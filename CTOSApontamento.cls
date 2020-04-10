VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOSApontamento"
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

'Public gobjAnotacao As ClassAnotacoes

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

'Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoRastroLote As AdmEvento
Attribute objEventoRastroLote.VB_VarHelpID = -1
Private WithEvents objEventoOP_OP As AdmEvento
Attribute objEventoOP_OP.VB_VarHelpID = -1
Private WithEvents objEventoProdutoOP_OP As AdmEvento
Attribute objEventoProdutoOP_OP.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoProdutoOP_Grid As AdmEvento
Attribute objEventoProdutoOP_Grid.VB_VarHelpID = -1

'mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const DATA1 As String = "Data"
Private Const ESTORNO1 As String = "Estorno"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const CCL1 As String = "Ccl"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const OP_CODIGO As String = "OP_Codigo"
Private Const PRODUTO_OP As String = "Produto_OP"
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
Private Const CONTACONTABILPROD1 As String = "ContaContabilProd"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"

Dim gcolcolRastreamentoSerie As Collection
Dim iTipoMovtoAnt As Integer

'Declaração das Variáveis Globais
Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFrameAtualItens As Integer
Dim iLinhaAntiga As Integer
Dim lCodigoAntigo As Long

Public gobjMovEst As ClassMovEstoque

Public objGrid As AdmGrid

Public iGrid_Sequencial_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_Descricao_Col As Integer
Public iGrid_OPCodigo_Col As Integer
Public iGrid_Benef_Col  As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Ccl_Col As Integer
Public iGrid_ProdutoOP_Col As Integer
Public iGrid_Estorno_Col As Integer
Public iGrid_ContaContabilEst_Col As Integer
Public iGrid_ContaContabilProducao_Col As Integer
Public iGrid_Lote_Col As Integer
Public iGrid_FilialOP_Col As Integer

Public objGridMO As AdmGrid
Public iGrid_MOCodigo_Col As Integer
Public iGrid_MONomeRed_Col As Integer
Public iGrid_MOTipo_Col As Integer
Public iGrid_MOHoras_Col As Integer
Public iGrid_MOOS_Col As Integer
Public iGrid_MOServico_Col As Integer

Public objGridMaq As AdmGrid
Public iGrid_Maquina_Col As Integer
Public iGrid_MaqQtd_Col As Integer
Public iGrid_MaqHoras_Col As Integer
Public iGrid_MaqOS_Col As Integer
Public iGrid_MaqServico_Col As Integer

Dim WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Dim WithEvents objEventoOPPadrao As AdmEvento
Attribute objEventoOPPadrao.VB_VarHelpID = -1
Dim WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Dim WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Dim WithEvents objEventoOP As AdmEvento
Attribute objEventoOP.VB_VarHelpID = -1
Dim WithEvents objEventoBotaoOP As AdmEvento
Attribute objEventoBotaoOP.VB_VarHelpID = -1
Dim WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Dim WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Dim WithEvents objEventoProdutoOP As AdmEvento
Attribute objEventoProdutoOP.VB_VarHelpID = -1
Dim WithEvents objEventoOP_OP_Todos As AdmEvento
Attribute objEventoOP_OP_Todos.VB_VarHelpID = -1
Dim WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1
Dim WithEvents objEventoMaq As AdmEvento
Attribute objEventoMaq.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 2

Const PGT_LIVRE = 0
Const PGT_QTD = 1
Const PGT_PROD = 2
Const PGT_DESCPRO = 3
Const PGT_MO_DESC = 4
Const PGT_TMO_DESC = 5
Const PGT_MAQ_NOME = 6
Const PGT_TMO_DESC2 = 7

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get CTBGerencial() As Object
     Set CTBGerencial = objUserControl.Controls("CTBGerencial")
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

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
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

Public Property Get BotaoProjetos() As Object
     Set BotaoProjetos = objUserControl.Controls("BotaoProjetos")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get BotaoSerie() As Object
     Set BotaoSerie = objUserControl.Controls("BotaoSerie")
End Property

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get BotaoLote() As Object
     Set BotaoLote = objUserControl.Controls("BotaoLote")
End Property

Public Property Get BotaoCcls() As Object
     Set BotaoCcls = objUserControl.Controls("BotaoCcls")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get BotaoEstoque() As Object
     Set BotaoEstoque = objUserControl.Controls("BotaoEstoque")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get BotaoGeraReq() As Object
     Set BotaoGeraReq = objUserControl.Controls("BotaoGeraReq")
End Property

Public Property Get OP() As Object
     Set OP = objUserControl.Controls("OP")
End Property

Public Property Get ProdutoOPGera() As Object
     Set ProdutoOPGera = objUserControl.Controls("ProdutoOPGera")
End Property

Public Property Get QuantidadeOP() As Object
     Set QuantidadeOP = objUserControl.Controls("QuantidadeOP")
End Property

Public Property Get LblUM() As Object
     Set LblUM = objUserControl.Controls("LblUM")
End Property

Public Property Get OPLabel() As Object
     Set OPLabel = objUserControl.Controls("OPLabel")
End Property

Public Property Get ProdutoOPLabel() As Object
     Set ProdutoOPLabel = objUserControl.Controls("ProdutoOPLabel")
End Property

Public Property Get Estorno() As Object
     Set Estorno = objUserControl.Controls("Estorno")
End Property

Public Property Get OPCodigoPadrao() As Object
     Set OPCodigoPadrao = objUserControl.Controls("OPCodigoPadrao")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get OPCodigo() As Object
     Set OPCodigo = objUserControl.Controls("OPCodigo")
End Property

Public Property Get BotaoOP() As Object
     Set BotaoOP = objUserControl.Controls("BotaoOP")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get FilialOP() As Object
     Set FilialOP = objUserControl.Controls("FilialOP")
End Property

Public Property Get Benef() As Object
     Set Benef = objUserControl.Controls("Benef")
End Property

Public Property Get Lote() As Object
     Set Lote = objUserControl.Controls("Lote")
End Property

Public Property Get ContaContabilProducao() As Object
     Set ContaContabilProducao = objUserControl.Controls("ContaContabilProducao")
End Property

Public Property Get ContaContabilEst() As Object
     Set ContaContabilEst = objUserControl.Controls("ContaContabilEst")
End Property

Public Property Get AlmoxPadrao() As Object
     Set AlmoxPadrao = objUserControl.Controls("AlmoxPadrao")
End Property

Public Property Get ProdutoOP() As Object
     Set ProdutoOP = objUserControl.Controls("ProdutoOP")
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

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get GridMovs() As Object
     Set GridMovs = objUserControl.Controls("GridMovs")
End Property

Public Property Get CclPadrao() As Object
     Set CclPadrao = objUserControl.Controls("CclPadrao")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get HORA() As Object
     Set HORA = objUserControl.Controls("Hora")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get LabelProjeto() As Object
     Set LabelProjeto = objUserControl.Controls("LabelProjeto")
End Property

Public Property Get QuantDisponivel() As Object
     Set QuantDisponivel = objUserControl.Controls("QuantDisponivel")
End Property

Public Property Get OPPadraoLabel() As Object
     Set OPPadraoLabel = objUserControl.Controls("OPPadraoLabel")
End Property

Public Property Get CodigoLabel() As Object
     Set CodigoLabel = objUserControl.Controls("CodigoLabel")
End Property

Public Property Get CclPadraoLabel() As Object
     Set CclPadraoLabel = objUserControl.Controls("CclPadraoLabel")
End Property

Public Property Get AlmoxPadraoLabel() As Object
     Set AlmoxPadraoLabel = objUserControl.Controls("AlmoxPadraoLabel")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
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

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get FrameItens() As Object
     Set FrameItens = objUserControl.Controls("FrameItens")
End Property

Public Property Get MOCodigo() As Object
     Set MOCodigo = objUserControl.Controls("MOCodigo")
End Property

Public Property Get MONomeRed() As Object
     Set MONomeRed = objUserControl.Controls("MONomeRed")
End Property

Public Property Get MOHoras() As Object
     Set MOHoras = objUserControl.Controls("MOHoras")
End Property

Public Property Get MOTipo() As Object
     Set MOTipo = objUserControl.Controls("MOTipo")
End Property

Public Property Get MOOS() As Object
     Set MOOS = objUserControl.Controls("MOOS")
End Property

Public Property Get MOServico() As Object
     Set MOServico = objUserControl.Controls("MOServico")
End Property

Public Property Get Maquina() As Object
     Set Maquina = objUserControl.Controls("Maquina")
End Property

Public Property Get MaqQtd() As Object
     Set MaqQtd = objUserControl.Controls("MaqQtd")
End Property

Public Property Get MaqHoras() As Object
     Set MaqHoras = objUserControl.Controls("MaqHoras")
End Property

Public Property Get MaqOS() As Object
     Set MaqOS = objUserControl.Controls("MaqOS")
End Property

Public Property Get MaqServico() As Object
     Set MaqServico = objUserControl.Controls("MaqServico")
End Property

Public Property Get GridMO() As Object
     Set GridMO = objUserControl.Controls("GridMO")
End Property

Public Property Get GridMaq() As Object
     Set GridMaq = objUserControl.Controls("GridMaq")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Benef_Click()

Dim lErro As Long

On Error GoTo Erro_Benef_Click

    iAlterado = REGISTRO_ALTERADO
    
    If GridMovs.Row <> 0 And objGrid.iLinhasExistentes <> 0 Then
    
        lErro = QuantDisponivel_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega as séries na coleção global
        If gcolcolRastreamentoSerie.Count >= GridMovs.Row Then
            lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    
    End If
    
    Exit Sub

Erro_Benef_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206198)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Benef_GotFocus()
'trata o evento gotfocus associado ao campo Benef

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Benef_KeyPress(KeyAscii As Integer)
'trata o evento keypress associado ao campo Benef

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Benef_Validate(Cancel As Boolean)
'trata o evento validate associado ao campo Benef

Dim lErro As Long

    Set objGrid.objControle = Benef
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub BotaoExcluir_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 206199

    Set objMovEstoque.objApontamento = New ClassOSAp
    objMovEstoque.objApontamento.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
        
    'Envia aviso perguntando se realmente deseja excluir o apontamento
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_APONTAMENTOSRV", objMovEstoque.objApontamento.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a producao
        lErro = CF("MovEstoque_ApSRV_Exclui", objMovEstoque, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Call Limpa_Tela_Apontamento

        Data.PromptInclude = False
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
        Data.PromptInclude = True
        
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206199
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206200)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("ApontamentoSRV_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    lCodigoAntigo = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206201)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovs.Row = 0 Then gError 206202
    
    If GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = "" Then gError 206203

    sConta = String(STRING_CONTA, 0)

    If GridMovs.Col = iGrid_ContaContabilEst_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ElseIf GridMovs.Col = iGrid_ContaContabilProducao_Col Then

        lErro = CF("Conta_Formata", ContaContabilProducao.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206202
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 206203
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206204)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If GridMovs.Col <> iGrid_ContaContabilEst_Col And GridMovs.Col <> iGrid_ContaContabilProducao_Col Then
        Me.Show
        Exit Sub
    End If
        
    If objPlanoConta.sConta <> "" Then
   
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 206205
        
        If GridMovs.Col = iGrid_ContaContabilEst_Col Then
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
        Else
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxuta
            ContaContabilProducao.PromptInclude = True
        End If

        GridMovs.TextMatrix(GridMovs.Row, GridMovs.Col) = objGrid.objControle.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 206205
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206206)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'se o codigo foi trocado
    If lCodigoAntigo <> StrParaLong(Trim(Codigo.Text)) Then

        'se o codigo novo está preenchido
        If Len(Trim(Codigo.ClipText)) > 0 Then
        
            Set objMovEstoque.objApontamento = New ClassOSAp
            objMovEstoque.objApontamento.lCodigo = Codigo.Text
            objMovEstoque.objApontamento.iFilialEmpresa = giFilialEmpresa
            
            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If

        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
        
   End If

   Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206209)

    End Select

    Exit Sub


End Sub

Public Sub ContaContabilEst_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ContaContabilEst_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub ContaContabilEst_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub ContaContabilEst_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilEst
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ContaContabilEst(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilEst

    Set objGrid.objControle = ContaContabilEst

    If Len(Trim(ContaContabilEst.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilEst.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 206210
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaMascarada
            ContaContabilEst.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEst.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError ERRO_SEM_MENSAGEM
    
            'conta não cadastrada
            If lErro = 5700 Then gError 206211
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206210
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 206211
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEst.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206212)

    End Select

    Exit Function

End Function

Public Sub ContaContabilProducao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ContaContabilProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub ContaContabilProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub ContaContabilProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabilProducao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ContaContabilProducao(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilProducao

    Set objGrid.objControle = ContaContabilProducao

    If Len(Trim(ContaContabilProducao.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilProducao.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 206213
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaMascarada
            ContaContabilProducao.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError ERRO_SEM_MENSAGEM
    
            'conta não cadastrada
            If lErro = 5700 Then gError 206214
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Saida_Celula_ContaContabilProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducao:

    Saida_Celula_ContaContabilProducao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206213
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 206214
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducao.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206215)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao As String
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFrameAtualItens = 1
    
    Set gcolcolRastreamentoSerie = New Collection

    'Inicializa todos os objeventos
    Set gobjMovEst = New ClassMovEstoque

    Set objEventoCodigo = New AdmEvento
    Set objEventoOPPadrao = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoOP = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoProdutoOP = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoProdutoOP = New AdmEvento
    Set objEventoProdutoOP_OP = New AdmEvento
    Set objEventoOP_OP = New AdmEvento
    Set objEventoOP_OP_Todos = New AdmEvento
    Set objEventoProdutoOP_Grid = New AdmEvento
    Set objEventoRastroLote = New AdmEvento
    Set objEventoMO = New AdmEvento
    Set objEventoMaq = New AdmEvento
    
'    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
'    Set gobjTelaProjetoInfo.objUserControl = Me
'    Set gobjTelaProjetoInfo.objTela = Me

    'Mostra a Data Atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Inicializa máscara de Produto , ProdutoOPGera , ProdutoOP
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoOPGera)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoOP)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MOServico)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MaqServico)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa Máscara para CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Ccl.Mask = sMascaraCclPadrao
    CclPadrao.Mask = sMascaraCclPadrao

    'Formata Quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    MOHoras.Format = FORMATO_ESTOQUE
    MaqHoras.Format = FORMATO_ESTOQUE

    'Inicializa GridMovs
    Set objGrid = New AdmGrid
    Set objGridMO = New AdmGrid
    Set objGridMaq = New AdmGrid

    lErro = Inicializa_GridMovs(objGrid)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_GridMO(objGridMO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_GridMaq(objGridMaq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_SERVICOS)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iLinhaAntiga = 0
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206216)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM 'Erro já tratado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206217)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridMovs(objGridInt As AdmGrid) As Long
'Inicializa o Grid

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridMovs

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Peça")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Escaninho")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("O.S.")
    objGridInt.colColuna.Add ("Serviço OS")
    objGridInt.colColuna.Add ("Lote/OP/Série Ini.")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta CTB Estoque")
    objGridInt.colColuna.Add ("Conta CTB Produção")
    objGridInt.colColuna.Add ("Estorno")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Benef.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (OPCodigo.Name)
    objGridInt.colCampo.Add (ProdutoOP.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ContaContabilEst.Name)
    objGridInt.colCampo.Add (ContaContabilProducao.Name)
    objGridInt.colCampo.Add (Estorno.Name)
    
    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Benef_Col = 4
    iGrid_Quantidade_Col = 5
    iGrid_Almoxarifado_Col = 6
    iGrid_OPCodigo_Col = 7
    iGrid_ProdutoOP_Col = 8
    iGrid_Lote_Col = 9
    iGrid_FilialOP_Col = 10
    iGrid_Ccl_Col = 11
    iGrid_ContaContabilEst_Col = 12
    iGrid_ContaContabilProducao_Col = 13
    iGrid_Estorno_Col = 14
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovs

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridMovs.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMovs = SUCESSO

    Exit Function

Erro_Inicializa_GridMovs:

    Inicializa_GridMovs = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206218)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim dtHora As Date

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela

    sTabela = "ApontamentoSRV"

    If Len(Trim(HORA.ClipText)) > 0 Then
        dtHora = CDate(HORA.Text)
    Else
        dtHora = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", StrParaDbl(Codigo.Text), 0, "Codigo"
    colCampoValor.Add "Data", StrParaDate(Data.Text), 0, "Data"
    colCampoValor.Add "Hora", CDbl(dtHora), 0, "Hora"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206219)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do Banco de Dados
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    Set objMovEstoque.objApontamento = New ClassOSAp
    objMovEstoque.objApontamento.lCodigo = colCampoValor.Item("Codigo").vValor
    objMovEstoque.objApontamento.dtData = colCampoValor.Item("Data").vValor
    objMovEstoque.objApontamento.dtHora = colCampoValor.Item("Hora").vValor
    objMovEstoque.objApontamento.iFilialEmpresa = giFilialEmpresa

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206220)

    End Select

    Exit Sub

End Sub

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long
'Preenche a Tela

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_Tela

    lErro = CF("MovimentoEstoqueAp_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    'Limpa a tela sem Fechar o Comando de setas
    'Função genérica para Limpar a Tela
    Call Limpa_Tela_Apontamento

    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.objApontamento.lCodigo)
    Codigo.PromptInclude = True

    If lErro = SUCESSO Then
        
        For iIndice = 1 To objGrid.iLinhasExistentes 'm
            GridMovs.TextMatrix(iIndice, 0) = iIndice
        Next
    
        'Limpa o Grid
        Call Grid_Limpa(objGrid)
    
        'Se o grid permite excluir e incluir Linhas
        If objGrid.iProibidoIncluir <> GRID_PROIBIDO_INCLUIR And objGrid.iProibidoExcluir <> GRID_PROIBIDO_EXCLUIR Then
            'prepara o Grid para não permitir inserir e excluir Linhas
            objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
            objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
            Call Grid_Inicializa(objGrid)
        End If
        
        Set gobjMovEst = objMovEstoque

        If objMovEstoque.dtData <> DATA_NULA Then
            Data.PromptInclude = False
            Data.Text = Format(objMovEstoque.dtData, "dd/mm/yy")
            Data.PromptInclude = True
        Else
            Data.PromptInclude = False
            Data.Text = ""
            Data.PromptInclude = True
        End If
    
        HORA.PromptInclude = False
        'este teste está correto
        If objMovEstoque.dtData <> DATA_NULA Then HORA.Text = Format(objMovEstoque.dtHora, "hh:mm:ss")
        HORA.PromptInclude = True
    
        'Preenche o Grid
        lErro = Preenche_GridMovs(objMovEstoque.colItens)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Preenche_GridAp(objMovEstoque.objApontamento)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.objApontamento.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError ERRO_SEM_MENSAGEM

    '    '###################################################
    '    'Inserido por Wagner 04/08/2006
    '    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objMovEstoque.objProjetoInfo, objMovEstoque.lCodigo, PRJ_CR_TIPO_REQPROD, "", objMovEstoque.iFilialEmpresa)
    '    If lErro <> SUCESSO Then gError 185850
    '    '###################################################
    
    End If

    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206221
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_REQPRODUCAO", gErr, objMovEstoque.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206222)

    End Select

    Exit Function

End Function

Public Sub Preenche_Grid(ByVal iLinha As Integer, ByVal objGrid As AdmGrid, ParamArray avParam() As Variant)

Dim lErro As Long
Dim iIndice As Integer
Dim vValor As Variant
Dim vValorFormatado As Variant, sProdForm As String
Dim iColuna As Integer
Dim iTipo As Integer
Dim objMO As ClassMO
Dim objTMO As ClassTiposDeMaodeObra
Dim objMaq As ClassMaquinas

On Error GoTo Erro_Preenche_Grid

    For iIndice = 0 To UBound(avParam)
    
        If iIndice Mod 2 = 1 Then
            iColuna = (iIndice + 1) / 2
            vValor = avParam(iIndice)
            
            Select Case iTipo
            
                Case PGT_LIVRE
                    vValorFormatado = vValor
            
                Case PGT_QTD
                    vValorFormatado = Formata_Estoque(CDbl(vValor))
                    
                Case PGT_PROD
                
                    sProdForm = String(STRING_PRODUTO, 0)

                    'coloca o Produto no formato que é exibido na tela
                    lErro = Mascara_RetornaProdutoTela(CStr(vValor), sProdForm)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    vValorFormatado = sProdForm
                    
                Case PGT_MO_DESC
                
                    Set objMO = New ClassMO
                
                    objMO.lCodigo = vValor
                    
                    'Lê os dados da mao de obra a ser excluido
                    lErro = CF("MO_Le", objMO)
                    If lErro <> SUCESSO And lErro <> 193817 Then gError ERRO_SEM_MENSAGEM
                    
                    vValorFormatado = objMO.sNomeReduzido
                    
                Case PGT_TMO_DESC
                
                    Set objTMO = New ClassTiposDeMaodeObra
                
                    objTMO.iCodigo = vValor
                    
                    'Lê os dados da mao de obra a ser excluido
                    lErro = CF("TiposDeMaodeObra_Le", objTMO)
                    If lErro <> SUCESSO And lErro <> 193817 Then gError ERRO_SEM_MENSAGEM
                
                    vValorFormatado = objTMO.sDescricao
                    
                Case PGT_TMO_DESC2
                
                    Set objTMO = New ClassTiposDeMaodeObra
                    Set objMO = New ClassMO
                
                    objMO.lCodigo = vValor
                    
                    'Lê os dados da mao de obra a ser excluido
                    lErro = CF("MO_Le", objMO)
                    If lErro <> SUCESSO And lErro <> 193817 Then gError ERRO_SEM_MENSAGEM
                
                    objTMO.iCodigo = objMO.iTipo
                    
                    'Lê os dados da mao de obra a ser excluido
                    lErro = CF("TiposDeMaodeObra_Le", objTMO)
                    If lErro <> SUCESSO And lErro <> 193817 Then gError ERRO_SEM_MENSAGEM
                
                    vValorFormatado = objTMO.sDescricao
                    
                Case PGT_MAQ_NOME
                
                    Set objMaq = New ClassMaquinas
                    
                    objMaq.lNumIntDoc = vValor
                
                    'Le a Máquina atraves do NumIntDoc
                    lErro = CF("Maquinas_Le_NumIntDoc", objMaq)
                    If lErro <> SUCESSO And lErro <> 106353 Then gError ERRO_SEM_MENSAGEM
               
                    vValorFormatado = objMaq.sNomeReduzido
            
            End Select
            
            objGrid.objGrid.TextMatrix(iLinha, iColuna) = CStr(vValorFormatado)
        Else
            iTipo = avParam(iIndice)
        End If

    Next

    Exit Sub

Erro_Preenche_Grid:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206223)

    End Select

    Exit Sub

End Sub

Private Function Preenche_GridAp(ByVal objOSAp As ClassOSAp) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOSApMO As ClassOSApMO
Dim objOSApMaq As ClassOSApMaq

On Error GoTo Erro_Preenche_GridAp

    iIndice = 0
    For Each objOSApMO In objOSAp.colMO
        iIndice = iIndice + 1
        Call Preenche_Grid(iIndice, objGridMO, PGT_LIVRE, objOSApMO.lCodigoMO, PGT_MO_DESC, objOSApMO.lCodigoMO, PGT_TMO_DESC, objOSApMO.iTipo, PGT_QTD, objOSApMO.dHorasGastas, PGT_LIVRE, objOSApMO.sOS, PGT_PROD, objOSApMO.sProdutoOS)
    Next
    objGridMO.iLinhasExistentes = objOSAp.colMO.Count

    iIndice = 0
    For Each objOSApMaq In objOSAp.colMaq
        iIndice = iIndice + 1
        Call Preenche_Grid(iIndice, objGridMaq, PGT_MAQ_NOME, objOSApMaq.lNumIntDocMaq, PGT_LIVRE, objOSApMaq.iQuantidade, PGT_QTD, objOSApMaq.dHorasGastas, PGT_LIVRE, objOSApMaq.sOS, PGT_PROD, objOSApMaq.sProdutoOS)
    Next
    objGridMaq.iLinhasExistentes = objOSAp.colMaq.Count

    Preenche_GridAp = SUCESSO

    Exit Function

Erro_Preenche_GridAp:

    Preenche_GridAp = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206224)

    End Select

    Exit Function
    
End Function

Private Function Preenche_GridMovs(colItens As colItensMovEstoque) As Long
'Preeche o grid com os dados passados em colItens

Dim iIndice As Integer
Dim sProdutoMascarado As String, sProdutoOPMascarado As String
Dim sCclMascarado As String
Dim lErro As Long
Dim objTipoMovEst As ClassTipoMovEst
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sContaEnxutaEst As String
Dim sContaEnxutaProducao As String
Dim colRatreamentoMovto As New Collection
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim objFilialOP As New AdmFiliais
Dim colRastreamentoSerie As Collection
Dim iBenef As Integer

On Error GoTo Erro_Preenche_GridMovs

    Set gcolcolRastreamentoSerie = New Collection

    'Preenche gridmovs
    For Each objItemMovEstoque In colItens

        iIndice = iIndice + 1

        'Mascara o Produto
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 206225

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True
        
        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilEst <> "" Then
        
            sContaEnxutaEst = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True
            
            GridMovs.TextMatrix(iIndice, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then
        
            sContaEnxutaProducao = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxutaProducao
            ContaContabilProducao.PromptInclude = True
            
            GridMovs.TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) = ContaContabilProducao.Text
            
        End If
        
        GridMovs.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridMovs.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemMovEstoque.sProdutoDesc
        GridMovs.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemMovEstoque.sSiglaUM
        GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
        GridMovs.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed
        GridMovs.TextMatrix(iIndice, iGrid_OPCodigo_Col) = objItemMovEstoque.sOPCodigo

        If objItemMovEstoque.iPossuiGrade = MARCADO Then GridMovs.TextMatrix(iIndice, 0) = "# " & GridMovs.TextMatrix(iIndice, 0)
        
        iBenef = BENEF_COMBO_DISP
        
        Call Combo_Seleciona_ItemData(Benef, iBenef)
        
        GridMovs.TextMatrix(iIndice, iGrid_Benef_Col) = Benef.Text

        'Mascara o ProdutoOP
        sProdutoOPMascarado = String(STRING_PRODUTO, 0)

        If Len(Trim(objItemMovEstoque.sProdutoOP)) > 0 Then

            lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProdutoOP, sProdutoOPMascarado)
            If lErro <> SUCESSO Then gError 206226
    
            ProdutoOP.PromptInclude = False
            ProdutoOP.Text = sProdutoOPMascarado
            ProdutoOP.PromptInclude = True
            
        Else
    
            sProdutoOPMascarado = ""
            
        End If
    
        GridMovs.TextMatrix(iIndice, iGrid_ProdutoOP_Col) = sProdutoOPMascarado

        'Mascara o Ccl
        If objItemMovEstoque.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemMovEstoque.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Else

            sCclMascarado = ""

        End If

        GridMovs.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        Set colRatreamentoMovto = New Collection
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError ERRO_SEM_MENSAGEM
        
        'Se existe rastreamento
        If colRatreamentoMovto.Count > 0 Then
                        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto(1)
            
            gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie
            
            If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovs.TextMatrix(iIndice, iGrid_Lote_Col) = objRatreamentoMovto.sLote
            
            If objRatreamentoMovto.iFilialOP > 0 Then
            
                objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP

                'Le a Filial Empresa da OP para pegar a descrição
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                GridMovs.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
            
            End If
        
        Else
            Set colRastreamentoSerie = New Collection
            gcolcolRastreamentoSerie.Add colRastreamentoSerie
        
        End If
        
        If objItemMovEstoque.lNumIntDocEst = 0 Then
            GridMovs.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"
        Else
            GridMovs.TextMatrix(iIndice, iGrid_Estorno_Col) = "1"
        End If

    Next

    objGrid.iLinhasExistentes = colItens.Count
    
    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Preenche_GridMovs = SUCESSO

    Exit Function

Erro_Preenche_GridMovs:

    Preenche_GridMovs = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206225
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objItemMovEstoque.sProduto)

        Case 206226
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objItemMovEstoque.sProdutoOP)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206227)

    End Select

    Exit Function

End Function

Public Sub CodigoLabel_Click()

Dim objAp As New ClassOSAp
Dim colSelecao As New Collection

    If Len(Trim(Codigo.ClipText)) <> 0 Then objAp.lCodigo = CLng(Codigo.Text)

    'mario completar a colecao
    Call Chama_Tela("ApontamentoSRVLista", colSelecao, objAp, objEventoCodigo)
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objAp As ClassOSAp
Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_objCodigoEvento_evSelecao

    Set objAp = obj1
    
    objMovEstoque.lCodigo = objAp.lCodigoMovEst
    objMovEstoque.iFilialEmpresa = objAp.iFilialEmpresa
    Set objMovEstoque.objApontamento = objAp

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objCodigoEvento_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206228)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOP_OP_evSelecao(obj1 As Object)

Dim objItemOP As ClassOS

    Set objItemOP = obj1

    'Preeche OP com o codigo de objOrdemProducao
    OP.Text = objItemOP.sCodigo

    Me.Show

End Sub

Private Sub objEventoOP_OP_Todos_evSelecao(obj1 As Object)

Dim objOP As ClassOS

    Set objOP = obj1

    'Preeche OP com o codigo de objOrdemProducao
    OP.Text = objOP.sCodigo

    Me.Show

End Sub

Private Sub objEventoProdutoOP_Grid_evSelecao(obj1 As Object)

Dim objItemOS As ClassOS
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoProdutoOP_Grid_evSelecao

    Set objItemOS = obj1

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objItemOS.sServico, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 206229

    If iFrameAtualItens = 1 Then
        If Not (Me.ActiveControl Is ProdutoOP) Then
                GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = sProdutoMascarado
        End If
        ProdutoOP.PromptInclude = False
        ProdutoOP.Text = sProdutoMascarado
        ProdutoOP.PromptInclude = True
    ElseIf iFrameAtualItens = 2 Then
        If Not (Me.ActiveControl Is MOOS) Then
                GridMO.TextMatrix(GridMO.Row, iGrid_MOServico_Col) = sProdutoMascarado
        End If
        MOServico.PromptInclude = False
        MOServico.Text = sProdutoMascarado
        MOServico.PromptInclude = True
    Else
        If Not (Me.ActiveControl Is MaqOS) Then
                GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqServico_Col) = sProdutoMascarado
        End If
        MaqServico.PromptInclude = False
        MaqServico.Text = sProdutoMascarado
        MaqServico.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_Grid_evSelecao:

    Select Case gErr

        Case 206229
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objItemOS.sServico)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206230)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProdutoOP_OP_evSelecao(obj1 As Object)

Dim objItemOP As ClassOS
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoOP_OP_evSelecao

    Set objItemOP = obj1

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objItemOP.sServico, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 206231

    ProdutoOPGera.PromptInclude = False
    ProdutoOPGera.Text = sProdutoMascarado
    ProdutoOPGera.PromptInclude = True
    
    Call ProdutoOPGera_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_OP_evSelecao:

    Select Case gErr
        
        Case 206231
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objItemOP.sServico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206232)

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
        
            Case TAB_Movimentos
                Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_PRODUCAO_MOVIMENTOS
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_PRODUCAO_CONTABILIZACAO
                        
        End Select
    
    End If

End Sub

Public Sub OPPadraoLabel_Click()

Dim objOrdemProducao As New ClassOS
Dim colSelecao As New Collection

    Call Chama_Tela("OSAbertaLista", colSelecao, objOrdemProducao, objEventoOPPadrao)

End Sub

Private Sub objEventoOPPadrao_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOS

    Set objOrdemProducao = obj1

    'Preenche OPCodigo Padrão
    OPCodigoPadrao.Text = objOrdemProducao.sCodigo

    Me.Show

End Sub

Public Sub CclPadraoLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)

Dim objCcl As ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 206233

    'Preeche o CclPadrao com o codigo de objCcl
    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 206233
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206234)

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

Public Sub OPLabel_Click()

Dim objOrdemProducao As New ClassOS
Dim colSelecao As New Collection
Dim objItemOP As New ClassItemOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_OPLabel_Click
    
    lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se o Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add Trim(OP.Text)
                                
        Call Chama_Tela("ItemOSLista", colSelecao, objOrdemProducao, objEventoOP_OP, "Codigo = ?")
        
    Else
    
        If Len(Trim(OP.Text)) <> 0 Then objOrdemProducao.sCodigo = Trim(OP.Text)

        Call Chama_Tela("OSAbertaLista", colSelecao, objOrdemProducao, objEventoOP_OP_Todos)
    
    End If
    
    Exit Sub
    
Erro_OPLabel_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206235)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOS
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
    
On Error GoTo Erro_objEventoOP_evSelecao

    Set objOrdemProducao = obj1
    
    If iFrameAtualItens = 1 Then
        'Preenche a OP
        GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = objOrdemProducao.sCodigo
        OPCodigo.Text = objOrdemProducao.sCodigo
    ElseIf iFrameAtualItens = 2 Then
        'Preenche a OP
        GridMO.TextMatrix(GridMO.Row, iGrid_MOOS_Col) = objOrdemProducao.sCodigo
        MOOS.Text = objOrdemProducao.sCodigo
    Else
        'Preenche a OP
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqOS_Col) = objOrdemProducao.sCodigo
        MaqOS.Text = objOrdemProducao.sCodigo
    End If

    Me.Show

    Exit Sub
    
Erro_objEventoOP_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206236)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub ProdutoOPLabel_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim objItemOP As New ClassOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoOPLabel_Click
    
    If Len(Trim(OP.Text)) = 0 Then gError 206237
    
    lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objItemOP.sServico = sProdutoFormatado
        
    colSelecao.Add Trim(OP.Text)
        
    Call Chama_Tela("ItemOSLista", colSelecao, objItemOP, objEventoProdutoOP_OP, "Codigo = ?")
    
    Exit Sub
    
Erro_ProdutoOPLabel_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206237
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREECHIDA", gErr)
            OP.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206238)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProdutoOP_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoOP_evSelecao

    Set objProduto = obj1

    lErro = CF("Produto_Formata", objProduto.sCodigo, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 206239

        ProdutoOPGera.PromptInclude = False
        ProdutoOPGera.Text = sProdutoMascarado
        ProdutoOPGera.PromptInclude = True
        
        'Preenche o label de UM
        LblUM.Caption = objProduto.sSiglaUMEstoque
            
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206239
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206240)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoOP_Click(Index As Integer)

Dim objOrdemProducao As New ClassOS
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoOP_Click
    
    If (GridMovs.Row = 0) Then gError 206241

    If iFrameAtualItens = 1 Then
        sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)
    
        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Se na Linha corrente Produto estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 206242
        
    ElseIf iFrameAtualItens = 2 Then
        If GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col) = 0 Then gError 206243
    Else
        If GridMaq.TextMatrix(GridMaq.Row, iGrid_Maquina_Col) = 0 Then gError 206244
    End If

    Call Chama_Tela("OSAbertaLista", colSelecao, objOrdemProducao, objEventoOP)
   
    Exit Sub

Erro_BotaoOP_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206241
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 206242
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 206243
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOMO_NAO_PREENCHIDO_GRID2", gErr, GridMO.Row)
        
        Case 206244
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_PREENCHIDA_GRID", gErr, GridMaq.Row)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206245)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOS
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOS
Dim lErro As Long

On Error GoTo Erro_objEventoBotaoOP_evSelecao

    Set objOrdemProducao = obj1
    
    If iFrameAtualItens = 1 Then
        If (GridMovs.Row) > 0 Then
    
            lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Se Produto desta Linha Estiver Preenchido
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = objOrdemProducao.sCodigo
            End If
    
        End If
    ElseIf iFrameAtualItens = 3 Then
        If GridMO.Row > 0 Then
            If GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col) <> 0 Then
                GridMO.TextMatrix(GridMO.Row, iGrid_MOOS_Col) = objOrdemProducao.sCodigo
            End If
        End If
    Else
        If GridMaq.Row > 0 Then
            If GridMaq.TextMatrix(GridMaq.Row, iGrid_Maquina_Col) <> 0 Then
                GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqOS_Col) = objOrdemProducao.sCodigo
            End If
        End If
    End If

    Exit Sub

Erro_objEventoBotaoOP_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206246)

    End Select

    Exit Sub

End Sub

Public Sub BotaoCcls_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo BotaoCcls_Click
    
    If (GridMovs.Row = 0) Then gError 206247

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 206248

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
   
    Exit Sub

BotaoCcls_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206247
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 206248
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206249)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    If GridMovs.Row > 0 Then

        'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))) <> 0 Then

            'Mascara para colocar na tela

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            'Coloca o valor do Ccl na coluna correspondente
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Ccl_Col) = sCclMascarado

            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206250)

    End Select

    Exit Sub

End Sub

Public Sub BotaoServicos_Click(Index As Integer)

Dim colSelecao As New Collection
Dim lErro As Long
Dim objItemOP As New ClassOS
Dim sOS As String
    
On Error GoTo Erro_BotaoServicos_Click

    If iFrameAtualItens = 1 Then
        If GridMovs.Row = 0 Then gError 206251
        sOS = GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col)
    ElseIf iFrameAtualItens = 2 Then
        If GridMO.Row = 0 Then gError 206252
        sOS = GridMO.TextMatrix(GridMO.Row, iGrid_MOOS_Col)
    Else
        If GridMaq.Row = 0 Then gError 206253
        sOS = GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqOS_Col)
    End If
        
    'se a OP estiver preenchida, mostra só os produtos da OP em questão
    If Len(Trim(sOS)) = 0 Then gError 206254
        
    colSelecao.Add Trim(sOS)
    
    Call Chama_Tela("ItemOSLista", colSelecao, objItemOP, objEventoProdutoOP_Grid, "Codigo = ?")
        
    Exit Sub
   
Erro_BotaoServicos_Click:

    Select Case gErr
              
        Case 206251 To 206253
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 206254
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO_OP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206255)
        
    End Select
    
    Exit Sub
   
End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemMovEst As ClassItemMovEstoque) As Long

Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sContaEnxuta As String
Dim sAlmoxarifadoPadrao As String
Dim colRastreamentoSerie As New Collection

On Error GoTo Erro_ProdutoLinha_Preenche

    If Trim(Len(objProduto.sContaContabilProducao)) > 0 Then
        
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
        If lErro <> SUCESSO Then gError 206256
        
        'preenche  a ContaContabilProducao
        ContaContabilProducao.PromptInclude = False
        ContaContabilProducao.Text = sContaEnxuta
        ContaContabilProducao.PromptInclude = True
        
        'preenche Conta De Producao
        GridMovs.TextMatrix(GridMovs.Row, iGrid_ContaContabilProducao_Col) = ContaContabilProducao.Text
            
    End If

    Call Combo_Seleciona_ItemData(Benef, BENEF_COMBO_DISP)
    GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col) = Benef.Text
    
    'Unidade de Medida
    GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovs.TextMatrix(GridMovs.Row, iGrid_Descricao_Col) = objProduto.sDescricao

    If Len(Trim(objProduto.sGrade)) = 0 Then

        'Almoxarifado
        '(Utiliza Almoxarifado Padrão caso esteja preenchido)
        If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then ' And Len(Trim(Almoxarifado.ClipText)) = 0 Then
            lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadrao)
            If lErro = SUCESSO Then
                GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
            Else
                GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = ""
            End If
            
        Else
            'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
            lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
        End If

    End If

    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovs.TextMatrix(GridMovs.Row, iGrid_Ccl_Col) = CclPadrao.Text

    'OPCODIGO
    If Len(Trim(objProduto.sGrade)) = 0 And Len(Trim(OPCodigoPadrao.Text)) <> 0 Then GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = OPCodigoPadrao.Text

    'Preenche Estorno com Valor 0 (Checked = False)
    GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) = "0"

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovs.Row - GridMovs.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolcolRastreamentoSerie.Add colRastreamentoSerie
        
        gobjMovEst.colItens.Add1 objItemMovEst
        
'        Call gobjTelaProjetoInfo.Adiciona_Linha(GridMovs.Row)
    
        If Len(Trim(objProduto.sGrade)) = 0 Then

            objItemMovEst.iPossuiGrade = DESMARCADO
            
        Else
        
            objItemMovEst.iPossuiGrade = MARCADO
            GridMovs.TextMatrix(GridMovs.Row, 0) = "# " & GridMovs.TextMatrix(GridMovs.Row, 0)
        
        End If
        
        objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
        objItemMovEst.sProduto = objProduto.sCodigo
    
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206256
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objProduto.sContaContabilProducao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206257)

    End Select

    Exit Function

End Function

Public Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If (GridMovs.Row = 0) Then gError 206258

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 206259

    colSelecao.Add sProdutoFormatado

    Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206258
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 206259
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206260)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    If (GridMovs.Row > 0) Then

        sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
            lErro = Preenche_ContaContabilEst(GridMovs.Row, GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then
                'Calcula a Quantidade Disponível nesse Almoxarifado
                 lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
                If lErro <> SUCESSO Then gError 206261
            Else
                lErro = QuantLote_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206261
            'Limpa a Quantidade Disponível da Tela e o Almoxarifado
            QuantDisponivel.Caption = ""
            Almoxarifado.Text = ""
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206262)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal obj1 As Object) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objMovEstoque As ClassMovEstoque
Dim objAp As ClassOSAp

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not obj1 Is Nothing Then
    
        If TypeName(obj1) = "ClassMovEstoque" Then
            Set objMovEstoque = obj1
            objMovEstoque.objApontamento.lCodigoMovEst = objMovEstoque.lCodigo
        Else
            Set objAp = obj1
            Set objMovEstoque = New ClassMovEstoque
            Set objMovEstoque.objApontamento = objAp
            objMovEstoque.lCodigo = objAp.lCodigoMovEst
        End If
        objMovEstoque.iFilialEmpresa = giFilialEmpresa
            
        lErro = Preenche_Tela(objMovEstoque)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206263
            Call Rotina_Erro(vbOKOnly, "ERRO_MOV_EST_NAO_APSRV", gErr, objMovEstoque.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206264)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub OP_Validate(bMantemFoco As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOS
Dim objItemOP As New ClassItemOS
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_OP_Validate

    'Verifica se foi Preenchido
    If Len(Trim(OP.Text)) > 0 Then

        objOrdemProducao.sCodigo = OP.Text
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa

        lErro = CF("OS_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 195102 Then gError 34748

        If lErro = 195102 Then gError 206265
        
        lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34750

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OP.Text
            objItemOP.sServico = sProdutoFormatado

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 34751

            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 206266

        End If

    End If

    Exit Sub

Erro_OP_Validate:

    bMantemFoco = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206265
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OSCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then
                Call Chama_Tela("OrdemProducao", objOrdemProducao)
            End If

        Case 206266
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206267)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoOPGera_Validate(bMantemFoco As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemOP As New ClassItemOS
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_ProdutoOPGera_Validate

    If Len(Trim(ProdutoOPGera.ClipText)) > 0 Then

        sProduto = ProdutoOPGera.Text

        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 25041 Then gError 206268

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
           If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 206269
        End If
        
        LblUM.Caption = objProduto.sSiglaUMEstoque
        
        If Len(Trim(OP.Text)) > 0 Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OP.Text
            objItemOP.sServico = objProduto.sCodigo

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError ERRO_SEM_MENSAGEM
            If lErro = 34711 Then gError 206270

        End If

    End If

    Exit Sub

Erro_ProdutoOPGera_Validate:

    bMantemFoco = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206268
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", sProduto)

             If vbMsg = vbYes Then
                objProduto.sCodigo = sProduto
                Call Chama_Tela("Produto", objProduto)
            End If
    
        Case 206269
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr)

        Case 206270
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206271)

    End Select

    Exit Sub

End Sub

Public Sub QuantidadeOP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_QuantidadeOP_Validate

    If Len(Trim(QuantidadeOP.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeOP.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_QuantidadeOP_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206272)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGeraReq_Click()

Dim lErro As Long
Dim dQuantidade As Double
Dim objOP As New ClassOS
Dim objMovEstoque As New ClassMovEstoque
Dim objItemOP As New ClassItemOS
Dim vbMsg As VbMsgBoxResult
Dim iProdutoOPPreenchido As Integer
Dim sProdutoOPFormatado As String
Dim iIndice As Integer

On Error GoTo Erro_BotaoGeraReq_Click
          
    'se já existem linhas preenchidas no grid
    If objGrid.iLinhasExistentes > 0 Then

        DoEvents
        
        'pergunta se quer apagá-las
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_APAGAR_GRID")
        
        'se a resposta for afirmativa ==> remove as linhas do grid
        If vbMsg = vbYes Then
        
'            For iIndice = objGrid.iLinhasExistentes To 1 Step -1
'                Call gobjTelaProjetoInfo.Remove_Linha(iIndice)
'            Next
        
            Call Grid_Limpa(objGrid)
            Call Grid_Limpa(objGridMO)
            Call Grid_Limpa(objGridMaq)

            'Remove os ítens de gobjMovEst
            Set gobjMovEst.colItens = New colItensMovEstoque
            Set gobjMovEst.objApontamento = New ClassOSAp

        End If
        
    End If

    'se a ordem de produção não foi preenchida ==> erro
    If Len(Trim(OP.Text)) = 0 Then gError 206273

    'se a quantidade foi preenchida ==> guarda a quantidade
    If Len(Trim(QuantidadeOP.ClipText)) > 0 Then
        dQuantidade = CDbl(QuantidadeOP.Text)
    End If

    'se a quantidade for maior que 0 e o produto não foi preenchido ==> erro
    If (dQuantidade > 0) And (Len(Trim(ProdutoOPGera.ClipText)) = 0) Then gError 206274

    'se o produto estiver preenchido
    If Len(Trim(ProdutoOPGera.ClipText)) > 0 Then

        'coloca o produto no formato do BD
        lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoOPFormatado, iProdutoOPPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'se o produto não foi preenchido
    If iProdutoOPPreenchido <> PRODUTO_PREENCHIDO Then

        objOP.sCodigo = OP.Text
        objOP.iFilialEmpresa = giFilialEmpresa

        'trata de gerar as requisições para todos os itens da OP
        lErro = OP_Requisicoes_Gera(objOP, objMovEstoque)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else

        'se o produto foi preenchido
        objItemOP.sCodigo = OP.Text
        objItemOP.iFilialEmpresa = giFilialEmpresa
        objItemOP.sServico = sProdutoOPFormatado
        If (dQuantidade > 0) Then objItemOP.dQuantidade = dQuantidade

        'gera as requisições para o produto em questão
        lErro = ItemOP_Requisicao_Gera(objItemOP, objMovEstoque)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'preenche o grid de itens requisitados com os itens dos kits que compõem a ordem de produção
    lErro = BotaoGeraReq_Preenche_Grid(objMovEstoque.colItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Preenche_GridAp(objMovEstoque.objApontamento)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_BotaoGeraReq_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206273
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREECHIDA", gErr)

        Case 206274
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SEM_PRODUTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206275)

    End Select

    Exit Sub

End Sub

Private Function BotaoGeraReq_Preenche_Grid(colItensMovEstoque As colItensMovEstoque) As Long
'preenche o grid de itens requisitados com os itens dos kits que compoe a ordem de produção

Dim sAlmoxarifadoPadrao As String
Dim sContaEnxuta As String
Dim lErro As Long
Dim sProdutoMascarado As String
Dim sProdutoOPMascarado As String
Dim iLinha As Integer
Dim sCclFormata As String
Dim iIndice As Integer, iCclPreenchida As Integer
Dim objItemMovEst As ClassItemMovEstoque
Dim colRastreamentoSerie As Collection

On Error GoTo Erro_BotaoGeraReq_Preenche_Grid

    If colItensMovEstoque.Count > 0 Then

        For iIndice = 1 To colItensMovEstoque.Count

            sProdutoMascarado = String(STRING_PRODUTO, 0)
            sProdutoOPMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoTela(colItensMovEstoque.Item(iIndice).sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 206276

            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True

            lErro = Valida_Repeticao(Produto.Text, OP.Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            iLinha = objGrid.iLinhasExistentes + 1

            GridMovs.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
            GridMovs.TextMatrix(iLinha, iGrid_Descricao_Col) = colItensMovEstoque.Item(iIndice).sProdutoDesc
            GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = colItensMovEstoque.Item(iIndice).sSiglaUM
            GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(colItensMovEstoque.Item(iIndice).dQuantidade)
            GridMovs.TextMatrix(iLinha, iGrid_OPCodigo_Col) = OP.Text
            GridMovs.TextMatrix(iLinha, iGrid_Estorno_Col) = colItensMovEstoque.Item(iIndice).lNumIntDocEst

            objGrid.iLinhasExistentes = iLinha
            
'            Call gobjTelaProjetoInfo.Adiciona_Linha(iLinha)

            If Trim(Len(colItensMovEstoque.Item(iIndice).sContaContabilAplic)) > 0 Then
        
                lErro = Mascara_RetornaContaEnxuta(colItensMovEstoque.Item(iIndice).sContaContabilAplic, sContaEnxuta)
                If lErro <> SUCESSO Then gError 206278
                
                'preenche  a ContaContabilProducao
                ContaContabilProducao.PromptInclude = False
                ContaContabilProducao.Text = sContaEnxuta
                ContaContabilProducao.PromptInclude = True
                
                'preenche Conta De Producao
                GridMovs.TextMatrix(iLinha, iGrid_ContaContabilProducao_Col) = ContaContabilProducao.Text
                    
            End If
            
            Call Combo_Seleciona_ItemData(Benef, colItensMovEstoque.Item(iIndice).iBenef)
            GridMovs.TextMatrix(iLinha, iGrid_Benef_Col) = Benef.Text

            lErro = Mascara_RetornaProdutoTela(colItensMovEstoque.Item(iIndice).sProdutoOP, sProdutoOPMascarado)
            If lErro <> SUCESSO Then gError 206277

            GridMovs.TextMatrix(iLinha, iGrid_ProdutoOP_Col) = sProdutoOPMascarado

            If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then
                GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
            Else

                'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
                lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", colItensMovEstoque.Item(iIndice).sProduto, sAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 52219 And lErro <> 52221 Then gError ERRO_SEM_MENSAGEM

                If lErro = SUCESSO Then GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
            
            End If

            'coloca o Ccl no formato do bd
            lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            If iCclPreenchida = CCL_PREENCHIDA Then GridMovs.TextMatrix(iLinha, iGrid_Ccl_Col) = CclPadrao.Text

            Set objItemMovEst = New ClassItemMovEstoque
            gobjMovEst.colItens.Add1 objItemMovEst

            'preenche a conta contabil de estoque, se houver alguma associada ao produto/almoxarifado em questão
            lErro = Preenche_ContaContabilEst(iLinha, GridMovs.TextMatrix(iLinha, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set colRastreamentoSerie = New Collection
            
            gcolcolRastreamentoSerie.Add colRastreamentoSerie

        Next

        lErro = Grid_Refresh_Checkbox(objGrid)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    BotaoGeraReq_Preenche_Grid = SUCESSO

    Exit Function

Erro_BotaoGeraReq_Preenche_Grid:

    BotaoGeraReq_Preenche_Grid = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206276
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, colItensMovEstoque.Item(iIndice).sProduto)
        
        Case 206277
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOTELA", gErr, colItensMovEstoque.Item(iIndice).sProdutoOP)
        
        Case 206278
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, colItensMovEstoque.Item(iIndice).sContaContabilAplic)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206279)
        
    End Select
    
    Exit Function
    
End Function

Private Function OP_Requisicoes_Gera(ByVal objOP As ClassOS, ByVal objMovEstoque As ClassMovEstoque) As Long
'Gera Requisições para todos os Produtos da OP a partir das quantidades que faltam produzir

Dim lErro As Long
Dim objItemOP As ClassItemOS
Dim objOSAp As New ClassOSAp

On Error GoTo Erro_OP_Requisicoes_Gera

    Set objMovEstoque.objApontamento = objOSAp

    lErro = CF("OS_Le_ComItens", objOP)
    If lErro <> SUCESSO And lErro <> 30401 Then gError ERRO_SEM_MENSAGEM
    
    'se a ordem de producao nao tiver itens associados ==> erro
    If lErro = 30401 Then gError 206280

    For Each objItemOP In objOP.colItens

        lErro = OP_Requisicoes_Gera1(objItemOP, objMovEstoque)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next

    OP_Requisicoes_Gera = SUCESSO

    Exit Function

Erro_OP_Requisicoes_Gera:

    OP_Requisicoes_Gera = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206280
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMPRODUCAO_SEM_ITENS", gErr, objOP.sCodigo, objOP.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206281)

    End Select

    Exit Function

End Function

Private Function OP_Requisicoes_Gera1(ByVal objItemOP As ClassItemOS, ByVal objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objOSAp As ClassOSAp
Dim objOSApMO As ClassOSApMO
Dim objOSApMaq As ClassOSApMaq
Dim objItemOSOp As ClassItemOSOperacoes
Dim objItemOSOpMO As ClassItemOSOperacoesMO
Dim objItemOSOpMaq As ClassItemOSOperacoesMaquinas
Dim objItemOSOpMP As ClassItemOSOperacoesPecas
Dim dFator As Double, dFatorUM As Double, dFatorOP As Double
Dim bAchou As Double, objItemMovEstoque As ClassItemMovEstoque
Dim colMO As Collection

On Error GoTo Erro_OP_Requisicoes_Gera1

'ACUMULA e PREENCHE A COLECAO DE MOVESTOQUE
    Set objOSAp = objMovEstoque.objApontamento
    
    If StrParaDbl(QuantidadeOP.Text) > 0 Then
        dFatorOP = StrParaDbl(QuantidadeOP.Text) / objItemOP.dQuantidade
    Else
        dFatorOP = 1
    End If

    For Each objItemOSOp In objItemOP.colItemOSOperacoes
        
        For Each objItemOSOpMO In objItemOSOp.colMO
        
            bAchou = False
            For Each objOSApMO In objOSAp.colMO
                If objOSApMO.iTipo = objItemOSOpMO.lCodigoMO And objOSApMO.sProdutoOS = objItemOP.sServico Then
                    objOSApMO.dHorasGastas = objOSApMO.dHorasGastas + objItemOSOpMO.dHoras * dFatorOP
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objOSApMO = New ClassOSApMO
                Set colMO = New Collection
                objOSAp.colMO.Add objOSApMO
                
                objOSApMO.dHorasGastas = objItemOSOpMO.dHoras * dFatorOP
                objOSApMO.iSeq = objOSAp.colMO.Count
                objOSApMO.iTipo = objItemOSOpMO.lCodigoMO
                objOSApMO.sOS = objItemOP.sCodigo
                objOSApMO.sProdutoOS = objItemOP.sServico
                
                lErro = CF("MO_Le_Tipo", objOSApMO.iTipo, colMO)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If colMO.Count > 0 Then objOSApMO.lCodigoMO = colMO.Item(1).lCodigo
                
            End If
            
        Next
        
        For Each objItemOSOpMaq In objItemOSOp.colMaquinas
        
            bAchou = False
            For Each objOSApMaq In objOSAp.colMaq
                If objOSApMaq.lNumIntDocMaq = objItemOSOpMaq.lNumIntDocMaq And objOSApMaq.sProdutoOS = objItemOP.sServico Then
                    objOSApMaq.dHorasGastas = objOSApMaq.dHorasGastas + objItemOSOpMaq.dHoras * dFatorOP
                    If objOSApMaq.iQuantidade < objItemOSOpMaq.iQuantidade Then objOSApMaq.iQuantidade = objItemOSOpMaq.iQuantidade
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objOSApMaq = New ClassOSApMaq
                objOSAp.colMaq.Add objOSApMaq
                
                objOSApMaq.dHorasGastas = objItemOSOpMaq.dHoras * dFatorOP
                objOSApMaq.iSeq = objOSAp.colMaq.Count
                objOSApMaq.lNumIntDocMaq = objItemOSOpMaq.lNumIntDocMaq
                objOSApMaq.sOS = objItemOP.sCodigo
                objOSApMaq.sProdutoOS = objItemOP.sServico
                objOSApMaq.iQuantidade = objItemOSOpMaq.iQuantidade
                
            End If
            
        Next
        
        For Each objItemOSOpMP In objItemOSOp.colPecas
        
            'Carrega sProdutoFormatado em objProduto
            objProduto.sCodigo = objItemOSOpMP.sProduto
    
            'Tenta ler o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
            'Ausência de Produto no BD
            If lErro = 28030 Then gError 206282
        
            If objItemOSOpMP.iComposicao = PRODUTOKIT_COMPOSICAO_FIXA Then
                dFator = 1
            Else
                dFator = dFatorOP
            End If
   
            If Trim(objProduto.sContaContabilProducao) = "" Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
            
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
            
                If lErro = 22531 Then gError 206283
            
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                        
            End If
    
            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                bAchou = False
                For Each objItemMovEstoque In objMovEstoque.colItens
                    If objItemOSOpMP.sProduto = objItemMovEstoque.sProduto And objItemMovEstoque.sProdutoOP = objItemOP.sServico Then
                    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemOSOpMP.sUMProduto, objItemMovEstoque.sSiglaUM, dFatorUM)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                        bAchou = True
                        objItemMovEstoque.dQuantidade = objItemMovEstoque.dQuantidade + objItemOSOpMP.dQuantidade * dFator * dFatorUM
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    objMovEstoque.colItens.Add 0, 0, 0, 0, objItemOSOpMP.sProduto, objProduto.sDescricao, objItemOSOpMP.sUMProduto, objItemOSOpMP.dQuantidade * dFator, 0, "", 0, "", 0, objItemOP.sCodigo, objItemOP.sServico, objProduto.sContaContabilProducao, "", 0, Nothing, Nothing, DATA_NULA
                End If
                
            End If
                       
        Next

    Next

    OP_Requisicoes_Gera1 = SUCESSO

    Exit Function

Erro_OP_Requisicoes_Gera1:

    OP_Requisicoes_Gera1 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206282
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 206283
             Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206284)

    End Select

    Exit Function
    
End Function

Public Function ItemOP_Requisicao_Gera(objItemOP As ClassItemOS, objMovEstoque As ClassMovEstoque) As Long
'Gera Requisições para o Produto da OP a partir da quantidade que falta produzir

Dim lErro As Long
Dim objOSAp As New ClassOSAp

On Error GoTo Erro_ItemOP_Requisicao_Gera

    Set objMovEstoque.objApontamento = objOSAp

    lErro = CF("ItemOS_Le", objItemOP)
    If lErro <> SUCESSO And lErro <> 34711 Then gError ERRO_SEM_MENSAGEM
    
    If lErro = 34711 Then gError 206285

    lErro = OP_Requisicoes_Gera1(objItemOP, objMovEstoque)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    ItemOP_Requisicao_Gera = SUCESSO

    Exit Function

Erro_ItemOP_Requisicao_Gera:

    ItemOP_Requisicao_Gera = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206285
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206286)

    End Select

    Exit Function

End Function

Public Sub OPCodigoPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOS
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_OPCodigoPadrao_Validate

    If Len(Trim(OPCodigoPadrao.Text)) > 0 Then

        objOrdemProducao.sCodigo = OPCodigoPadrao.Text
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa

        lErro = CF("OS_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 195102 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO Then gError 206287

        If Len(Trim(OP.Text)) = 0 Then OP.Text = OPCodigoPadrao.Text

    End If

    Exit Sub

Erro_OPCodigoPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206287
            Call Rotina_Erro(vbOKOnly, "ERRO_OSCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206288)

    End Select
    
    Exit Sub

End Sub

Public Sub CclPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    'Verifica se o CclPadrao foi Preenchida
    If Len(Trim(CclPadrao.ClipText)) > 0 Then

        lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 5703 Then gError 206289

    End If

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206289
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206290)

    End Select

    Exit Sub

End Sub

Public Sub AlmoxPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadrao_Validate

    'Verifica se Almoxarifado Padrão foi informado
    If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then

        lErro = TP_Almoxarifado_Filial_Le(AlmoxPadrao, objAlmoxarifado, 0)
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError ERRO_SEM_MENSAGEM

        If lErro = 25136 Then gError 206291

        If lErro = 25143 Then gError 206292

    End If

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206291, 206292
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206293)

    End Select

    Exit Sub

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206294)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206295)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206296)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206297)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer, lNumIntDoc As Long
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        'Verifica se produto está preenchido
        sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)
    
        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
        End If
    
        If gobjMovEst.colItens.Count >= GridMovs.Row And GridMovs.Row <> 0 Then
            lNumIntDoc = gobjMovEst.colItens(GridMovs.Row).lNumIntDoc
        Else
            lNumIntDoc = 0
        End If
    
        Select Case objControl.Name
        
            Case Produto.Name
    
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    Produto.Enabled = False
                Else
                    Produto.Enabled = True
                End If
    
            Case UnidadeMed.Name
    
                If iProdutoPreenchido <> PRODUTO_PREENCHIDO Or lNumIntDoc <> 0 Or left(GridMovs.TextMatrix(GridMovs.Row, 0), 1) = "#" Then
                    UnidadeMed.Enabled = False
                Else
                    
                    UnidadeMed.Enabled = True
        
                    objClasseUM.iClasse = objProduto.iClasseUM
        
                    'Preenche a List da Combo UnidadeMed com as UM's do Produto
                    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    'Guardo o valor da Unidade de Medida da Linha
                    sUnidadeMed = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
        
                    'Limpar as Unidades utilizadas anteriormente
                    UnidadeMed.Clear
        
                    For Each objUnidadeDeMedida In colSiglas
                        UnidadeMed.AddItem objUnidadeDeMedida.sSigla
                    Next
        
                    UnidadeMed.AddItem ""
        
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
        
            Case Quantidade.Name, Almoxarifado.Name, OPCodigo.Name, ProdutoOP.Name
            
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 And left(GridMovs.TextMatrix(GridMovs.Row, 0), 1) <> "#" Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case Ccl.Name, ContaContabilEst.Name, ContaContabilProducao.Name
    
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case Lote.Name
    
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And left(GridMovs.TextMatrix(GridMovs.Row, 0), 1) <> "#" Then

                    If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                        objControl.Enabled = False
                    Else
                        objControl.Enabled = True
                    End If
                Else
                    objControl.Enabled = False
                End If
            
            Case FilialOP.Name
        
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And left(GridMovs.TextMatrix(GridMovs.Row, 0), 1) <> "#" Then
        
                    If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                Else
                    objControl.Enabled = False
                End If
            
            Case Estorno.Name
        
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc <> 0 Then
                    Estorno.Enabled = True
                Else
                    Estorno.Enabled = False
                End If
        
            Case Benef.Name
            
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
                    Benef.Enabled = True
                Else
                    Benef.Enabled = False
                End If
                
            Case MONomeRed.Name, MOTipo.Name
                objControl.Enabled = False
                
            Case Maquina.Name, MOCodigo
                objControl.Enabled = True
                
            Case MaqOS.Name, MaqServico.Name, MaqQtd.Name, MaqHoras.Name
                If Len(Trim(GridMaq.TextMatrix(GridMaq.Row, iGrid_Maquina_Col))) > 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            
            Case MOOS.Name, MOServico.Name, MOHoras.Name
                If Len(Trim(GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col))) > 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            
        End Select

    End If
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206298)

    End Select

    Exit Sub

End Sub

Public Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridMovs_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer
Dim lNumIntDoc As Long
Dim iLinhasExistentes As Integer
Dim lErro As Long

On Error GoTo Erro_GridMovs_KeyDown
     
    If gobjMovEst.colItens.Count >= GridMovs.Row Then
        lNumIntDoc = gobjMovEst.colItens(GridMovs.Row).lNumIntDoc
    Else
        lNumIntDoc = 0
    End If

    If lNumIntDoc = 0 Then
        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then
            'Guarda iLinhasExistentes
            iLinhasExistentesAnterior = objGrid.iLinhasExistentes

            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridMovs.Row

        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
            'Exclui de colItens o Item correspondente, se houver
            gobjMovEst.colItens.Remove iLinhaAnterior

            gcolcolRastreamentoSerie.Remove iLinhaAnterior

'            Call gobjTelaProjetoInfo.Remove_Linha(iLinhaAnterior)

            For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
                If gobjMovEst.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
                    GridMovs.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
                Else
                    GridMovs.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
                End If
                
            Next

            GridMovs.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

            lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    End If

    Exit Sub
    
Erro_GridMovs_KeyDown:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206299)
    
    End Select
    
    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig
Dim iPossuiGrade As Integer
Dim sProdutoPai As String
Dim iIndice As Integer
Dim colItensRomaneioGrade As New Collection
Dim objItemMovEst As New ClassItemMovEstoque
Dim objRomaneioGrade As New ClassRomaneioGrade
Dim sProdPreenchido As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) <> 0 Then
    
        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError ERRO_SEM_MENSAGEM

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError ERRO_SEM_MENSAGEM
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 206300

            If Trim(objProduto.sContaContabilProducao) = "" Then
            
                objTipoDeProduto.iTipo = objProduto.iTipo
        
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
        
                If lErro = 22531 Then gError 206302
        
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
                                
                'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
                If Trim(objProduto.sContaContabilProducao) = "" Then
                            
                    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                    objCTBConfig.iFilialEmpresa = giFilialEmpresa
                    
                    lErro = CF("CTBConfig_Le", objCTBConfig)
                    If lErro <> SUCESSO And lErro <> 9755 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
            
                End If
                                
            End If

            'se nao for um produto pai de grade
            If Len(Trim(objProduto.sGrade)) = 0 Then

                If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 206301

                lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                lErro = Preenche_ContaContabilEst(GridMovs.Row, Produto.Text)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then
                    lErro = QuantDisponivel_Calcula1(Produto.Text, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), objProduto)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                Else
                    lErro = QuantLote_Calcula1(Produto.Text, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)), objProduto)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                End If
        
                If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                    GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
                End If

            'se for um produto pai de grade
            Else
            
                'Verifica se há filhos válidos da grade pai
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Se nao existir, erro
                If colItensRomaneioGrade.Count = 0 Then gError 206303
                
                objItemMovEst.sProduto = objProduto.sCodigo
                objItemMovEst.sDescricao = objProduto.sDescricao
                objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
                objItemMovEst.sOPCodigo = OPCodigoPadrao.Text
                objItemMovEst.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
                        
                Set objRomaneioGrade = New ClassRomaneioGrade
                
                objRomaneioGrade.sNomeTela = Me.Name
                Set objRomaneioGrade.objTela = Me
                
                Set objRomaneioGrade.objObjetoTela = objItemMovEst
                            
                Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
                If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
        
                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                lErro = Preenche_ContaContabilEst(GridMovs.Row, Produto.Text)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206300
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 206301
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206302
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206303
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206304)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        If gobjMovEst.colItens(GridMovs.Row).lNumIntDoc = 0 Then

            If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then

                lErro = Testa_QuantRequisitada(StrParaDbl(Quantidade.Text))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End If

        End If

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(Quantidade.Text), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206305)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError ERRO_SEM_MENSAGEM

        If lErro = 25130 Then gError 206306

        If lErro = 25135 Then gError 206307

        GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            lErro = QuantLote_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    Call Preenche_ContaContabilEst(GridMovs.Row, GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
            
    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206306

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
            If vbMsg = vbYes Then
                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 206307

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))
            If vbMsg = vbYes Then
                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206308)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_OPCodigo(objGridInt As AdmGrid, objControle As Object) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String, sProdutoOPEnxuto As String
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOS, objItemOPUnico As ClassItemOS
Dim objOrdemProducao As New ClassOS
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim iCont As Integer
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_OPCodigo

    Set objGridInt.objControle = objControle
    
    If Len(Trim(objControle.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = objControle.Text

        lErro = CF("OS_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 195102 Then gError ERRO_SEM_MENSAGEM

        If lErro = 195102 Then gError 206309

        If iFrameAtualItens = 1 Then
            sProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col)
        ElseIf iFrameAtualItens = 2 Then
            sProduto = GridMO.TextMatrix(GridMO.Row, iGrid_MOServico_Col)
        Else
            sProduto = GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqServico_Col)
        End If
        
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = objControle.Text
            objItemOP.sServico = sProdutoFormatado

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 206310
            
        Else
        
            lErro = CF("OS_Le_ComItens", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 195060 Then gError ERRO_SEM_MENSAGEM
            
            For Each objItemOP In objOrdemProducao.colItens
                If objItemOPUnico Is Nothing Then Set objItemOPUnico = objItemOP
                iCont = iCont + 1
            Next
            
            If iCont = 1 Then
            
                'Mascara o ProdutoOP
                lErro = Mascara_RetornaProdutoEnxuto(objItemOPUnico.sServico, sProdutoOPEnxuto)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                If iFrameAtualItens = 1 Then
                    ProdutoOP.PromptInclude = False
                    ProdutoOP.Text = sProdutoOPEnxuto
                    ProdutoOP.PromptInclude = True
            
                    GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = ProdutoOP.Text
                ElseIf iFrameAtualItens = 2 Then
                    MOServico.PromptInclude = False
                    MOServico.Text = sProdutoOPEnxuto
                    MOServico.PromptInclude = True
            
                    GridMO.TextMatrix(GridMO.Row, iGrid_MOServico_Col) = MOServico.Text
                Else
                    MaqServico.PromptInclude = False
                    MaqServico.Text = sProdutoOPEnxuto
                    MaqServico.PromptInclude = True
            
                    GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqServico_Col) = MaqServico.Text
                End If
                    
            End If
        End If
          
    Else
        
        If iFrameAtualItens = 1 Then
            ProdutoOP.PromptInclude = False
            ProdutoOP.Text = ""
            ProdutoOP.PromptInclude = True
    
            GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = ""
        ElseIf iFrameAtualItens = 2 Then
            MOServico.PromptInclude = False
            MOServico.Text = ""
            MOServico.PromptInclude = True
    
            GridMO.TextMatrix(GridMO.Row, iGrid_MOServico_Col) = ""
        Else
            MaqServico.PromptInclude = False
            MaqServico.Text = ""
            MaqServico.PromptInclude = True
    
            GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqServico_Col) = ""
        End If

    End If

    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_OPCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_OPCodigo:

    Saida_Celula_OPCodigo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206309
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OSCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("OrdemProducao", objOrdemProducao)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 206310
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206311
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsg = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206312)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoOP(objGridInt As AdmGrid, objControle As Object) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemOP As New ClassItemOS
Dim objKit As New ClassKit
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iAchou As Integer, iIndice As Integer
Dim vbMsg As VbMsgBoxResult
Dim sOS As String

On Error GoTo Erro_Saida_Celula_ProdutoOP

    Set objGridInt.objControle = objControle
    iAchou = 0

    If Len(Trim(objControle.ClipText)) > 0 Then

        lErro = CF("Produto_Critica", objControle.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError ERRO_SEM_MENSAGEM

        If lErro = 25041 Then gError 206312
        
        If iFrameAtualItens = 1 Then
            sOS = GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col)
        ElseIf iFrameAtualItens = 2 Then
            sOS = GridMO.TextMatrix(GridMO.Row, iGrid_MOOS_Col)
        Else
            sOS = GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqOS_Col)
        End If

        If Len(Trim(sOS)) > 0 Then

            objItemOP.sCodigo = sOS
            objItemOP.iFilialEmpresa = giFilialEmpresa

            objItemOP.sServico = objProduto.sCodigo

            lErro = CF("ItemOS_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 206313

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ProdutoOP = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoOP:

    Saida_Celula_ProdutoOP = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206312
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoOP.Text)
            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdutoOP.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 206313
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PARTICIPA_OS", gErr, objItemOP.sServico, objItemOP.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206314)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Trim(Ccl.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError ERRO_SEM_MENSAGEM

        If lErro = 5703 Then gError 206315

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206315
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsgRes = vbYes Then
                objCcl.sCcl = sCclFormatada
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("CclTela", objCcl)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206316)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    If Len(UnidadeMed.Text) > 0 Then

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
        If gobjMovEst.colItens(GridMovs.Row).lNumIntDoc = 0 Then
    
            'Se a quantidade está preenchida e não se trata de estorno
            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then
    
                dQuantidade = CDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))
    
                'Testa a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            End If
    
        End If

    Else
    
        QuantDisponivel.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206317)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Apontamento

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206318)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsgRes As VbMsgBoxResult
Dim sReqMaiorOP As String
Dim dtData As Date

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 206319

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 206320

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovs
    'If objGrid.iLinhasExistentes = 0 Then gError 206321

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 206322

        'Verifica se o Almoxarifado foi informado
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 And left(GridMovs.TextMatrix(iIndice, 0), 1) <> "#" Then gError 206323

        'Verifica se OPCodigo foi preenchido
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_OPCodigo_Col))) = 0 And left(GridMovs.TextMatrix(iIndice, 0), 1) <> "#" Then gError 206324

        'Verifica se ProdutoOP foi preenchido
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_ProdutoOP_Col))) = 0 And left(GridMovs.TextMatrix(iIndice, 0), 1) <> "#" Then gError 206325

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 206326

    Next
    
    Set objMovEstoque.objApontamento = New ClassOSAp
    objMovEstoque.objApontamento.lCodigo = StrParaLong(Codigo.Text)
    objMovEstoque.objApontamento.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovimentoEstoqueAp_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2")
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

    End If
    
    dtData = StrParaDate(Data.Text)
    If dtData = DATA_NULA Then dtData = Date

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(dtData)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    '###############################################
'    'Inserido por Wagner 10/08/2006
'    lErro = gobjTelaProjetoInfo.Valida_Dados
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'    '###############################################
    
    'Grava no BD(inclusive dados contábeis)
    lErro = CF("MovEstoque_ApSRV_Grava", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    'gravar anotacao, se houver
'    If Not (gobjAnotacao Is Nothing) Then
'
'        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
'
'            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
'            gobjAnotacao.sID = CStr(objMovEstoque.iFilialEmpresa) & "," & CStr(objMovEstoque.lCodigo)
'            gobjAnotacao.dtDataAlteracao = gdtDataHoje
'
'            lErro = CF("Anotacoes_Grava", gobjAnotacao)
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        End If
'
'    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206319
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
            
        Case 206320
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
            Data.SetFocus

        Case 206321
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 206322
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 206323
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 206324
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCODIGO_NAO_PREENCHIDO", gErr, iIndice)

        Case 206325
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO", gErr, iIndice)

        Case 206326
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206327)

    End Select

    Exit Function

End Function

Private Sub MovimentoEstorno(iMovimento As Integer)

Dim iIndice As Integer

    'Verifica se existe alguma Linha com Estorno marcado.
    For iIndice = 1 To objGrid.iLinhasExistentes

        If GridMovs.TextMatrix(iIndice, iGrid_Estorno_Col) = "1" Then
            iMovimento = MOVIMENTO_ESTORNO
            Exit For
        End If
    Next

    If iMovimento <> MOVIMENTO_ESTORNO Then iMovimento = MOVIMENTO_NORMAL

End Sub

Private Function Move_Tela_Memoria(objMovEstoque As ClassMovEstoque) As Long
'Preenche objMovEstoque (inclusive colItens)

Dim iIndice As Integer
Dim lCodigo As Long
Dim lErro As Long
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Data.Text)) > 0 Then
        objMovEstoque.dtData = StrParaDate(Data.Text)
    Else
        objMovEstoque.dtData = DATA_NULA
    End If

'hora
    If Len(Trim(HORA.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(HORA.Text)
    Else
        objMovEstoque.dtHora = Time
    End If

    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    objMovEstoque.iTipoMov = 0

    For iIndice = 1 To objGrid.iLinhasExistentes

        lErro = Move_Itens_Memoria(iIndice, objMovEstoque)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next
    
    lErro = Move_Ap_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    '##################################################
'    'Inserido por Wagner 09/08/2006
'    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_REQPROD)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    Set objMovEstoque.objProjetoInfo = objProjetoInfo
'    '##################################################

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206328)

    End Select

    Exit Function

End Function

Private Function Move_Ap_Memoria(objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objOSAp As New ClassOSAp

On Error GoTo Erro_Move_Ap_Memoria

    Set objMovEstoque.objApontamento = objOSAp
    
    objOSAp.lCodigo = StrParaLong(Codigo.Text)
    objOSAp.lCliente = objMovEstoque.lCliente
    objOSAp.iFilialEmpresa = objMovEstoque.iFilialEmpresa
    objOSAp.dtData = objMovEstoque.dtData
    objOSAp.dtHora = objMovEstoque.dtHora

    lErro = Move_MO_Memoria(objOSAp.colMO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Move_Maq_Memoria(objOSAp.colMaq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Move_Ap_Memoria = SUCESSO

    Exit Function

Erro_Move_Ap_Memoria:

    Move_Ap_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206329)

    End Select

    Exit Function
    
End Function

Private Function Move_MO_Memoria(ByVal colItens As Collection) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim objOSApMO As ClassOSApMO
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_MO_Memoria

    For iIndice = 1 To objGridMO.iLinhasExistentes
        Set objOSApMO = New ClassOSApMO
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMO.TextMatrix(iIndice, iGrid_MOServico_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 205891

        objOSApMO.lCodigoMO = StrParaLong(GridMO.TextMatrix(iIndice, iGrid_MOCodigo_Col))
        objOSApMO.dHorasGastas = StrParaDbl(GridMO.TextMatrix(iIndice, iGrid_MOHoras_Col))
        objOSApMO.sOS = GridMO.TextMatrix(iIndice, iGrid_MOOS_Col)
        objOSApMO.sProdutoOS = sProdutoFormatado
        
        If objOSApMO.lCodigoMO = 0 Then gError 205892
        If objOSApMO.dHorasGastas = 0 Then gError 205893
        If Len(Trim(objOSApMO.sOS)) = 0 Then gError 205894
        If Len(Trim(objOSApMO.sProdutoOS)) = 0 Then gError 205895
        
        objOSApMO.iSeq = iIndice
        colItens.Add objOSApMO
    Next

    Move_MO_Memoria = SUCESSO

    Exit Function

Erro_Move_MO_Memoria:

    Move_MO_Memoria = gErr

    Select Case gErr
    
        Case 205891

        Case 205892
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOMO_NAO_PREENCHIDO_GRID2", gErr, iIndice)

        Case 205893
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAS_NAO_PREENCHIDA_GRID2", gErr, iIndice)

        Case 205894
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCODIGO_NAO_PREENCHIDO", gErr, iIndice)

        Case 205895
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205896)

    End Select

    Exit Function
    
End Function

Private Function Move_Maq_Memoria(ByVal colItens As Collection) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim objOSApMaq As ClassOSApMaq
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMaq As ClassMaquinas

On Error GoTo Erro_Move_Maq_Memoria

    For iIndice = 1 To objGridMaq.iLinhasExistentes
    
        Set objOSApMaq = New ClassOSApMaq
        Set objMaq = New ClassMaquinas
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMaq.TextMatrix(iIndice, iGrid_MaqServico_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 205900
        
        objMaq.iFilialEmpresa = giFilialEmpresa
        objMaq.sNomeReduzido = GridMaq.TextMatrix(iIndice, iGrid_Maquina_Col)
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaq)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 205901

        objOSApMaq.lNumIntDocMaq = objMaq.lNumIntDoc
        objOSApMaq.iQuantidade = StrParaInt(GridMaq.TextMatrix(iIndice, iGrid_MaqQtd_Col))
        objOSApMaq.dHorasGastas = StrParaDbl(GridMaq.TextMatrix(iIndice, iGrid_MaqHoras_Col))
        objOSApMaq.sOS = GridMaq.TextMatrix(iIndice, iGrid_MaqOS_Col)
        objOSApMaq.sProdutoOS = sProdutoFormatado
        
        If objOSApMaq.lNumIntDocMaq = 0 Then gError 205902
        If objOSApMaq.dHorasGastas = 0 Then gError 205903
        If objOSApMaq.iQuantidade = 0 Then gError 205904
        If Len(Trim(objOSApMaq.sOS)) = 0 Then gError 205905
        If Len(Trim(objOSApMaq.sProdutoOS)) = 0 Then gError 205906
        
        objOSApMaq.iSeq = iIndice
        colItens.Add objOSApMaq
    Next
    
    Move_Maq_Memoria = SUCESSO

    Exit Function

Erro_Move_Maq_Memoria:

    Move_Maq_Memoria = gErr

    Select Case gErr
    
        Case 205900, 205901

        Case 205902
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 205903
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAS_NAO_PREENCHIDA_GRID2", gErr, iIndice)

        Case 205904
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 205905
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCODIGO_NAO_PREENCHIDO", gErr, iIndice)

        Case 205906
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205907)

    End Select

    Exit Function
    
End Function

Private Function Move_Itens_Memoria(iIndice As Integer, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim sProdutoOPFormatado As String
Dim iProdutoPreenchido As Integer
Dim iProdutoOPPreenchido As Integer
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objAlmoxarifado As ClassAlmoxarifado
Dim sContaFormatadaEst As String
Dim iContaPreenchida As Integer
Dim sContaFormatadaProducao As String
Dim colRastreamentoMovto As New Collection
Dim iTipoMovEstoque As Integer
Dim objItemMovEst As ClassItemMovEstoque
Dim lBenef As Long

On Error GoTo Erro_Move_Itens_Memoria

    With GridMovs
    
        Call Combo_Obtem_ItemData(Benef, .TextMatrix(iIndice, iGrid_Benef_Col), lBenef)

        'Se o Item for extorno
        If .TextMatrix(iIndice, iGrid_Estorno_Col) = MARCADO Then
            
            'verifica se o lNumIntDoc do elemento correspondente em colItens está zerado
            If gobjMovEst.colItens(iIndice).lNumIntDoc = 0 Then gError 206330
                        
            iTipoMovEstoque = MOV_EST_ESTORNO_APONT_SRV
            
        Else
            iTipoMovEstoque = MOV_EST_APONT_SRV
        End If
        
        'Pega o Código do Almoxarifado
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = .TextMatrix(iIndice, iGrid_Almoxarifado_Col)

        If Len(Trim(objAlmoxarifado.sNomeReduzido)) <> 0 Then

            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM
    
            If lErro = 25056 Then gError 206331

        End If

        sProdutoFormatado = ""
        sProdutoOPFormatado = ""

        If Len(Trim(.TextMatrix(iIndice, iGrid_Ccl_Col))) > 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", .TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Else

            sCclFormatada = ""

        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilEst_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilEst_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Else
            sContaFormatadaEst = ""
        End If
                
        If .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Else
            sContaFormatadaProducao = ""
        End If
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_ProdutoOP_Col), sProdutoOPFormatado, iProdutoOPPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Move os dados do rastreamento para a Memória
        lErro = Move_RastroEstoque_Memoria(iIndice, colRastreamentoMovto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set objItemMovEst = objMovEstoque.colItens.Add(gobjMovEst.colItens(iIndice).lNumIntDoc, iTipoMovEstoque, 0, 0, sProdutoFormatado, .TextMatrix(iIndice, iGrid_Descricao_Col), .TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, .TextMatrix(iIndice, iGrid_Almoxarifado_Col), 0, sCclFormatada, CLng(.TextMatrix(iIndice, iGrid_Estorno_Col)), .TextMatrix(iIndice, iGrid_OPCodigo_Col), sProdutoOPFormatado, sContaFormatadaProducao, sContaFormatadaEst, 0, colRastreamentoMovto, Nothing, DATA_NULA)
        
        Set objItemMovEst.colItensRomaneioGrade = gobjMovEst.colItens(iIndice).colItensRomaneioGrade

    End With

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206330
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case 206331
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206332)

    End Select

    Exit Function

End Function

Function Move_RastroEstoque_Memoria(iLinha As Integer, colRastreamentoMovto As Collection) As Long
'Move o Rastro dos Itens de Movimento

Dim objProduto As New ClassProduto, lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastreamentoMovto As New ClassRastreamentoMovto

On Error GoTo Erro_Move_RastroEstoque_Memoria
    
    lErro = CF("Produto_Formata", GridMovs.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    If lErro = 28030 Then gError 206333
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            
                objRastreamentoMovto.sLote = GridMovs.TextMatrix(iLinha, iGrid_Lote_Col)
            
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            
            'Se colocou o Número do Lote e não colocou a filial ---> Erro
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                
                If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 206334
            
                objRastreamentoMovto.sLote = GridMovs.TextMatrix(iLinha, iGrid_Lote_Col)
                objRastreamentoMovto.iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))
            
            End If
            
            
            'Se colocou a Filial da O.P. e não colocou o lote ---> Erro
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 Then
                'Se não colocou o Número do Lote ---> Erro
                If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 206335
            End If
           
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iLinha)
                objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
        End If
        
        'Se o Número do Lote
        If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
        
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then objRastreamentoMovto.dQuantidade = CDbl(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col))
            objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
            objRastreamentoMovto.sProduto = sProdutoFormatado
            
            If objProduto.iRastro <> PRODUTO_RASTRO_NUM_SERIE Then
                colRastreamentoMovto.Add objRastreamentoMovto
            End If
        
        End If
        
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 206333
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 206334
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 206335
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206336)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Apontamento

    'Coloca a Data atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206337)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Apontamento()

Dim lErro As Long
Dim iIndice As Integer

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)
    
    LblUM.Caption = ""
    
    For iIndice = 1 To objGrid.iLinhasExistentes
        GridMovs.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa o Grid
    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridMO)
    Call Grid_Limpa(objGridMaq)

    'prepara o Grid para permitir inserir e excluir Linhas
    objGrid.iProibidoIncluir = 0
    objGrid.iProibidoExcluir = 0
    Call Grid_Inicializa(objGrid)
    
    'Mostra a Data Atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Set gobjMovEst = New ClassMovEstoque

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    lCodigoAntigo = 0
    
    Set gcolcolRastreamentoSerie = New Collection

'    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
'    Set gobjTelaProjetoInfo.objUserControl = Me
'    Set gobjTelaProjetoInfo.objTela = Me
        
'    Set gobjAnotacao = Nothing

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

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

    Set gobjInfoUsu = Nothing

    Set gobjMovEst = Nothing

    Set objEventoCodigo = Nothing
    Set objEventoOPPadrao = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoOP = Nothing
    Set objEventoCcl = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoProdutoOP = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoProdutoOP = Nothing
    Set objEventoProdutoOP_OP = Nothing
    Set objEventoOP_OP = Nothing
    Set objEventoOP_OP_Todos = Nothing
    Set objEventoProdutoOP_Grid = Nothing
    Set objEventoRastroLote = Nothing
    Set objEventoMO = Nothing
    Set objEventoMaq = Nothing
    
    '???
    Set objEventoBotaoOP = Nothing
    
    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    Set objGridMO = Nothing
    Set objGridMaq = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set gcolcolRastreamentoSerie = Nothing
    
'    Set gobjTelaProjetoInfo = Nothing
    
'    Set gobjAnotacao = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
   
End Sub

Public Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CclPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub OPCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub OPCodigoPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AlmoxPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estorno_Click()

Dim lErro As Long

On Error GoTo Erro_Estorno_Click

    iAlterado = REGISTRO_ALTERADO

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub
    
Erro_Estorno_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206338)
    
    End Select
    
    Exit Sub

End Sub

Public Sub OP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoOPGera_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub GridMovs_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Public Sub GridMovs_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Public Sub GridMovs_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Public Sub GridMovs_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If


End Sub

Public Sub GridMovs_LeaveCell()
    If objGrid.iSaidaCelula = SUCESSO Then

    Call Saida_Celula(objGrid)

    End If

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objGridInt.objGrid Is GridMovs Then
        
            Select Case GridMovs.Col
    
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_OPCodigo_Col
    
                    lErro = Saida_Celula_OPCodigo(objGridInt, OPCodigo)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_Benef_Col
                    lErro = Saida_Celula_Benef(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_Almoxarifado_Col
    
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_Ccl_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_Estorno_Col
                    lErro = Saida_Celula_Padrao(objGridInt, Estorno)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_ProdutoOP_Col
                    lErro = Saida_Celula_ProdutoOP(objGridInt, ProdutoOP)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_ContaContabilEst_Col
                    lErro = Saida_Celula_ContaContabilEst(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_ContaContabilProducao_Col
                    lErro = Saida_Celula_ContaContabilProducao(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_Lote_Col
                    lErro = Saida_Celula_Lote(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case Else
                    lErro = Saida_Celula_Grid(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            End Select
            
        ElseIf objGridInt.objGrid Is GridMO Then
    
            Select Case GridMO.Col
    
                Case iGrid_MOCodigo_Col
                    lErro = Saida_Celula_MOCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_MOOS_Col
                    lErro = Saida_Celula_OPCodigo(objGridInt, MOOS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_MOServico_Col
                    lErro = Saida_Celula_ProdutoOP(objGridInt, MOServico)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_MOHoras_Col
                    lErro = Saida_Celula_Horas(objGridInt, MOHoras)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
            End Select
                
        ElseIf objGridInt.objGrid Is GridMaq Then

             Select Case GridMaq.Col
    
                Case iGrid_Maquina_Col
                    lErro = Saida_Celula_Maquina(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_MaqOS_Col
                    lErro = Saida_Celula_OPCodigo(objGridInt, MaqOS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_MaqServico_Col
                    lErro = Saida_Celula_ProdutoOP(objGridInt, MaqServico)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_MaqHoras_Col
                    lErro = Saida_Celula_Horas(objGridInt, MaqHoras)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
                Case iGrid_MaqQtd_Col
                    lErro = Saida_Celula_Padrao(objGridInt, MaqQtd)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
            End Select
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 206339

        iAlterado = REGISTRO_ALTERADO

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206339
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206340)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Benef(objGridInt As AdmGrid) As Long
'faz a critica da celula de Benef do grid que está deixando de ser a corrente


Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_Benef

    Set objGridInt.objControle = Benef
    
    If Len(Trim(Benef.Text)) > 0 Then

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col) = Benef.Text

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))) <> 0 Then

            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
                lErro = QuantDisponivel_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Else
                lErro = QuantLote_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
        Else
            QuantDisponivel.Caption = ""
        End If
    Else
        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Saida_Celula_Benef = SUCESSO

    Exit Function

Erro_Saida_Celula_Benef:

   Saida_Celula_Benef = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206341)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOS
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
    
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
        If lErro = 28030 Then gError 206342
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then

                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 206343

                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then

                If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col))) > 0 Then


                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col))

                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM

                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 206344

                    'Preenche a Quantidade do Lote
                    lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Else
        
                    'Preenche a Quantidade do Lote
                    lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                End If
                            
            'Se for rastro por série
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
            End If
            
        End If
    
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
            
    If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then

        lErro = Testa_QuantRequisitada(StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206342
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206343, 206344
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206345)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOP As New AdmFiliais
Dim iCodigo As Integer
Dim objOrdemProducao As New ClassOS
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 206346
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 206347
                    
        End If
        
        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) > 0 Then
                    
            lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objRastroLote.sCodigo = GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
            
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 206348
                            
                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
            
        End If
        
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
        
    If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then

        lErro = Testa_QuantRequisitada(StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206346
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206347
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206348
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206349)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Grid(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Grid

    Select Case GridMovs.Col

        Case iGrid_Produto_Col

            Set objGridInt.objControle = Produto

        Case iGrid_UnidadeMed_Col

            Set objGridInt.objControle = UnidadeMed

        Case iGrid_Quantidade_Col

            Set objGridInt.objControle = Quantidade

        Case iGrid_Almoxarifado_Col

            Set objGridInt.objControle = Almoxarifado

        Case iGrid_OPCodigo_Col

            Set objGridInt.objControle = OPCodigo

        Case iGrid_Ccl_Col

            Set objGridInt.objControle = Ccl

        Case iGrid_Estorno_Col

            Set objGridInt.objControle = Estorno

        Case iGrid_ProdutoOP_Col

            Set objGridInt.objControle = ProdutoOP
            
        Case iGrid_ContaContabilEst_Col
        
            Set objGridInt.objControle = ContaContabilEst

    End Select

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Grid = SUCESSO

    Exit Function

Erro_Saida_Celula_Grid:

    Saida_Celula_Grid = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206350)

    End Select

    Exit Function

End Function

Public Sub GridMovs_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Public Sub GridMovs_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Sub GridMovs_RowColChange()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_GridMovs_RowColChange

    Call Grid_RowColChange(objGrid)

    If (GridMovs.Row <> iLinhaAntiga) Then

        'Guarda a Linha usada
        iLinhaAntiga = GridMovs.Row

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

    End If

    Exit Sub

Erro_GridMovs_RowColChange:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206351)

    End Select

    Exit Sub

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


Public Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Estorno_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Estorno_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Estorno_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Estorno
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub OPCodigo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub OPCodigo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub OPCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = OPCodigo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

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

Public Sub ProdutoOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub ProdutoOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub ProdutoOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ProdutoOP
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function QuantDisponivel_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantDisponivel_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
    
        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado, objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206352)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Calcula1(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long
'descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lBenef As Long

On Error GoTo Erro_QuantDisponivel_Calcula1

    QuantDisponivel.Caption = ""

    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        If GridMovs.Row >= GridMovs.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
                If lErro = 28030 Then gError 206353
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM
    
            If lErro = 25056 Then gError 206354
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError ERRO_SEM_MENSAGEM
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Call Combo_Obtem_ItemData(Benef, GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col), lBenef)
                
                If lBenef = BENEF_COMBO_BENEF Then
                    QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef3 * dFator)
                ElseIf lBenef = BENEF_COMBO_DISP Then
                    QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
                Else
                    QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantOutras3 * dFator)
                End If
                
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivel.Caption = ""
    
        End If

    End If
    
    QuantDisponivel_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula1:

    QuantDisponivel_Calcula1 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206353
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 206354
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206355)

    End Select

    Exit Function

End Function

Private Function QuantLote_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantLote_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantLote_Calcula1(sProduto, sAlmoxarifado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Else
    
        lErro = QuantLote_Calcula1(sProduto, sAlmoxarifado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)), objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    QuantLote_Calcula = SUCESSO

    Exit Function

Erro_QuantLote_Calcula:

    QuantLote_Calcula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206356)

    End Select

    Exit Function

End Function

Private Function QuantLote_Calcula1(sProduto As String, sAlmoxarifado As String, sLote As String, iFilialOP As Integer, Optional objProduto As ClassProduto) As Long
'descobre a quantidade Lote e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLoteSaldo As New ClassRastreamentoLoteSaldo
Dim lBenef As Long

On Error GoTo Erro_QuantLote_Calcula1

    QuantDisponivel.Caption = ""

    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        If GridMovs.Row >= GridMovs.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
                If lErro = 28030 Then gError 206357
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM
    
            If lErro = 25056 Then gError 206358
    
            objRastreamentoLoteSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
            objRastreamentoLoteSaldo.sProduto = sProdutoFormatado
            objRastreamentoLoteSaldo.sLote = sLote
            objRastreamentoLoteSaldo.iFilialOP = iFilialOP
            
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoLoteSaldo)
            If lErro <> SUCESSO And lErro <> 78633 Then gError ERRO_SEM_MENSAGEM
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 78633 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Call Combo_Obtem_ItemData(Benef, GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col), lBenef)
        
                QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDispNossa * dFator)
    
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivel.Caption = ""
    
        End If

    End If
    
    QuantLote_Calcula1 = SUCESSO

    Exit Function

Erro_QuantLote_Calcula1:

    QuantLote_Calcula1 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206357
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 206358
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206359)

    End Select

    Exit Function

End Function

Private Function Testa_Quantidade() As Long

Dim dQuantidade As Double
Dim lErro As Long

On Error GoTo Erro_Testa_Quantidade

    If GridMovs.Row >= GridMovs.FixedRows And gobjMovEst.colItens.Count >= GridMovs.Row Then

        If gobjMovEst.colItens(GridMovs.Row).lNumIntDoc = 0 Then
    
            'Se a quantidade está preenchida e não se trata de linha estornada
            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then
    
                dQuantidade = CDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))
    
                'Calcula a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            End If
    
        End If

    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206360)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_QuantRequisitada(ByVal dQuantAtual As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoAtual As String
Dim sAlmoxarifado As String
Dim sAlmoxarifadoAtual As String
Dim sUnidadeAtual As String
Dim sUnidadeProd As String
Dim dQuantidadeProd As String
Dim dFator As Double
Dim objProduto As New ClassProduto, sLoteAtual As String, sLote As String
Dim dQuantTotal As Double, iFilialOPAtual As Integer, iFilialOP As Integer
Dim vbMsg As VbMsgBoxResult
Dim objControle As Control

On Error GoTo Erro_Testa_QuantRequisitada

    sProdutoAtual = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)
    sAlmoxarifadoAtual = GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col)
    sUnidadeAtual = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
    iFilialOPAtual = Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col))
    sLoteAtual = GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col)

    If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 Then

        lErro = CF("Produto_Formata", sProdutoAtual, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
        If lErro = 28030 Then gError 206361
    
        For iIndice = 1 To objGrid.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If GridMovs.Row <> iIndice Then
    
                sCodProduto = GridMovs.TextMatrix(iIndice, iGrid_Produto_Col)
                sAlmoxarifado = GridMovs.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
                iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(iIndice, iGrid_FilialOP_Col))
                sLote = GridMovs.TextMatrix(iIndice, iGrid_Lote_Col)
    
                lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) And iFilialOPAtual = iFilialOP And UCase(sLoteAtual) = UCase(sLote) Then
    
                    'Verifica se há alguma QuanTidade informada
                    If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 Then
    
                        sUnidadeProd = GridMovs.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
                        dQuantidadeProd = CDbl(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))
    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                        dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
    
                    End If
    
                End If
    
            End If
    
        Next
    
        dQuantTotal = dQuantTotal + dQuantAtual

        If dQuantTotal > CDbl(QuantDisponivel.Caption) Then
            GridMovs.Enabled = False
            vbMsg = Rotina_Aviso(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)
            GridMovs.Enabled = True
        End If

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206361
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206362)

    End Select

    Exit Function

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
Dim dQuantidade As Double
Dim sProduto As String
Dim sUM As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case QUANT_ESTOQUE
            For iLinha = 1 To objGrid.iLinhasExistentes
            
                If Len(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    If Len(GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    
                        lErro = CF("UMEstoque_Conversao", GridMovs.TextMatrix(iLinha, iGrid_Produto_Col), GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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
                objMnemonicoValor.colValor.Add CLng(Codigo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case DATA1
            If Len(Data.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case CCL1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_Ccl_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_Ccl_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ALMOXARIFADO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ESTORNO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                objMnemonicoValor.colValor.Add CInt(GridMovs.TextMatrix(iLinha, iGrid_Estorno_Col))
            Next

        Case PRODUTO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case UNIDADE_MED
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case DESCRICAO_ITEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_Descricao_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_Descricao_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case OP_CODIGO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_OPCodigo_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_OPCodigo_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case PRODUTO_OP
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_ProdutoOP_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_ProdutoOP_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case QUANTIDADE1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridMovs.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case CONTACONTABILEST1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_ContaContabilEst_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_ContaContabilEst_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case CONTACONTABILPROD1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovs.TextMatrix(iLinha, iGrid_ContaContabilProducao_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovs.TextMatrix(iLinha, iGrid_ContaContabilProducao_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case Else
            gError 206363

        End Select

        Calcula_Mnemonico = SUCESSO

        Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206363
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206364)

    End Select

    Exit Function

End Function

Private Function Preenche_ContaContabilEst(ByVal iLinha As Integer, sProduto As String) As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabilEst
        
    If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(sProduto)) > 0 Then
    
        'preenche o objEstoqueProduto
        objAlmoxarifado.sNomeReduzido = GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
        
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 25060 Then gError 206365
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 206366
        
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
        
            GridMovs.TextMatrix(iLinha, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
        
        End If
    
    End If
    
    Preenche_ContaContabilEst = SUCESSO
    
    Exit Function
    
Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = gErr
    
        Select Case gErr
            
            Case ERRO_SEM_MENSAGEM
            
            Case 206365
                 Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
            
            Case 206366
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
       
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206367)
        
        End Select
        
        Exit Function
        
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_PRODUCAO_MOVIMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Apontamento de Serviços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OSApontamento"
    
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
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is OPCodigoPadrao Then
            Call OPPadraoLabel_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is OP Then
            Call OPLabel_Click
        ElseIf Me.ActiveControl Is ProdutoOPGera Then
            Call ProdutoOPLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is ProdutoOP Then
            Call BotaoServicos_Click(0)
        ElseIf Me.ActiveControl Is MOServico Then
            Call BotaoServicos_Click(1)
        ElseIf Me.ActiveControl Is MaqServico Then
            Call BotaoServicos_Click(2)
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is OPCodigo Then
            Call BotaoOP_Click(0)
        ElseIf Me.ActiveControl Is MOOS Then
            Call BotaoOP_Click(1)
        ElseIf Me.ActiveControl Is MaqOS Then
            Call BotaoOP_Click(2)
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is ContaContabilEst Or Me.ActiveControl Is ContaContabilProducao Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Lote Then
            Call BotaoLote_Click
        ElseIf Me.ActiveControl Is MOCodigo Then
            Call BotaoMO_Click
        ElseIf Me.ActiveControl Is Maquina Then
            Call BotaoMaquinas_Click
        End If
    
    End If

End Sub


Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
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

'Verifica se já está em outra linha do Grid o Mesmo produto para mesma OP
Private Function Valida_Repeticao(ByVal sProdutoLinha As String, ByVal sOPLinha As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim sCodOP As String

On Error GoTo Erro_Valida_Repeticao

    If gobjEST.iOPRepeticaoItem = MARCADO Then
        For iIndice = 1 To objGrid.iLinhasExistentes
            If iIndice <> GridMovs.Row Then
                sProduto = GridMovs.TextMatrix(iIndice, iGrid_Produto_Col)
                sCodOP = GridMovs.TextMatrix(iIndice, iGrid_OPCodigo_Col)
                If sProduto = sProdutoLinha And sCodOP = sOPLinha Then gError 206368
            End If
        Next
    End If

    Valida_Repeticao = SUCESSO

    Exit Function

Erro_Valida_Repeticao:

    Valida_Repeticao = gErr

    Select Case gErr

         Case 206368
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE2", gErr, sProdutoLinha, sOPLinha, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206369)

    End Select

    Exit Function
    
End Function

Public Sub BotaoLote_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sLote As String
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_BotaoLote_Click

    If (GridMovs.Row = 0) Then gError 206370

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)
    sLote = GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 206371
    
    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))) = 0 Then gError 206372
    
    Set objAlmoxarifado = New ClassAlmoxarifado
    
    objAlmoxarifado.sNomeReduzido = GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(sLote)) > 0 Then
        objRastroLoteSaldo.sLote = sLote
    End If

    colSelecao.Add sProdutoFormatado
    colSelecao.Add objAlmoxarifado.iCodigo

    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoRastroLote, "Produto = ? AND Almoxarifado = ?")

    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206370
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 206371
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 206372
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, GridMovs.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206373)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRastroLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoRastroLote_evSelecao

    Set objRastroLoteSaldo = obj1

    If (GridMovs.Row > 0) Then
        Lote.Text = objRastroLoteSaldo.sLote
        
        'Carrega as séries na coleção global
        lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col) = objRastroLoteSaldo.sLote
    End If

    objProduto.sCodigo = objRastroLoteSaldo.sProduto
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    'Se for rastro por lote
     If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
        'Preenche a Quantidade do Lote
        lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else

        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoRastroLote_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206374)

    End Select

    Exit Sub

End Sub

Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim bPodeAlterarQtd As Boolean
Dim lBenef As Long

On Error GoTo Erro_BotaoSerie_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridMovs.Row = 0 Then gError 206375
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))) = 0 Then gError 206376
    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then gError 206377
    If StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)) = 0 Then gError 206378
        
    'Formata o produto
    lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
       
    objItemMovEstoque.dQuantidade = StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))
    objItemMovEstoque.iItemNF = GridMovs.Row
    objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col)
    objItemMovEstoque.sProduto = sProdutoFormatado
    objItemMovEstoque.sSiglaUM = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)

    Call Combo_Obtem_ItemData(Benef, GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col), lBenef)

    'Se o Item for estorno
    If GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) = MARCADO Then
        'Seta o Tipo como estorno
        If lBenef = BENEF_COMBO_BENEF Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_REQ_PRODUCAO_BENEF3
        ElseIf lBenef = BENEF_COMBO_OUTROS Then
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_REQ_PRODUCAO_OUTROS
        Else
            objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_REQ_PRODUCAO
        End If
    Else
        If lBenef = BENEF_COMBO_BENEF Then
            objItemMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3
        ElseIf lBenef = BENEF_COMBO_OUTROS Then
            objItemMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_OUTROS
        Else
            objItemMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO
        End If
    End If
                
    bPodeAlterarQtd = True
    If gobjMovEst.colItens.Count >= GridMovs.Row Then
        If gobjMovEst.colItens(GridMovs.Row).lNumIntDoc <> 0 Then
            bPodeAlterarQtd = False
        End If
    End If
    
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela_Modal("RastreamentoSerie", gcolcolRastreamentoSerie.Item(GridMovs.Row), objItemMovEstoque, Me.Name, bPodeAlterarQtd)
                    
    lErro = Acerta_Quantidade_Rastreada(GridMovs.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 206375
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 206376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 206377
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, GridMovs.Row)
        
        Case 206378
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, GridMovs.Row)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206379)
    
    End Select
    
    Exit Sub

End Sub

Public Function Carrega_Series(colRastreamentoMovto As Collection, ByVal dQuantidade As Double, ByVal sLoteIni As String, ByVal dQuantidadeAnterior As Double, ByVal sLoteIniAnterior As String, ByVal iLinha As Integer)
'Gera as séries a partir da série inicial e quantidade

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoSerie As ClassRastreamentoLote
Dim objRastreamentoSerieIni As ClassRastreamentoLote
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objAlmoxarifado As ClassAlmoxarifado
Dim vbResult As VbMsgBoxResult
Dim colRastreamentoMovtoAux As New Collection
Dim iTipoMovto As Integer
Dim lBenef As Long

On Error GoTo Erro_Carrega_Series

    'Formata o produto
    lErro = CF("Produto_Formata", GridMovs.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 206380

    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

        Call Combo_Obtem_ItemData(Benef, GridMovs.TextMatrix(iLinha, iGrid_Benef_Col), lBenef)

        'Se o Item for estorno
        If GridMovs.TextMatrix(iLinha, iGrid_Estorno_Col) = MARCADO Then
            'Seta o Tipo como estorno
            If lBenef = BENEF_COMBO_BENEF Then
                iTipoMovto = MOV_EST_ESTORNO_REQ_PRODUCAO_BENEF3
            ElseIf lBenef = BENEF_COMBO_OUTROS Then
                iTipoMovto = MOV_EST_ESTORNO_REQ_PRODUCAO_OUTROS
            Else
                iTipoMovto = MOV_EST_ESTORNO_REQ_PRODUCAO
            End If
        Else
            
            If lBenef = BENEF_COMBO_BENEF Then
                iTipoMovto = MOV_EST_REQ_PRODUCAO_BENEF3
            ElseIf lBenef = BENEF_COMBO_OUTROS Then
                iTipoMovto = MOV_EST_REQ_PRODUCAO_OUTROS
            Else
                iTipoMovto = MOV_EST_REQ_PRODUCAO
            End If
            
        End If
                
        If dQuantidadeAnterior <> 0 And Len(Trim(sLoteIniAnterior)) <> 0 And iTipoMovtoAnt = iTipoMovto Then
            
            If Abs(dQuantidade - dQuantidadeAnterior) > QTDE_ESTOQUE_DELTA Or sLoteIni <> sLoteIniAnterior Then
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_MODIFICACAO_SERIES")
                If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
            Else
                vbResult = vbNo
        
            End If
            
        Else
            vbResult = vbYes
            iTipoMovtoAnt = iTipoMovto
        End If
                    
        If vbResult = vbYes Then
                            
            If Len(Trim(sLoteIni)) <> 0 Then
            
                If Not IsNumeric(right(sLoteIni, objProduto.iSerieParteNum)) Then gError 206382
                
                Set objRastreamentoSerieIni = New ClassRastreamentoLote
                
                objRastreamentoSerieIni.sProduto = objProduto.sCodigo
                objRastreamentoSerieIni.iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))
                objRastreamentoSerieIni.sCodigo = sLoteIni
                
                lErro = CF("RastreamentoLote_Le", objRastreamentoSerieIni)
                If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM
                
                If lErro <> SUCESSO Then gError 206381
                              
                Set objItemMovEstoque = New ClassItemMovEstoque
                
                objItemMovEstoque.dQuantidade = Fix(dQuantidade)
                objItemMovEstoque.iItemNF = GridMovs.Row
                objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col)
                objItemMovEstoque.sProduto = sProdutoFormatado
                objItemMovEstoque.sSiglaUM = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
                objItemMovEstoque.iTipoMov = iTipoMovto
               
                Set objAlmoxarifado = New ClassAlmoxarifado
                
                objAlmoxarifado.sNomeReduzido = objItemMovEstoque.sAlmoxarifadoNomeRed
        
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM
                
                objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                
                lErro = CF("Rastreamento_Serie_Gera", objItemMovEstoque, objProduto, sLoteIni, colRastreamentoMovtoAux)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End If
            
            'Remove os dados anteriores
            For iIndice = colRastreamentoMovto.Count To 1 Step -1
                colRastreamentoMovto.Remove iIndice
            Next
            
            'Coloca os novos dados
            For Each objRastreamentoMovto In colRastreamentoMovtoAux
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
            
        End If
        
    End If

    Carrega_Series = SUCESSO
    
    Exit Function

Erro_Carrega_Series:

    Carrega_Series = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206380
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
               
        Case 206381
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo, objRastreamentoSerieIni.iFilialOP)
        
        Case 206382
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, right(sLoteIni, objProduto.iSerieParteNum))
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206383)

    End Select

    Exit Function

End Function

Public Function Acerta_Quantidade_Rastreada(ByVal iLinha As Integer)
'Acerta a quantidade do grid com base na quantidadse da coleção global de movimentos de séries

Dim lErro As Long
Dim dQuantidade As Double
Dim objRastreamentoSerie As ClassRastreamentoMovto

On Error GoTo Erro_Acerta_Quantidade_Rastreada

    For Each objRastreamentoSerie In gcolcolRastreamentoSerie.Item(iLinha)
    
        dQuantidade = dQuantidade + objRastreamentoSerie.dQuantidade
    
    Next
                
    If gobjMovEst.colItens(iLinha).lNumIntDoc = 0 Then

        If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(iLinha, iGrid_Estorno_Col) <> "1" Then

            lErro = Testa_QuantRequisitada(dQuantidade)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    End If
        
    GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    Acerta_Quantidade_Rastreada = SUCESSO
    
    Exit Function

Erro_Acerta_Quantidade_Rastreada:

    Acerta_Quantidade_Rastreada = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206384)

    End Select

    Exit Function

End Function

''#####################################
''Inserido por Wagner 03/08/2006
'Sub BotaoProjetos_Click()
'    Call gobjTelaProjetoInfo.BotaoProjetos_Click
'End Sub
'
'Sub LabelProjeto_Click()
'    Call gobjTelaProjetoInfo.LabelProjeto_Click
'End Sub
'
'Sub Projeto_GotFocus()
'    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
'End Sub
'
'Sub Projeto_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Sub Projeto_Validate(Cancel As Boolean)
'    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
'End Sub
'
'Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim objItensPRJCR As ClassItensPRJCR
'
'On Error GoTo Erro_Obter_ItensPRJCR
'
'    For iIndice = 1 To objGrid.iLinhasExistentes
'
'        Set objItensPRJCR = New ClassItensPRJCR
'
'        objItensPRJCR.sItem = CStr(iIndice)
'        objItensPRJCR.sDescricao = GridMovs.TextMatrix(iIndice, iGrid_Produto_Col)
'
'        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))
'
'        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_REQPROD
'
'        colItensPRJ.Add objItensPRJCR
'
'    Next
'
'    If colItensPRJ.Count = 0 Then
'        Set objItensPRJCR = New ClassItensPRJCR
'        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_REQPROD
'        colItensPRJ.Add objItensPRJCR
'    End If
'
'    Obter_ItensPRJCR = SUCESSO
'
'    Exit Function
'
'Erro_Obter_ItensPRJCR:
'
'    Obter_ItensPRJCR = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)
'
'    End Select
'
'    Exit Function
'End Function
'
'Sub Etapa_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Sub Etapa_Click()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Sub Etapa_Validate(Cancel As Boolean)
'    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
'End Sub
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
'
'Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)
'
'Dim lErro As Long
'
'On Error GoTo Erro_Anotacao_Extrai
'
'    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
'    If Len(Trim(Codigo.Text)) > 0 Then
'        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
'    Else
'        objAnotacao.sID = ""
'        If Not (gobjAnotacao Is Nothing) Then
'            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
'            objAnotacao.sTitulo = gobjAnotacao.sTitulo
'        End If
'    End If
'
'    Exit Sub
'
'Erro_Anotacao_Extrai:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206385)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)
'
'Dim lErro As Long
'
'On Error GoTo Erro_Anotacao_Preenche
'
'    'guarda o texto digitado
'    Set gobjAnotacao = objAnotacao
'
'    Exit Sub
'
'Erro_Anotacao_Preenche:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206386)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Public Sub TabStrip1_Click()

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtualItens Then

        If TabStrip_PodeTrocarTab(iFrameAtualItens, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameItens(TabStrip1.SelectedItem.Index).Visible = True
        FrameItens(iFrameAtualItens).Visible = False
        'Armazena novo valor de iFrameAtualItens
        iFrameAtualItens = TabStrip1.SelectedItem.Index
    
    End If

End Sub

Private Function Inicializa_GridMO(objGridInt As AdmGrid) As Long
'Inicializa o Grid

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridMO

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Mão-de-obra")
    objGridInt.colColuna.Add ("Nome Reduzido")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Horas Gastas")
    objGridInt.colColuna.Add ("O.S.")
    objGridInt.colColuna.Add ("Serviço OS")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (MOCodigo.Name)
    objGridInt.colCampo.Add (MONomeRed.Name)
    objGridInt.colCampo.Add (MOTipo.Name)
    objGridInt.colCampo.Add (MOHoras.Name)
    objGridInt.colCampo.Add (MOOS.Name)
    objGridInt.colCampo.Add (MOServico.Name)
    
    'Colunas do Grid
    iGrid_MOCodigo_Col = 1
    iGrid_MONomeRed_Col = 2
    iGrid_MOTipo_Col = 3
    iGrid_MOHoras_Col = 4
    iGrid_MOOS_Col = 5
    iGrid_MOServico_Col = 6
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMO

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMO.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMO = SUCESSO

    Exit Function

Erro_Inicializa_GridMO:

    Inicializa_GridMO = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206387)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridMaq(objGridInt As AdmGrid) As Long
'Inicializa o Grid

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridMaq

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Máquina")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Horas Gastas")
    objGridInt.colColuna.Add ("O.S.")
    objGridInt.colColuna.Add ("Serviço OS")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Maquina.Name)
    objGridInt.colCampo.Add (MaqQtd.Name)
    objGridInt.colCampo.Add (MaqHoras.Name)
    objGridInt.colCampo.Add (MaqOS.Name)
    objGridInt.colCampo.Add (MaqServico.Name)
    
    'Colunas do Grid
    iGrid_Maquina_Col = 1
    iGrid_MaqQtd_Col = 2
    iGrid_MaqHoras_Col = 3
    iGrid_MaqOS_Col = 4
    iGrid_MaqServico_Col = 5
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMaq

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMaq.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMaq = SUCESSO

    Exit Function

Erro_Inicializa_GridMaq:

    Inicializa_GridMaq = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206388)

    End Select

    Exit Function

End Function

Public Sub GridMaq_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq, iAlterado)
    End If

End Sub

Public Sub GridMaq_GotFocus()
    Call Grid_Recebe_Foco(objGridMaq)
End Sub

Public Sub GridMaq_EnterCell()
    Call Grid_Entrada_Celula(objGridMaq, iAlterado)
End Sub

Public Sub GridMaq_LeaveCell()
    Call Saida_Celula(objGridMaq)
End Sub

Public Sub GridMaq_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq, iAlterado)
    End If

End Sub

Public Sub GridMaq_RowColChange()
    Call Grid_RowColChange(objGridMaq)
End Sub

Public Sub GridMaq_Scroll()
    Call Grid_Scroll(objGridMaq)
End Sub

Public Sub GridMaq_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMaq)
End Sub

Public Sub GridMaq_LostFocus()
    Call Grid_Libera_Foco(objGridMaq)
End Sub

Public Sub GridMO_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Public Sub GridMO_GotFocus()
    Call Grid_Recebe_Foco(objGridMO)
End Sub

Public Sub GridMO_EnterCell()
    Call Grid_Entrada_Celula(objGridMO, iAlterado)
End Sub

Public Sub GridMO_LeaveCell()
    Call Saida_Celula(objGridMO)
End Sub

Public Sub GridMO_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Public Sub GridMO_RowColChange()
    Call Grid_RowColChange(objGridMO)
End Sub

Public Sub GridMO_Scroll()
    Call Grid_Scroll(objGridMO)
End Sub

Public Sub GridMO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMO)
End Sub

Public Sub GridMO_LostFocus()
    Call Grid_Libera_Foco(objGridMO)
End Sub

Public Sub MOCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MOCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Public Sub MOCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Public Sub MOCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOCodigo
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MOHoras_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MOHoras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Public Sub MOHoras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Public Sub MOHoras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOHoras
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MOOS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MOOS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Public Sub MOOS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Public Sub MOOS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOOS
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MOServico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MOServico_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Public Sub MOServico_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Public Sub MOServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOServico
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Maquina_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Maquina_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Public Sub Maquina_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Public Sub Maquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = Maquina
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MaqQtd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MaqQtd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Public Sub MaqQtd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Public Sub MaqQtd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqQtd
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MaqHoras_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MaqHoras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Public Sub MaqHoras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Public Sub MaqHoras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqHoras
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MaqOS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MaqOS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Public Sub MaqOS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Public Sub MaqOS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqOS
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MaqServico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MaqServico_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Public Sub MaqServico_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Public Sub MaqServico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqServico
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoMO_Click()

Dim lErro As Long
Dim objMO As New ClassMO
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is MOCodigo Then
        objMO.lCodigo = StrParaInt(MOCodigo.Text)
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMO.Row = 0 Then gError 206389

        objMO.lCodigo = StrParaInt(GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col))
    End If
    
    Call Chama_Tela("MaodeObraLista", colSelecao, objMO, objEventoMO)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 206389
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206390)

    End Select

    Exit Sub

End Sub

Public Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMO As ClassMO
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_objEventoMO_evSelecao

    Set objMO = obj1
    
    MOCodigo.Text = CStr(objMO.lCodigo)
    
    GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col) = CStr(objMO.lCodigo)
    GridMO.TextMatrix(GridMO.Row, iGrid_MONomeRed_Col) = objMO.sNomeReduzido

    objTiposDeMaodeObra.iCodigo = objMO.iTipo
    
    'Lê o TiposDeMaodeObra que está sendo Passado
    lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
    If lErro <> SUCESSO And lErro <> 137598 Then gError ERRO_SEM_MENSAGEM
    
    GridMO.TextMatrix(GridMO.Row, iGrid_MOTipo_Col) = objTiposDeMaodeObra.sDescricao
    GridMO.TextMatrix(GridMO.Row, iGrid_MOOS_Col) = OPCodigoPadrao.Text
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridMO.Row - GridMO.FixedRows = objGridMO.iLinhasExistentes Then
        objGridMO.iLinhasExistentes = objGridMO.iLinhasExistentes + 1
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206391)

    End Select

    Exit Sub

End Sub

Public Sub BotaoMaquinas_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaquinas_Click

    If GridMaq.Row <= objGridMaq.iLinhasExistentes Then

        If Me.ActiveControl Is Maquina Then
                
            objMaquinas.sNomeReduzido = Maquina.Text
            
        Else
        
            'Verifica se tem alguma linha selecionada no Grid
            If GridMaq.Row = 0 Then gError 206392
    
            objMaquinas.sNomeReduzido = GridMaq.TextMatrix(GridMaq.Row, iGrid_Maquina_Col)
            
        End If
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError ERRO_SEM_MENSAGEM
        
    End If
        
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaq)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206392
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206393)

    End Select

    Exit Sub
    
End Sub

Public Sub objEventoMaq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim iLinha As Integer
Dim iIndice As Integer

On Error GoTo Erro_objEventoMaq_evSelecao

    Set objMaquinas = obj1
        
    For iIndice = 1 To objGridMaq.iLinhasExistentes
        If iIndice <> GridMaq.Row Then
            'Tem a mesma máquina, erro
            If objMaquinas.sNomeReduzido = GridMaq.TextMatrix(iIndice, iGrid_Maquina_Col) Then
                gError 206394
            End If
        End If
    Next
        
    'Mostra os dados da Maquina na tela
    GridMaq.TextMatrix(GridMaq.Row, iGrid_Maquina_Col) = objMaquinas.sNomeReduzido
    GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqOS_Col) = OPCodigoPadrao.Text
    
    Maquina.Text = objMaquinas.sNomeReduzido
       
    If GridMaq.Row - GridMaq.FixedRows = objGridMaq.iLinhasExistentes Then
        objGridMaq.iLinhasExistentes = objGridMaq.iLinhasExistentes + 1
    End If
       
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaq_evSelecao:

    Select Case gErr

        Case 206394
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA", gErr, objMaquinas.sNomeReduzido, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206395)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objItemOP As New ClassItemOS
    
On Error GoTo Erro_BotaoProdutos_Click

    If GridMovs.Row = 0 Then gError 206396
       
    ' Formata o Produto
    lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado
        
    'Lista de produtos que  podem participar da producao e inventariados
    Call Chama_Tela("ProdutoEstoquePCPLista", colSelecao, objProduto, objEventoProduto)
   
   Exit Sub
   
Erro_BotaoProdutos_Click:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 206396
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206397)
        
    End Select
    
    Exit Sub
   
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig
Dim objItemMovEst As New ClassItemMovEstoque

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If GridMovs.Row = 0 Then gError 206398

    lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If (GridMovs.Col = iGrid_Produto_Col) Then
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 206399

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If lErro = 28030 Then gError 206400
        
        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        If Not (Me.ActiveControl Is Produto) Then

            GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = Produto.Text
    
            If Trim(objProduto.sContaContabilProducao) = "" Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
                
                If lErro = 22531 Then gError 206401
                
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                
                'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
                If Trim(objProduto.sContaContabilProducao) = "" Then
                                    
                    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                    objCTBConfig.iFilialEmpresa = giFilialEmpresa
                            
                    lErro = CF("CTBConfig_Le", objCTBConfig)
                    If lErro <> SUCESSO And lErro <> 9755 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
                    
                End If
                
            End If
    
            'Preenche a Linha do Grid
            lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Preenche_ContaContabilEst(GridMovs.Row, GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then
                lErro = QuantDisponivel_Calcula1(sProdutoMascarado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Else
                lErro = QuantLote_Calcula1(sProdutoMascarado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)), objProduto)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206398
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 206399
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, GridMovs.Row)

        Case 206400
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 206401
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206402)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Horas(objGridInt As AdmGrid, objControle As Object) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Horas

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objControle.Text = Formata_Estoque(objControle.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Horas = SUCESSO

    Exit Function

Erro_Saida_Celula_Horas:

    Saida_Celula_Horas = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206403)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMO As New ClassMO
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_MOCodigo

    Set objGridInt.objControle = MOCodigo
    
    'Se o campo foi preenchido
    If Len(MOCodigo.Text) > 0 Then
    
        objMO.lCodigo = StrParaLong(MOCodigo.Text)
        
        'Lê os dados da mao de obra a ser excluido
        lErro = CF("MO_Le", objMO)
        If lErro <> SUCESSO And lErro <> 193817 Then gError ERRO_SEM_MENSAGEM
    
        'Verifica se mao de obra não está cadastrado
        If lErro <> SUCESSO Then gError 206404
        
        objTiposDeMaodeObra.iCodigo = objMO.iTipo
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError ERRO_SEM_MENSAGEM
    
        If lErro <> SUCESSO Then gError 206405

        GridMO.TextMatrix(GridMO.Row, iGrid_MONomeRed_Col) = objMO.sNomeReduzido
        GridMO.TextMatrix(GridMO.Row, iGrid_MOTipo_Col) = objTiposDeMaodeObra.sDescricao
        If Len(Trim(OPCodigoPadrao.Text)) > 0 Then GridMO.TextMatrix(GridMO.Row, iGrid_MOOS_Col) = OPCodigoPadrao.Text
       
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_MOCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCodigo:

    Saida_Celula_MOCodigo = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
        Case 206404
            Call Rotina_Erro(vbOKOnly, "ERRO_MO_INEXISTENTE", gErr, objMO.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206405
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206406)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Maquina(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodigoItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim objMaquinas As New ClassMaquinas
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Maquina

    Set objGridInt.objControle = Maquina

    'Se o campo foi preenchido
    If Len(Trim(Maquina.Text)) > 0 Then
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", Maquina, objMaquinas)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        For iIndice = 1 To objGridMaq.iLinhasExistentes
            If iIndice <> GridMaq.Row Then
                'Tem a mesma máquina, erro
                If objMaquinas.sNomeReduzido = GridMaq.TextMatrix(iIndice, iGrid_Maquina_Col) Then
                    gError 206407
                End If
            End If
        Next
    
        GridMaq.TextMatrix(GridMaq.Row, iGrid_Maquina_Col) = objMaquinas.sNomeReduzido
        If Len(Trim(OPCodigoPadrao.Text)) > 0 Then GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqOS_Col) = OPCodigoPadrao.Text
           
        'verifica se precisa preencher o grid com uma nova linha
        If GridMaq.Row - GridMaq.FixedRows = objGridMaq.iLinhasExistentes Then
            objGridMaq.iLinhasExistentes = objGridMaq.iLinhasExistentes + 1
        End If
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Maquina = SUCESSO

    Exit Function

Erro_Saida_Celula_Maquina:

    Saida_Celula_Maquina = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 206407
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA", gErr, objMaquinas.sNomeReduzido, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206408)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub
