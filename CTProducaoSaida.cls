VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProducaoSaida"
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

Public gobjAnotacao As ClassAnotacoes

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoRastroLote As AdmEvento 'Inserido por Wagner
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

Dim gcolcolRastreamentoSerie As Collection 'Inserido por Wagner 13/03/2006
Dim iTipoMovtoAnt As Integer

'Declaração das Variáveis Globais
Public iAlterado As Integer
Dim iFrameAtual As Integer
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
Public iGrid_ProdutoOP_Col
Public iGrid_Estorno_Col As Integer
Public iGrid_ContaContabilEst_Col As Integer
Public iGrid_ContaContabilProducao_Col As Integer
Public iGrid_Lote_Col As Integer
Public iGrid_FilialOP_Col As Integer

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

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 2

'--- inicio dos properties get dos controles da tela

Public Property Get IgnorarReqProdAnt() As Object
     Set IgnorarReqProdAnt = objUserControl.Controls("IgnorarReqProdAnt")
End Property

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

'--- fim dos properties get dos controles da tela

Public Sub Benef_Click()

Dim lErro As Long

On Error GoTo Erro_Benef_Click

    iAlterado = REGISTRO_ALTERADO
    
    If GridMovs.Row <> 0 And objGrid.iLinhasExistentes <> 0 Then
    
        lErro = QuantDisponivel_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 91282
    
        '#################################################################
        'Inserido por Wagner 13/03/2006
        'Carrega as séries na coleção global
        If gcolcolRastreamentoSerie.Count >= GridMovs.Row Then
            lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
            If lErro <> SUCESSO Then gError 141911
        End If
        '#################################################################
    
    End If
    
    Exit Sub

Erro_Benef_Click:

    Select Case gErr
    
        Case 91282, 141911
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165376)
    
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
    If Len(Trim(Codigo.Text)) = 0 Then gError 89904

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
        
    '''03/09/01 - Marcelo inclusao da pergunta se deseja excluir a Requisicao
    'Envia aviso perguntando se realmente deseja excluir a Requisicao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_REQUISICAOPRODUCAO", objMovEstoque.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a producao
        lErro = CF("MovEstoque_ReqProducao_Exclui", objMovEstoque, objContabil)
        If lErro <> SUCESSO Then gError 89905

        Call Limpa_Tela_ProducaoSaida

        Data.PromptInclude = False
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
        Data.PromptInclude = True
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 89904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 89905
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165377)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGrade_Click()

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_BotaoGrade_Click

    If GridMovs.Row > 0 And GridMovs.Row <= objGrid.iLinhasExistentes Then
    
        Set objItemMovEst = gobjMovEst.colItens(GridMovs.Row)
        
        If objItemMovEst.iPossuiGrade = DESMARCADO Then gError 126587
        
        objItemMovEst.iBenef = StrParaInt(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
        objItemMovEst.sOPCodigo = OPCodigoPadrao.Text
        objItemMovEst.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
        objItemMovEst.sDescricao = objItemMovEst.sProdutoDesc
        objItemMovEst.sSiglaUMEst = objItemMovEst.sSiglaUM

        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        Set objRomaneioGrade.objTela = Me
        
        Set objRomaneioGrade.objObjetoTela = objItemMovEst
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
    
        Call Atualiza_Grid_Movimentos(objItemMovEst)
            
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
      
        Case 126587
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_GRADE", gErr, GridMovs.Row)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165378)
            
    End Select
    
    Exit Sub

End Sub

Sub Atualiza_Grid_Movimentos(objItemMovEst As ClassItemMovEstoque)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemMovEst.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        
    Next

    GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemMovEst.dQuantidade = dQuantidade

    Exit Sub

End Sub


Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("MovEstoque_Automatico", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then gError 55712

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    lCodigoAntigo = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 55712
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165379)
    
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

    If GridMovs.Row = 0 Then gError 52194
    
    If GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = "" Then gError 52195

    sConta = String(STRING_CONTA, 0)

    If GridMovs.Col = iGrid_ContaContabilEst_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52196
    
    ElseIf GridMovs.Col = iGrid_ContaContabilProducao_Col Then


        lErro = CF("Conta_Formata", ContaContabilProducao.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52197

    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 52194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 52195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 52196, 52197

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165380)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
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
        If lErro <> SUCESSO Then gError 52198
        
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

        Case 52198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165381)

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
        
            objMovEstoque.lCodigo = Codigo.Text
            
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_REQ_PRODUCAO)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 34776
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 78890
            
            If lErro = SUCESSO Then
            
                If objMovEstoque.iTipoMov <> MOV_EST_REQ_PRODUCAO Then gError 34898
            
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
            
                If vbMsg = vbNo Then gError 55324
            
                lErro = Preenche_Tela(objMovEstoque)
                If lErro <> SUCESSO Then gError 34907
            
            End If
        
        End If

        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
        
   End If

   Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 34776, 34907
    
        Case 34898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_PSAIDA", gErr, objMovEstoque.lCodigo)
            lCodigoAntigo = 0

        Case 55324
                
        Case 78890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165382)

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49639
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49640
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaMascarada
            ContaContabilEst.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEst.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49641
    
            'conta não cadastrada
            If lErro = 5700 Then gError 49642
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49717
    
    Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = gErr

    Select Case gErr

        Case 49639, 49641, 49717
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 49640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49642
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEst.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165383)

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 52045
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 52046
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaMascarada
            ContaContabilProducao.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 52047
    
            'conta não cadastrada
            If lErro = 5700 Then gError 52048
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52049
    
    Saida_Celula_ContaContabilProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducao:

    Saida_Celula_ContaContabilProducao = gErr

    Select Case gErr

        Case 52045, 52047, 52049
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52048
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducao.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165384)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao As String
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 13/03/2006

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
    Set objEventoRastroLote = New AdmEvento 'Inserido por Wagner

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me

    'Mostra a Data Atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78738
    
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
    If lErro <> SUCESSO Then gError 49663
    
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducao)
    If lErro <> SUCESSO Then gError 52199
    
    'Inicializa máscara de Produto , ProdutoOPGera , ProdutoOP
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 34715

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoOPGera)
    If lErro <> SUCESSO Then gError 34716

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoOP)
    If lErro <> SUCESSO Then gError 34717

    'Inicializa Máscara para CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 34718

    Ccl.Mask = sMascaraCclPadrao
    CclPadrao.Mask = sMascaraCclPadrao

    'Formata Quantidade
    Quantidade.Format = FORMATO_ESTOQUE

    'Inicializa GridMovs
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridMovs(objGrid)
    If lErro <> SUCESSO Then gError 34719

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then gError 39625

    iLinhaAntiga = 0
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 34715, 34716, 34717, 34718, 34719, 39625, 49663, 52199, 78738

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165385)

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
    If lErro <> SUCESSO Then gError 78739

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

        Case 78739 'Erro já tratado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165386)

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
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Escaninho")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ordem Produção")
    objGridInt.colColuna.Add ("Produto O.P.")
    objGridInt.colColuna.Add ("Lote / O.P./ Serie Ini.") 'Alterado por Wagner
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil de Estoque")
    objGridInt.colColuna.Add ("Conta Contábil de Produção")
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
    objGridInt.iLinhasVisiveis = 4

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165387)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela

    sTabela = "MovEstReqProd"

    'Lê os atributos de objMovEstoque que aparecem na Tela
    If Len(Trim(Codigo.ClipText)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    If Len(Trim(Data.ClipText)) <> 0 Then
        objMovEstoque.dtData = CDate(Data.Text)

    Else
        objMovEstoque.dtData = DATA_NULA

    End If

    If Len(Trim(HORA.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(HORA.Text)
    Else
        objMovEstoque.dtHora = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMovEstoque.lCodigo, 0, "Codigo"
    colCampoValor.Add "Data", objMovEstoque.dtData, 0, "Data"
    colCampoValor.Add "Hora", CDbl(objMovEstoque.dtHora), 0, "Hora"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "NumIntDocEst", OP_IGUAL, 0

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165388)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do Banco de Dados
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objMovEstoque.lCodigo = colCampoValor.Item("Codigo").vValor
    objMovEstoque.dtData = colCampoValor.Item("Data").vValor
    objMovEstoque.dtHora = colCampoValor.Item("Hora").vValor
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 34720

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 34720

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165389)

    End Select

    Exit Sub

End Sub

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long
'Preenche a Tela

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_Tela

    'Limpa a tela sem Fechar o Comando de setas
    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

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
    Set objMovEstoque.colItens = New colItensMovEstoque

    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovEstoqueItens_Le1", objMovEstoque, MOV_EST_REQ_PRODUCAO)
    If lErro <> SUCESSO And lErro <> 55387 Then gError 34721
    
    If lErro = 55387 Then gError 55402

    'Coloca os Dados na Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    Codigo.PromptInclude = True

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
    If lErro <> SUCESSO Then gError 34722

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.colItens(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39628

    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objMovEstoque.objProjetoInfo, objMovEstoque.lCodigo, PRJ_CR_TIPO_REQPROD, "", objMovEstoque.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 185850
    '###################################################
    
    lErro = CF2(Me, "ReqProd_Preenche_Tela", objMovEstoque)
    If lErro <> SUCESSO Then gError 185850

    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 34721, 34722, 39628, 91277, 117635, 185850

        Case 55402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_REQPRODUCAO", gErr, objMovEstoque.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165390)

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
Dim colRastreamentoSerie As Collection 'Inserido por Wagner 13/03/2006
Dim iBenef As Integer

On Error GoTo Erro_Preenche_GridMovs

    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 13/03/2006

    'Preenche gridmovs
    For Each objItemMovEstoque In colItens

        iIndice = iIndice + 1

        'Mascara o Produto
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 34723

        Call StrParaMasked2(Produto, sProdutoMascarado)
'        Produto.PromptInclude = False
'        Produto.Text = sProdutoMascarado
'        Produto.PromptInclude = True
        
        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilEst <> "" Then
        
            sContaEnxutaEst = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then gError 49718
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True
            
            GridMovs.TextMatrix(iIndice, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then
        
            sContaEnxutaProducao = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then gError 52055
            
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
        

'        If objItemMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3 Then
'            GridMovs.TextMatrix(iIndice, iGrid_Benef_Col) = "1"
'        Else
'            GridMovs.TextMatrix(iIndice, iGrid_Benef_Col) = "0"
'        End If

        If objItemMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3 Then
            iBenef = BENEF_COMBO_BENEF
        ElseIf objItemMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_OUTROS Then
            iBenef = BENEF_COMBO_OUTROS
        Else
            iBenef = BENEF_COMBO_DISP
        End If
        
        Call Combo_Seleciona_ItemData(Benef, iBenef)
        
        GridMovs.TextMatrix(iIndice, iGrid_Benef_Col) = Benef.Text

        'Mascara o ProdutoOP
        sProdutoOPMascarado = String(STRING_PRODUTO, 0)

        If Len(Trim(objItemMovEstoque.sProdutoOP)) > 0 Then

            lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProdutoOP, sProdutoOPMascarado)
            If lErro <> SUCESSO Then gError 34903
    
            Call StrParaMasked2(ProdutoOP, sProdutoOPMascarado)
'            ProdutoOP.PromptInclude = False
'            ProdutoOP.Text = sProdutoOPMascarado
'            ProdutoOP.PromptInclude = True
            
        Else
    
            sProdutoOPMascarado = ""
            
        End If
    
        GridMovs.TextMatrix(iIndice, iGrid_ProdutoOP_Col) = sProdutoOPMascarado

        'Mascara o Ccl
        If objItemMovEstoque.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemMovEstoque.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 34724

        Else

            sCclMascarado = ""

        End If

        GridMovs.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        Set colRatreamentoMovto = New Collection
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError 78422
        
        'Se existe rastreamento
        If colRatreamentoMovto.Count > 0 Then
                        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto(1)
            
            gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie 'Inserido por Wagner 13/03/2006
            
            If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovs.TextMatrix(iIndice, iGrid_Lote_Col) = objRatreamentoMovto.sLote
            
            If objRatreamentoMovto.iFilialOP > 0 Then
            
                objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP

                'Le a Filial Empresa da OP para pegar a descrição
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO Then gError 78791

                GridMovs.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
            
            End If
        
        '#####################################################
        'Inserido por Wagner 13/03/2006
        Else
            Set colRastreamentoSerie = New Collection
            gcolcolRastreamentoSerie.Add colRastreamentoSerie
        '#####################################################
        
        End If
        
        If objItemMovEstoque.lNumIntDocEst = 0 Then
            GridMovs.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"
        Else
            GridMovs.TextMatrix(iIndice, iGrid_Estorno_Col) = "1"
        End If

    Next


    objGrid.iLinhasExistentes = colItens.Count
    
    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 34725

    Preenche_GridMovs = SUCESSO

    Exit Function

Erro_Preenche_GridMovs:

    Preenche_GridMovs = gErr

    Select Case gErr

        Case 34724, 34725, 49718, 52055, 78791

        Case 34723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objItemMovEstoque.sProduto)

        Case 34903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objItemMovEstoque.sProdutoOP)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165391)

    End Select

    Exit Function

End Function

Public Sub CodigoLabel_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    If Len(Trim(Codigo.ClipText)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    'Janaina
    colSelecao.Add MOV_EST_REQ_PRODUCAO
    colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
    colSelecao.Add MOV_EST_REQ_PRODUCAO_OUTROS
    'Janaina
    
    'mario completar a colecao
    Call Chama_Tela("MovEstoqueLista1", colSelecao, objMovEstoque, objEventoCodigo)
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objMovEstoque As ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_objCodigoEvento_evSelecao

    Set objMovEstoque = obj1

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 34726

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objCodigoEvento_evSelecao:

    Select Case gErr

        Case 34726

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165392)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOP_OP_evSelecao(obj1 As Object)

Dim objItemOP As ClassItemOP

    Set objItemOP = obj1

    'Preeche OP com o codigo de objOrdemProducao
    OP.Text = objItemOP.sCodigo

    Me.Show

End Sub

Private Sub objEventoOP_OP_Todos_evSelecao(obj1 As Object)

Dim objOP As ClassOrdemDeProducao

    Set objOP = obj1

    'Preeche OP com o codigo de objOrdemProducao
    OP.Text = objOP.sCodigo

    Me.Show

End Sub

Private Sub objEventoProdutoOP_Grid_evSelecao(obj1 As Object)

Dim objItemOP As ClassItemOP
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoProdutoOP_OP_evSelecao

    Set objItemOP = obj1
    
    If GridMovs.Row > 0 Then
    
        lErro = CF("Produto_Formata", objItemOP.sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 55222

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            If (GridMovs.Col = iGrid_ProdutoOP_Col) And (iProdutoPreenchido = PRODUTO_PREENCHIDO) Then

                sProdutoMascarado = String(STRING_PRODUTO, 0)

                lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 34740

                GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = sProdutoMascarado

                Call StrParaMasked2(ProdutoOP, sProdutoMascarado)
'                ProdutoOP.PromptInclude = False
'                ProdutoOP.Text = sProdutoMascarado
'                ProdutoOP.PromptInclude = True
            
            End If
            
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_OP_evSelecao:

    Select Case gErr

        Case 34740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objItemOP.sProduto)

        Case 55222

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165393)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProdutoOP_OP_evSelecao(obj1 As Object)

Dim objItemOP As ClassItemOP
Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoProdutoOP_OP_evSelecao

    Set objItemOP = obj1

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 52097

    Call StrParaMasked2(ProdutoOPGera, sProdutoMascarado)
'    ProdutoOPGera.PromptInclude = False
'    ProdutoOPGera.Text = sProdutoMascarado
'    ProdutoOPGera.PromptInclude = True
    
    'Preenche o label de UM
    LblUM.Caption = objItemOP.sSiglaUM

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_OP_evSelecao:

    Select Case gErr

        Case 52096
        
        Case 52097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objItemOP.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165394)

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

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemProducao, objEventoOPPadrao)

End Sub

Private Sub objEventoOPPadrao_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOrdemDeProducao

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
    If lErro <> SUCESSO Then gError 34727

    'Preeche o CclPadrao com o codigo de objCcl
    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 34727
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165395)

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

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection
Dim objItemOP As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_OPLabel_Click
    
    lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52095

    'Se o Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        
        If Len(Trim(OP.Text)) <> 0 Then objItemOP.sCodigo = Trim(OP.Text)
                        
        Call Chama_Tela("ItemOP_ProdutoLista", colSelecao, objItemOP, objEventoOP_OP)
        
    Else
    
        If Len(Trim(OP.Text)) <> 0 Then objOrdemProducao.sCodigo = Trim(OP.Text)

        Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemProducao, objEventoOP_OP_Todos)
    
    End If
    
    Exit Sub
    
Erro_OPLabel_Click:
    
    Select Case gErr
    
        Case 52095
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165396)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOrdemDeProducao
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
    
On Error GoTo Erro_objEventoOP_evSelecao

    Set objOrdemProducao = obj1
    
    If (GridMovs.Row > 0) Then

        sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52201

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            '################################
            'Inserido por Wagner
            lErro = Valida_Repeticao(sCodProduto, objOrdemProducao.sCodigo)
            If lErro <> SUCESSO Then gError 132511
            '################################
            
            'Preenche a OP
            GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = objOrdemProducao.sCodigo
            OPCodigo.Text = objOrdemProducao.sCodigo
        
        End If
        
    End If

    Me.Show

    Exit Sub
    
Erro_objEventoOP_evSelecao:

    Select Case gErr
    
        Case 52201, 132511
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165397)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub ProdutoOPLabel_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim objItemOP As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoOPLabel_Click
    
    If Len(Trim(OP.Text)) > 0 Then
    
        lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 55223
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objItemOP.sProduto = sProdutoFormatado
            
        colSelecao.Add Trim(OP.Text)
            
        Call Chama_Tela("ItemOrdemProducao_OPLista", colSelecao, objItemOP, objEventoProdutoOP_OP)
            
    Else
        Error 52878
    End If
    
    
    Exit Sub
    
Erro_ProdutoOPLabel_Click:
    
    Select Case gErr
    
        Case 55223
        
        Case 52878
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREECHIDA", gErr)
            OP.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165398)

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
    If lErro <> SUCESSO Then gError 34728

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 34729

            Call StrParaMasked2(ProdutoOPGera, sProdutoMascarado)
'            ProdutoOPGera.PromptInclude = False
'            ProdutoOPGera.Text = sProdutoMascarado
'            ProdutoOPGera.PromptInclude = True
            
            'Preenche o label de UM
            LblUM.Caption = objProduto.sSiglaUMEstoque
            
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_evSelecao:

    Select Case gErr

        Case 34728
        
        Case 34729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165399)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoOP_Click()

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
On Error GoTo Erro_BotaoOP_Click
    
    If (GridMovs.Row = 0) Then gError 52205

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52206

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 52207

    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemProducao, objEventoOP)
   
    Exit Sub

Erro_BotaoOP_Click:

    Select Case gErr
        
        Case 52205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 52206
        
        Case 52207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165400)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOrdemDeProducao
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOP
Dim lErro As Long

On Error GoTo Erro_objEventoBotaoOP_evSelecao

    Set objOrdemProducao = obj1

    If (GridMovs.Row) > 0 Then

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34730

        'Se Produto desta Linha Estiver Preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = objOrdemProducao.sCodigo

        End If

    End If

    Exit Sub

Erro_objEventoBotaoOP_evSelecao:

    Select Case gErr

        Case 34730

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165401)

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
    
    If (GridMovs.Row = 0) Then gError 52209

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52210

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 52211

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
   
    Exit Sub

BotaoCcls_Click:

    Select Case gErr
        
        Case 52209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 52210
            
        Case 52211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165402)

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
            If lErro <> SUCESSO Then gError 34731

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

        Case 34731

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165403)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objItemOP As New ClassItemOP
    
On Error GoTo Erro_BotaoProdutos_Click

    If GridMovs.Row = 0 Then gError 55340

    'se a coluna escolhida for a do ProdutoOp then
    If GridMovs.Col = iGrid_ProdutoOP_Col Then
                
        ' Formata o ProdutoOp
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52250
        
        'se a OP estiver preenchida, mostra só os produtos da OP em questão
        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col))) <> 0 Then
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objItemOP.sProduto = sProdutoFormatado
            
            colSelecao.Add Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col))
            
            Call Chama_Tela("ItemOrdemProducao_OPLista", colSelecao, objItemOP, objEventoProdutoOP_Grid)
            
            Exit Sub
            
        Else
        
            gError 52830
        
        End If
        
    Else
        
        ' Formata o Produto
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52251
    
    End If
        
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado
        
   'Lista de produtos que  podem participar da producao e inventariados
   Call Chama_Tela("ProdutoEstoquePCPLista", colSelecao, objProduto, objEventoProduto)
   
   Exit Sub
   
Erro_BotaoProdutos_Click:

    Select Case gErr
        
        Case 52250, 52251
        
        Case 52830
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO_OP", gErr)
        
        Case 55340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165404)
        
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

    If GridMovs.Row = 0 Then gError 55343

    lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 34732

    If (GridMovs.Col = iGrid_Produto_Col) Then
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 55344

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 34733

        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 34734

        If lErro = 28030 Then gError 34735
        
        Call StrParaMasked2(Produto, sProdutoMascarado)
'        Produto.PromptInclude = False
'        Produto.Text = sProdutoMascarado
'        Produto.PromptInclude = True

        If Not (Me.ActiveControl Is Produto) Then

            GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = Produto.Text
    
            If Trim(objProduto.sContaContabilProducao) = "" Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 52212
                
                If lErro = 22531 Then gError 52213
                
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                
                'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
                If Trim(objProduto.sContaContabilProducao) = "" Then
                                    
                    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                    objCTBConfig.iFilialEmpresa = giFilialEmpresa
                            
                    lErro = CF("CTBConfig_Le", objCTBConfig)
                    If lErro <> SUCESSO And lErro <> 9755 Then gError 52820
                    
                    If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
                    
                End If
                
            End If
    
            'Preenche a Linha do Grid
            lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
            If lErro <> SUCESSO Then gError 34736
            
            lErro = Preenche_ContaContabilEst(GridMovs.Row, GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError 52231
            
            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then
                lErro = QuantDisponivel_Calcula1(sProdutoMascarado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError 34738
            Else
                lErro = QuantLote_Calcula1(sProdutoMascarado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)), objProduto)
                If lErro <> SUCESSO Then gError 78760
            End If
            
        End If

    Else

        iProdutoPreenchido = 0
        sProdutoMascarado = ""

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34739

        If (GridMovs.Col = iGrid_ProdutoOP_Col) Then

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 55345

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 34740

            GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = sProdutoMascarado

            Call StrParaMasked2(ProdutoOP, sProdutoMascarado)
'            ProdutoOP.PromptInclude = False
'            ProdutoOP.Text = sProdutoMascarado
'            ProdutoOP.PromptInclude = True

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 34732, 34734, 34736, 34738, 34739, 52212, 52231, 52820, 78760

        Case 34733, 34740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, objProduto.sCodigo)

        Case 34735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 52213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case 55343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 55344, 55345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, GridMovs.Row)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165405)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemMovEst As ClassItemMovEstoque) As Long

Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sContaEnxuta As String
Dim sAlmoxarifadoPadrao As String
Dim colRastreamentoSerie As New Collection 'Inserido por Wagner 13/03/2006

On Error GoTo Erro_ProdutoLinha_Preenche

    If Trim(Len(objProduto.sContaContabilProducao)) > 0 Then
        
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
        If lErro <> SUCESSO Then gError 55063
        
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
            If lErro <> SUCESSO Then gError 52230
    
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
    
        End If

    End If

    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 34741

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovs.TextMatrix(GridMovs.Row, iGrid_Ccl_Col) = CclPadrao.Text

    'OPCODIGO
    If Len(Trim(objProduto.sGrade)) = 0 And Len(Trim(OPCodigoPadrao.Text)) <> 0 Then
        GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = OPCodigoPadrao.Text
    ElseIf Len(Trim(objProduto.sGrade)) = 0 And Len(Trim(OP.Text)) <> 0 Then
        GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col) = OP.Text
    End If

    'Preenche Estorno com Valor 0 (Checked = False)
    GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) = "0"

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovs.Row - GridMovs.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolcolRastreamentoSerie.Add colRastreamentoSerie 'Inserido por Wagner 13/03/2006
        
        gobjMovEst.colItens.Add1 objItemMovEst
        
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridMovs.Row)
    
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

        Case 34741, 52230

        Case 55063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objProduto.sContaContabilProducao)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165406)

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

    If (GridMovs.Row = 0) Then gError 52207

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 34742

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 52208

    colSelecao.Add sProdutoFormatado

    Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 34742
        
        Case 52207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 52208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165407)

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
        If lErro <> SUCESSO Then gError 34743

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
            lErro = Preenche_ContaContabilEst(GridMovs.Row, GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError 55224

            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then
                'Calcula a Quantidade Disponível nesse Almoxarifado
                 lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
                If lErro <> SUCESSO Then gError 34744
            Else
                lErro = QuantLote_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
                If lErro <> SUCESSO Then gError 78761
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 34743, 55224, 78761

        Case 34744
            'Limpa a Quantidade Disponível da Tela e o Almoxarifado
            QuantDisponivel.Caption = ""
            Almoxarifado.Text = ""
            GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165408)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objMovEstoque Is Nothing Then

        objMovEstoque.iFilialEmpresa = giFilialEmpresa

        'Lê MovEstoque no Banco de Dados
        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 34744

        If lErro = SUCESSO Then 'Se ele existe

            If objMovEstoque.iTipoMov <> MOV_EST_REQ_PRODUCAO And objMovEstoque.iTipoMov <> MOV_EST_REQ_PRODUCAO_BENEF3 And objMovEstoque.iTipoMov <> MOV_EST_REQ_PRODUCAO_OUTROS Then gError 34745

            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError 34746

        ElseIf lErro = 30128 Then

            'Se ele não existe exibe apenas o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objMovEstoque.lCodigo)
            Codigo.PromptInclude = True

            lCodigoAntigo = objMovEstoque.lCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 34744, 34746

        Case 34745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOV_EST_REQ_NAO_PRODUCAO", gErr, objMovEstoque.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165409)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub OP_Validate(bMantemFoco As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As New ClassItemOP
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_OP_Validate

    'Verifica se foi Preenchido
    If Len(Trim(OP.Text)) > 0 Then

        objOrdemProducao.sCodigo = OP.Text
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa

        lErro = CF("OrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 34748

        If lErro = 30368 Then gError 34749

        'ordem de producao baixada
        If lErro = 55316 Then gError 55323
        
        lErro = CF2(Me, "ReqProd_OP_Validate", objOrdemProducao)
        If lErro <> SUCESSO Then gError 34750

        lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34750

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OP.Text
            objItemOP.sProduto = sProdutoFormatado

            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 34751

            If lErro = 34711 Then gError 34752

        End If

    End If

    Exit Sub

Erro_OP_Validate:

    bMantemFoco = True

    Select Case gErr

        Case 34748

        Case 34749
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OPCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then
                Call Chama_Tela("OrdemProducao", objOrdemProducao)
            End If

        Case 34750, 34751

        Case 34752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 55323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165410)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoOPGera_Validate(bMantemFoco As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemOP As New ClassItemOP
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_ProdutoOPGera_Validate

    If Len(Trim(ProdutoOPGera.ClipText)) > 0 Then

        sProduto = ProdutoOPGera.Text

        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 34753
        
        If lErro = 25041 Then gError 34754

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
           If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 34755
        End If
        
        LblUM.Caption = objProduto.sSiglaUMEstoque
        
        If Len(Trim(OP.Text)) > 0 Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OP.Text
            objItemOP.sProduto = objProduto.sCodigo

            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 34756
            If lErro = 34711 Then gError 34757

        End If

    End If

    Exit Sub

Erro_ProdutoOPGera_Validate:

    bMantemFoco = True
    
    Select Case gErr

        Case 34753, 34756

        Case 34754
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", sProduto)

             If vbMsg = vbYes Then
                objProduto.sCodigo = sProduto
                Call Chama_Tela("Produto", objProduto)
            End If
    
        Case 34755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr)

        Case 34757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165411)

    End Select

    Exit Sub

End Sub

Public Sub QuantidadeOP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_QuantidadeOP_Validate

    If Len(Trim(QuantidadeOP.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeOP.Text)
        If lErro <> SUCESSO Then gError 34851

    End If

    Exit Sub

Erro_QuantidadeOP_Validate:

    Cancel = True


    Select Case gErr

        Case 34851

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165412)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGeraReq_Click()

Dim lErro As Long
Dim dQuantidade As Double
Dim objOP As New ClassOrdemDeProducao
Dim colItensMovEstoque As New colItensMovEstoque
Dim objItemOP As New ClassItemOP
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
        
            For iIndice = objGrid.iLinhasExistentes To 1 Step -1
                Call gobjTelaProjetoInfo.Remove_Linha(iIndice)
            Next
        
            Call Grid_Limpa(objGrid)

            'Remove os ítens de gobjMovEst
            Set gobjMovEst.colItens = New colItensMovEstoque

        End If
        
    End If

    'se a ordem de produção não foi preenchida ==> erro
    If Len(Trim(OP.Text)) = 0 Then gError 34768

    'se a quantidade foi preenchida ==> guarda a quantidade
    If Len(Trim(QuantidadeOP.ClipText)) > 0 Then
        dQuantidade = CDbl(QuantidadeOP.Text)
    End If

    'se a quantidade for maior que 0 e o produto não foi preenchido ==> erro
    If (dQuantidade > 0) And (Len(Trim(ProdutoOPGera.ClipText)) = 0) Then gError 34769

    'se o produto estiver preenchido
    If Len(Trim(ProdutoOPGera.ClipText)) > 0 Then

        'coloca o produto no formato do BD
        lErro = CF("Produto_Formata", ProdutoOPGera.Text, sProdutoOPFormatado, iProdutoOPPreenchido)
        If lErro <> SUCESSO Then gError 34893

    End If

    'se o produto não foi preenchido
    If iProdutoOPPreenchido <> PRODUTO_PREENCHIDO Then

        objOP.sCodigo = OP.Text
        objOP.iFilialEmpresa = giFilialEmpresa

        'trata de gerar as requisições para todos os itens da OP
        lErro = OP_Requisicoes_Gera(objOP, colItensMovEstoque)
        If lErro <> SUCESSO Then gError 34770

    Else

        'se o produto foi preenchido
        objItemOP.sCodigo = OP.Text
        objItemOP.iFilialEmpresa = giFilialEmpresa
        objItemOP.sProduto = sProdutoOPFormatado
        If (dQuantidade > 0) Then objItemOP.dQuantidade = dQuantidade

        'gera as requisições para o produto em questão
        lErro = ItemOP_Requisicao_Gera(objItemOP, colItensMovEstoque)
        If lErro <> SUCESSO Then gError 34771

    End If

    'preenche o grid de itens requisitados com os itens dos kits que compõem a ordem de produção
    lErro = BotaoGeraReq_Preenche_Grid(colItensMovEstoque)
    If lErro <> SUCESSO Then gError 55064
    
    Exit Sub

Erro_BotaoGeraReq_Click:

    Select Case gErr

        Case 34768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREECHIDA", gErr)

        Case 34769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SEM_PRODUTO", gErr)

        Case 34770, 34771, 34893, 55064

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165413)

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
Dim colRastreamentoSerie As Collection 'Inserido por Wagner 13/03/2006
Dim objRatreamentoMovto As ClassRastreamentoMovto
Dim objFilialOP As AdmFiliais

On Error GoTo Erro_BotaoGeraReq_Preenche_Grid

    If colItensMovEstoque.Count > 0 Then
            
        lErro = CF2(Me, "ReqProd_BotaoGeraReq_Preenche_Grid", colItensMovEstoque)
        If lErro <> SUCESSO Then gError 34772

        For iIndice = 1 To colItensMovEstoque.Count

            sProdutoMascarado = String(STRING_PRODUTO, 0)
            sProdutoOPMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoTela(colItensMovEstoque.Item(iIndice).sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 34772

            Call StrParaMasked2(Produto, sProdutoMascarado)
'            Produto.PromptInclude = False
'            Produto.Text = sProdutoMascarado
'            Produto.PromptInclude = True

            '################################
            'Inserido por Wagner
            lErro = Valida_Repeticao(Produto.Text, OP.Text)
            If lErro <> SUCESSO Then gError 132514
            '################################

            iLinha = objGrid.iLinhasExistentes + 1

            GridMovs.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
            GridMovs.TextMatrix(iLinha, iGrid_Descricao_Col) = colItensMovEstoque.Item(iIndice).sProdutoDesc
            GridMovs.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = colItensMovEstoque.Item(iIndice).sSiglaUM
            GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(colItensMovEstoque.Item(iIndice).dQuantidade)
            
            If Len(Trim(colItensMovEstoque.Item(iIndice).sOPCodigo)) > 0 Then
                GridMovs.TextMatrix(iLinha, iGrid_OPCodigo_Col) = colItensMovEstoque.Item(iIndice).sOPCodigo
            Else
                GridMovs.TextMatrix(iLinha, iGrid_OPCodigo_Col) = OP.Text
            End If
            
            GridMovs.TextMatrix(iLinha, iGrid_Estorno_Col) = colItensMovEstoque.Item(iIndice).lNumIntDocEst

            objGrid.iLinhasExistentes = iLinha
            
            Call gobjTelaProjetoInfo.Adiciona_Linha(iLinha)

            If Trim(Len(colItensMovEstoque.Item(iIndice).sContaContabilAplic)) > 0 Then
        
                lErro = Mascara_RetornaContaEnxuta(colItensMovEstoque.Item(iIndice).sContaContabilAplic, sContaEnxuta)
                If lErro <> SUCESSO Then gError 52050
                
                'preenche  a ContaContabilProducao
                ContaContabilProducao.PromptInclude = False
                ContaContabilProducao.Text = sContaEnxuta
                ContaContabilProducao.PromptInclude = True
                
                'preenche Conta De Producao
                GridMovs.TextMatrix(iLinha, iGrid_ContaContabilProducao_Col) = ContaContabilProducao.Text
                    
            End If
            
            Call Combo_Seleciona_ItemData(Benef, colItensMovEstoque.Item(iIndice).iBenef)
            GridMovs.TextMatrix(iLinha, iGrid_Benef_Col) = Benef.Text

            If Len(Trim(colItensMovEstoque.Item(iIndice).sProdutoOP)) > 0 Then
                
                lErro = Mascara_RetornaProdutoTela(colItensMovEstoque.Item(iIndice).sProdutoOP, sProdutoOPMascarado)
                If lErro <> SUCESSO Then gError 34890
    
                GridMovs.TextMatrix(iLinha, iGrid_ProdutoOP_Col) = sProdutoOPMascarado

            End If

            If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then
                GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
            Else

                'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
                lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", colItensMovEstoque.Item(iIndice).sProduto, sAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 52219 And lErro <> 52221 Then gError 55058

                If lErro = SUCESSO Then GridMovs.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
            
            End If

            'coloca o Ccl no formato do bd
            lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
            If lErro <> SUCESSO Then gError 34773

            If iCclPreenchida = CCL_PREENCHIDA Then GridMovs.TextMatrix(iLinha, iGrid_Ccl_Col) = CclPadrao.Text

            Set objItemMovEst = New ClassItemMovEstoque
            gobjMovEst.colItens.Add1 objItemMovEst

            'preenche a conta contabil de estoque, se houver alguma associada ao produto/almoxarifado em questão
            lErro = Preenche_ContaContabilEst(iLinha, GridMovs.TextMatrix(iLinha, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError 55065
            
            'Se existe rastreamento
            If Not (colItensMovEstoque.Item(iIndice).colRastreamentoMovto Is Nothing) Then
                If colItensMovEstoque.Item(iIndice).colRastreamentoMovto.Count > 0 Then
                    'Seta o primeiro Lote
                    Set objRatreamentoMovto = colItensMovEstoque.Item(iIndice).colRastreamentoMovto(1)
                    
                    gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie
                    
                    If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovs.TextMatrix(iIndice, iGrid_Lote_Col) = objRatreamentoMovto.sLote
                    
                    If objRatreamentoMovto.iFilialOP > 0 Then
                    
                        Set objFilialOP = New AdmFiliais
                    
                        objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP
        
                        'Le a Filial Empresa da OP para pegar a descrição
                        lErro = CF("FilialEmpresa_Le", objFilialOP)
                        If lErro <> SUCESSO Then gError 55065
        
                        GridMovs.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
                    
                    End If
                
                Else
                    Set colRastreamentoSerie = New Collection
                    gcolcolRastreamentoSerie.Add colRastreamentoSerie
                End If
            Else
                Set colRastreamentoSerie = New Collection
                gcolcolRastreamentoSerie.Add colRastreamentoSerie
            End If
            
'            '#################################################
'            'Inserido por Wagner 13/03/2006
'            Set colRastreamentoSerie = New Collection
'
'            gcolcolRastreamentoSerie.Add colRastreamentoSerie
'            '#################################################

        Next

        lErro = Grid_Refresh_Checkbox(objGrid)
        If lErro <> SUCESSO Then gError 34774

    End If

    BotaoGeraReq_Preenche_Grid = SUCESSO

    Exit Function

Erro_BotaoGeraReq_Preenche_Grid:

    BotaoGeraReq_Preenche_Grid = gErr

    Select Case gErr

        Case 34772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, colItensMovEstoque.Item(iIndice).sProduto)
        
        Case 34773, 34774, 55058, 55065, 132514 'Inserido por Wagner
        
        Case 34890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_Mascara_RetornaProdutoTela", gErr, colItensMovEstoque.Item(iIndice).sProdutoOP)
        
        Case 52050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, colItensMovEstoque.Item(iIndice).sContaContabilAplic)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165414)
        
    End Select
    
    Exit Function
    
End Function

Private Function OP_Requisicoes_Gera(objOP As ClassOrdemDeProducao, colItensMovEstoque As colItensMovEstoque) As Long
'Gera Requisições para todos os Produtos da OP a partir das quantidades que faltam produzir

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objItemOPGrade As ClassItemOP
Dim objProduto As New ClassProduto
Dim colItensOPGrade As New Collection

On Error GoTo Erro_OP_Requisicoes_Gera

    lErro = CF("ItensOrdemProducao_Le", objOP)
    If lErro <> SUCESSO And lErro <> 30401 Then gError 34863
    
    'se a ordem de producao nao tiver itens associados ==> erro
    If lErro = 30401 Then gError 34864

    For Each objItemOP In objOP.colItens
    
        objItemOP.sCodigo = objOP.sCodigo
        objItemOP.iFilialEmpresa = objOP.iFilialEmpresa

        objProduto.sCodigo = objItemOP.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126575

        'Ausência de Produto no BD
        If lErro = 28030 Then gError 126576

        If Len(Trim(objProduto.sGrade)) > 0 Then

            lErro = CF("ItensOPGrade_Le", colItensOPGrade, objItemOP.lNumIntDoc)
            If lErro <> SUCESSO Then gError 126577
    
            For Each objItemOPGrade In colItensOPGrade
            
                If objItemOPGrade.dQuantidade > objItemOPGrade.dQuantidadeProd And (objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL Or objItemOP.iSituacao = ITEMOP_SITUACAO_SACR) Then
        
                    objItemOPGrade.dQuantidade = objItemOPGrade.dQuantidade - objItemOPGrade.dQuantidadeProd
                    objItemOPGrade.sSiglaUM = objItemOP.sSiglaUM
            
                    lErro = OP_Requisicoes_Gera1(objItemOPGrade, colItensMovEstoque)
                    If lErro <> SUCESSO Then gError 126578
                    
                End If
            
            Next

        Else
        
            If objItemOP.dQuantidade > objItemOP.dQuantidadeProd And (objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL Or objItemOP.iSituacao = ITEMOP_SITUACAO_SACR) Then
    
                objItemOP.dQuantidade = objItemOP.dQuantidade - objItemOP.dQuantidadeProd
        
                lErro = OP_Requisicoes_Gera1(objItemOP, colItensMovEstoque)
                If lErro <> SUCESSO Then gError 126579
                
            End If
        
        End If
        
    Next

    OP_Requisicoes_Gera = SUCESSO

    Exit Function

Erro_OP_Requisicoes_Gera:

    OP_Requisicoes_Gera = gErr

    Select Case gErr

        Case 34863, 126575, 126577, 126578, 126579
        
        Case 34864
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMPRODUCAO_SEM_ITENS", gErr, objOP.sCodigo, objOP.iFilialEmpresa)

        Case 126576
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165415)

    End Select

    Exit Function

End Function

Private Function OP_Requisicoes_Gera1(ByVal objItemOP As ClassItemOP, ByVal colItensMovEstoque As colItensMovEstoque) As Long

Dim lErro As Long
Dim objKit As New ClassKit
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos
Dim objProdutoKitRaiz As New ClassProdutoKit
Dim dQuantidadeConv As Double, dFator As Double, dQtdeInsumoParaProducaoTotal As Double
Dim iClasseUM As Integer, iIndice As Integer
Dim objProduto As New ClassProduto, dQuantidadeArr As Double
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objItemMovEst As ClassItemMovEstoque
Dim bPulaReq As Boolean, objEmpenho As ClassEmpenho
Dim objItemOPBD As New ClassItemOP

On Error GoTo Erro_OP_Requisicoes_Gera1

    objItemOPBD.lNumIntDoc = objItemOP.lNumIntDoc
    objItemOPBD.sCodigo = objItemOP.sCodigo
    objItemOPBD.sProduto = objItemOP.sProduto
    objItemOPBD.iFilialEmpresa = objItemOP.iFilialEmpresa
    lErro = CF("ItemOP_Le", objItemOPBD)
    If lErro <> SUCESSO And lErro <> 34711 Then gError ERRO_SEM_MENSAGEM

    Set objKit = New ClassKit
    
    objKit.sProdutoRaiz = objItemOP.sProduto
    objKit.sVersao = objItemOP.sVersao

    lErro = CF("Kit_Le", objKit)
    If lErro <> SUCESSO And lErro <> 21826 Then gError 34865
    
    'se não for encontrado um kit para o produto da ordem de producao ==> erro
    If lErro = 21826 Then gError 34866

    objProdutoKitRaiz.sProdutoRaiz = objKit.sProdutoRaiz
    objProdutoKitRaiz.sVersao = objKit.sVersao

    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKitRaiz)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 34867
    
    If lErro = 34875 Then gError 34868

    'Alteracao Daniel em 29/07/2002
    objProdutoKitProdutos.sProdutoRaiz = objItemOP.sProduto
    objProdutoKitProdutos.sVersao = objItemOP.sVersao
    
    'Se a quantidade estiver preenchida => Utiliza para base de calculo da quantidade da Embalagem
    If Len(Trim(QuantidadeOP.Text)) > 0 Then
        objProdutoKitProdutos.dQuantidade = StrParaDbl(QuantidadeOP.Text)
    Else
        objProdutoKitProdutos.dQuantidade = objItemOP.dQuantidade
    End If
    
    objProdutoKitProdutos.iClasseUM = objProdutoKitRaiz.iClasseUM
    objProdutoKitProdutos.sUnidadeMed = objItemOP.sSiglaUM
    
    'Obtem uma Colecao com os itens do kit para empenho (a embalagem do ProdutoRaiz será o ultimo elemento da colecao)
    lErro = CF("OrdemProducao_Le_Col_Empenho", objProdutoKitProdutos, objKit.colComponentes)
    If lErro <> SUCESSO And lErro <> 106393 Then gError 106411
    
    'Se nao Encontrou => Erro
    If lErro = 106393 Then gError 106412

    For Each objProdutoKitProdutos In objKit.colComponentes

        If objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_FIXA Then

            Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, objProdutoKitProdutos.dQuantidade, dQuantidadeArr)
            dQuantidadeConv = dQuantidadeArr
            dQtdeInsumoParaProducaoTotal = dQuantidadeArr

        ElseIf objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then

            lErro = CF("UM_Conversao_Trans", objProdutoKitRaiz.iClasseUM, objItemOP.sSiglaUM, objProdutoKitRaiz.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 34871

            Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, ((objItemOP.dQuantidade * dFator) / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade), dQuantidadeArr)
            dQuantidadeConv = dQuantidadeArr

            'Se o que falta produzir é menor que o valor informado considera que a produção total é o Produzido + Informado
            'Senão é a Quantidade do item da OP
            If objItemOPBD.dQuantidade - objItemOPBD.dQuantidadeProd - objItemOP.dQuantidade < -QTDE_ESTOQUE_DELTA Then
                Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, ((objItemOPBD.dQuantidadeProd + objItemOP.dQuantidade) * dFator / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade), dQuantidadeArr)
            Else
                Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, (objItemOPBD.dQuantidade * dFator / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade), dQuantidadeArr)
            End If
            dQtdeInsumoParaProducaoTotal = dQuantidadeArr

        End If

        'Carrega sProdutoFormatado em objProduto
        objProduto.sCodigo = objProdutoKitProdutos.sProduto

        'Tenta ler o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 55059

        'Ausência de Produto no BD
        If lErro = 28030 Then gError 55060

        If Trim(objProduto.sContaContabilProducao) = "" Then
            
            objTipoDeProduto.iTipo = objProduto.iTipo
        
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 55061
        
            If lErro = 22531 Then gError 55062
        
            objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                    
        End If

        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
        
            bPulaReq = False
            
            If IgnorarReqProdAnt.Value = vbUnchecked Then
            
                Set objEmpenho = New ClassEmpenho
                With objEmpenho
                    .iFilialEmpresa = objItemOP.iFilialEmpresa
                    .lNumIntDocItemOP = objItemOP.lNumIntDoc
                    .sProduto = objProdutoKitProdutos.sProduto
                End With
                
                lErro = CF("Empenho_Le_1", objEmpenho)
                If lErro <> SUCESSO And lErro <> 178592 Then gError ERRO_SEM_MENSAGEM
        
                If lErro = SUCESSO Then
                    If Abs(objItemOPBD.dQuantidade - objItemOPBD.dQuantidadeProd - objItemOP.dQuantidade) < QTDE_ESTOQUE_DELTA Then
                        'REQUISIÇÃO TOTAL DO QUE FALTA PRODUZIR
                        'Está requisitando o que falta produzir -> requisita o que falta requisitar para o todo
                        dQuantidadeConv = dQtdeInsumoParaProducaoTotal - objEmpenho.dQuantidadeRequisitada
                    Else
                        'REQUISIÇÃO PARCIAL DO QUE FALTA PRODUZIR
                        'Se a quantidade do insumo é maior do que falta requisitar para O TODO -> Reduz Senão -> Mantém
                        'O TODO é igual ao maior de (Quantidade Total do Item da OP) e (Quantidade Total Produzida + Quantidade Informada na Tela)
                        If dQuantidadeConv - (dQtdeInsumoParaProducaoTotal - objEmpenho.dQuantidadeRequisitada) >= QTDE_ESTOQUE_DELTA Then
                            dQuantidadeConv = dQtdeInsumoParaProducaoTotal - objEmpenho.dQuantidadeRequisitada
                        Else
                            'Respeita a necessidade para quantidade informada
                        End If
                    End If
    
                    Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, dQuantidadeConv, dQuantidadeArr)
                    dQuantidadeConv = dQuantidadeArr
                    
                    If dQuantidadeConv < QTDE_ESTOQUE_DELTA Then bPulaReq = True
                End If
                
            End If
                
            If bPulaReq = False Then
            
                Set objItemMovEst = colItensMovEstoque.Add(0, 0, 0, 0, objProdutoKitProdutos.sProduto, objProduto.sDescricao, objProdutoKitProdutos.sUnidadeMed, dQuantidadeConv, 0, "", 0, "", 0, objItemOP.sCodigo, objItemOP.sProduto, objProduto.sContaContabilProducao, "", 0, Nothing, Nothing, DATA_NULA)
                objItemMovEst.iBenef = objItemOP.iBeneficiamento
            End If

        End If

    Next

    OP_Requisicoes_Gera1 = SUCESSO

    Exit Function

Erro_OP_Requisicoes_Gera1:

    OP_Requisicoes_Gera1 = gErr

    Select Case gErr

        Case 34865, 34867, 34871, 55059, 55061, 106411, ERRO_SEM_MENSAGEM
        
        Case 34866
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_NAO_E_KIT", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 34868
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRODUTO_RAIZ", gErr, objKit.sProdutoRaiz)

        Case 55060
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 55062
             Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
        
        Case 106412
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, objKit.sProdutoRaiz)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165416)

    End Select

    Exit Function
    
End Function

Public Function ItemOP_Requisicao_Gera(objItemOP As ClassItemOP, colItensMovEstoque As colItensMovEstoque) As Long
'Gera Requisições para o Produto da OP a partir da quantidade que falta produzir

Dim lErro As Long
Dim dQuantidadePedida As Double, dQuantidadeConv As Double, dFator As Double, dQtdeInsumoParaProducaoTotal As Double
Dim vbMsg As VbMsgBoxResult
Dim objKit As New ClassKit
Dim objProdutoKitRaiz As New ClassProdutoKit
Dim iClasseUM As Integer
Dim objProduto As New ClassProduto, dQuantidadeArr As Double
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos
Dim objItemMovEst As ClassItemMovEstoque
Dim bPulaReq As Boolean, objEmpenho As ClassEmpenho

On Error GoTo Erro_ItemOP_Requisicao_Gera

    dQuantidadePedida = objItemOP.dQuantidade 'Quantidade Informada pelo Usuário na Tela

    lErro = CF("ItemOP_Le", objItemOP)
    If lErro <> SUCESSO And lErro <> 34711 Then gError 34852
    
    If lErro = 34711 Then gError 34853

    'Se o item da OP já teve a produção completa dá erro
    If objItemOP.dQuantidade = objItemOP.dQuantidadeProd Then gError 34854

    'Se ainda falta algo a produzir
    If objItemOP.dQuantidade > objItemOP.dQuantidadeProd Then
        
        If dQuantidadePedida = 0 Then
            'Se não foi informado a quantidade ela será o que falta para produzir
            dQuantidadePedida = objItemOP.dQuantidade - objItemOP.dQuantidadeProd

        ElseIf objItemOP.dQuantidade > 0 Then

            If objItemOP.dQuantidade - objItemOP.dQuantidadeProd < dQuantidadePedida Then
            
                'Se for informado quantidade e ela é maior do que falta produzir dá um aviso
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_PEDIDA_A_MAIOR", dQuantidadePedida, objItemOP.sProduto, (objItemOP.dQuantidade - objItemOP.dQuantidadeProd), OP.Text)

                'Para seguir o usuário tem que aceitar requisitar apenas para o que falta
                If vbMsg <> vbNo Then dQuantidadePedida = objItemOP.dQuantidade - objItemOP.dQuantidadeProd

            End If

        End If

    End If

    objKit.sProdutoRaiz = objItemOP.sProduto
    objKit.sVersao = objItemOP.sVersao

    'Le o Kit / Versao
    lErro = CF("Kit_Le", objKit)
    If lErro <> SUCESSO And lErro <> 21826 Then gError 106422
    
    'Se nao Encontrou => Erro
    If lErro = 21826 Then gError 106423
    
    objProdutoKitRaiz.sProdutoRaiz = objKit.sProdutoRaiz
    objProdutoKitRaiz.sVersao = objKit.sVersao

    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKitRaiz)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 34858
    If lErro = 34875 Then gError 34859
    
    'Alteracao Daniel em 29/07/2002
    objProdutoKitProdutos.sProdutoRaiz = objItemOP.sProduto
    objProdutoKitProdutos.sVersao = objItemOP.sVersao
        
    'Se a quantidade estiver preenchida => Utiliza para base de calculo da quantidade da Embalagem
    If Len(Trim(QuantidadeOP.Text)) > 0 Then
        objProdutoKitProdutos.dQuantidade = StrParaDbl(QuantidadeOP.Text)
    Else
        objProdutoKitProdutos.dQuantidade = objItemOP.dQuantidade
    End If

    objProdutoKitProdutos.iClasseUM = objProdutoKitRaiz.iClasseUM
    
    'Alterado por Shirley 11/09/2002
    objProdutoKitProdutos.sUnidadeMed = objItemOP.sSiglaUM
    
    'Obtem uma Colecao com os itens do kit para empenho (a embalagem do ProdutoRaiz será o ultimo elemento da colecao)
    lErro = CF("OrdemProducao_Le_Col_Empenho", objProdutoKitProdutos, objKit.colComponentes)
    If lErro <> SUCESSO And lErro <> 106393 Then gError 106413
    
    'Se nao Encontrou => Erro
    If lErro = 106393 Then gError 106414
    
    For Each objProdutoKitProdutos In objKit.colComponentes

        If objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_FIXA Then

            Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, objProdutoKitProdutos.dQuantidade, dQuantidadeArr)
            dQuantidadeConv = dQuantidadeArr
            dQtdeInsumoParaProducaoTotal = dQuantidadeArr

        ElseIf objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then

            lErro = CF("UM_Conversao_Trans", objProdutoKitRaiz.iClasseUM, objItemOP.sSiglaUM, objProdutoKitRaiz.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 34862

            Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, (dQuantidadePedida * dFator / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade), dQuantidadeArr)
            dQuantidadeConv = dQuantidadeArr

            'Se o que falta produzir é menor que o valor informado considera que a produção total é o Produzido + Informado
            'Senão é a Quantidade do item da OP
            If objItemOP.dQuantidade - objItemOP.dQuantidadeProd - dQuantidadePedida < -QTDE_ESTOQUE_DELTA Then
                Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, ((objItemOP.dQuantidadeProd + dQuantidadePedida) * dFator / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade), dQuantidadeArr)
            Else
                Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, (objItemOP.dQuantidade * dFator / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade), dQuantidadeArr)
            End If
            dQtdeInsumoParaProducaoTotal = dQuantidadeArr

        End If

        'Carrega sProdutoFormatado em objProduto
        objProduto.sCodigo = objProdutoKitProdutos.sProduto

        'Tenta ler o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 55066

        'Ausência de Produto no BD
        If lErro = 28030 Then gError 55067

        If Trim(objProduto.sContaContabilProducao) = "" Then
            
            objTipoDeProduto.iTipo = objProduto.iTipo
        
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 55068
        
            If lErro = 22531 Then gError 55069
        
            objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                    
        End If

        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
        
            bPulaReq = False
            
            If IgnorarReqProdAnt.Value = vbUnchecked Then
                
                Set objEmpenho = New ClassEmpenho
                With objEmpenho
                    .iFilialEmpresa = objItemOP.iFilialEmpresa
                    .lNumIntDocItemOP = objItemOP.lNumIntDoc
                    .sProduto = objProdutoKitProdutos.sProduto
                End With
                
                'Lê o empenho para abater o que já foi requisitado
                lErro = CF("Empenho_Le_1", objEmpenho)
                If lErro <> SUCESSO And lErro <> 178592 Then gError ERRO_SEM_MENSAGEM
                    
                If lErro = SUCESSO Then
                    If Abs(objItemOP.dQuantidade - objItemOP.dQuantidadeProd - dQuantidadePedida) < QTDE_ESTOQUE_DELTA Then
                        'REQUISIÇÃO TOTAL DO QUE FALTA PRODUZIR
                        'Está requisitando o que falta produzir -> requisita o que falta requisitar para o todo
                        dQuantidadeConv = dQtdeInsumoParaProducaoTotal - objEmpenho.dQuantidadeRequisitada
                    Else
                        'REQUISIÇÃO PARCIAL DO QUE FALTA PRODUZIR
                        'Se a quantidade do insumo é maior do que falta requisitar para O TODO -> Reduz Senão -> Mantém
                        'O TODO é igual ao maior de (Quantidade Total do Item da OP) e (Quantidade Total Produzida + Quantidade Informada na Tela)
                        If dQuantidadeConv - (dQtdeInsumoParaProducaoTotal - objEmpenho.dQuantidadeRequisitada) >= QTDE_ESTOQUE_DELTA Then
                            dQuantidadeConv = dQtdeInsumoParaProducaoTotal - objEmpenho.dQuantidadeRequisitada
                        Else
                            'Respeita a necessidade para quantidade informada
                        End If
                    End If
    
                    Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, dQuantidadeConv, dQuantidadeArr)
                    dQuantidadeConv = dQuantidadeArr
                    
                    If dQuantidadeConv < QTDE_ESTOQUE_DELTA Then bPulaReq = True
                End If
            End If
            
            If bPulaReq = False Then
                Set objItemMovEst = colItensMovEstoque.Add(0, 0, 0, 0, objProdutoKitProdutos.sProduto, objProduto.sDescricao, objProdutoKitProdutos.sUnidadeMed, dQuantidadeConv, 0, "", 0, "", 0, objItemOP.sCodigo, objItemOP.sProduto, objProduto.sContaContabilProducao, "", 0, Nothing, Nothing, DATA_NULA)
                objItemMovEst.iBenef = objItemOP.iBeneficiamento
            End If
            
        End If

    Next

    ItemOP_Requisicao_Gera = SUCESSO

    Exit Function

Erro_ItemOP_Requisicao_Gera:

    ItemOP_Requisicao_Gera = gErr

    Select Case gErr

        Case 34852, 34856, 34858, 34862, 34860, 55066, 55068, 106413, 106422, ERRO_SEM_MENSAGEM
        
        Case 106414
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, objKit.sProdutoRaiz)
        
        Case 106423
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_INEXISTENTE", Err, objKit.sVersao, objKit.sProdutoRaiz)

        Case 34855
            QuantidadeOP.SetFocus

        Case 34853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 34854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_PRODUZIDO", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 34857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_NAO_E_KIT", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 34859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRODUTO_RAIZ", gErr, objProdutoKitProdutos.sProduto)

        Case 34861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, objProdutoKitProdutos.sProduto)

        Case 55067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 55069
             lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165417)

    End Select

    Exit Function

End Function

Public Sub OPCodigoPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_OPCodigoPadrao_Validate

    If Len(Trim(OPCodigoPadrao.Text)) > 0 Then

        objOrdemProducao.sCodigo = OPCodigoPadrao.Text
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa

        lErro = CF("OrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 34758
        
        If lErro = 30368 Then gError 34759
        
        'ordem de producao baixada
        If lErro = 55316 Then gError 55317

        If Len(Trim(OP.Text)) = 0 Then OP.Text = OPCodigoPadrao.Text

    End If

    Exit Sub

Erro_OPCodigoPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 34758

        Case 34759
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)

        Case 55317
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165418)

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
        If lErro <> SUCESSO And lErro <> 5703 Then gError 34760
        
        If lErro = 5703 Then gError 34761

    End If

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 34760

        Case 34761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165419)

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
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 34762

        If lErro = 25136 Then gError 34763

        If lErro = 25143 Then gError 34764

    End If

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 34762

        Case 34763, 34764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165420)

    End Select

    Exit Sub

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 34765

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 34765

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165421)

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
    If lErro <> SUCESSO Then gError 89806

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 89806

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165422)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 34766

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 34766

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165423)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 34767

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 34767

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165424)

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
        sCodProduto = GridMovs.TextMatrix(iLinha, iGrid_Produto_Col)
    
        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34775
    
        If gobjMovEst.colItens.Count >= GridMovs.Row Then
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
                
                    objProduto.sCodigo = sProdutoFormatado
        
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 34777
        
                    If lErro = 28030 Then gError 34778
        
                    objClasseUM.iClasse = objProduto.iClasseUM
        
                    'Preenche a List da Combo UnidadeMed com as UM's do Produto
                    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                    If lErro <> SUCESSO Then gError 34779
        
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
                
                    objProduto.sCodigo = sProdutoFormatado
            
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 78215
            
                    If lErro = 28030 Then gError 78216
                
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
        
                    objProduto.sCodigo = sProdutoFormatado
        
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 78773
        
                    If lErro = 28030 Then gError 78774
        
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
        
        End Select

        If gobjCRFAT.iUsaBloqAcessoPorTelaControle = MARCADO Then
            lErro = CF("Rotina_Grid_Enable_BloqueiaAcesso", Me.Name, objControl)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

    End If
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 34775, 34777, 34778, 34779, 78215, 78216, 78323, 78324, 78773, 78774
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165425)

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
Dim iLinhasExistentes As Integer 'm
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

            gcolcolRastreamentoSerie.Remove iLinhaAnterior 'Inserido por Wagner 13/03/2006

            Call gobjTelaProjetoInfo.Remove_Linha(iLinhaAnterior)

            For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
                If gobjMovEst.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
                    GridMovs.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
                Else
                    GridMovs.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
                End If
                
            Next

            GridMovs.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

            lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 126664

        End If

    End If

    Exit Sub
    
Erro_GridMovs_KeyDown:

    Select Case gErr
    
        Case 126664
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165426)
    
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
    
        Call CF("PreencheProduto_Customizado", Produto, sProdPreenchido)

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 34781

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError 126490
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 34782

            If Trim(objProduto.sContaContabilProducao) = "" Then
            
                objTipoDeProduto.iTipo = objProduto.iTipo
        
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 52214
        
                If lErro = 22531 Then gError 52215
        
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
                                
                'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
                If Trim(objProduto.sContaContabilProducao) = "" Then
                            
                    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                    objCTBConfig.iFilialEmpresa = giFilialEmpresa
                    
                    lErro = CF("CTBConfig_Le", objCTBConfig)
                    If lErro <> SUCESSO And lErro <> 9755 Then gError 52823
            
                    If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
            
                End If
                                
            End If

            'se nao for um produto pai de grade
            If Len(Trim(objProduto.sGrade)) = 0 Then

                If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 34783

                lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
                If lErro <> SUCESSO Then gError 34784
                
                lErro = Preenche_ContaContabilEst(GridMovs.Row, Produto.Text)
                If lErro <> SUCESSO Then gError 52232
        
                If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then
                    lErro = QuantDisponivel_Calcula1(Produto.Text, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), objProduto)
                    If lErro <> SUCESSO Then gError 55219
                Else
                    lErro = QuantLote_Calcula1(Produto.Text, GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)), objProduto)
                    If lErro <> SUCESSO Then gError 78762
                End If
        
                If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                    GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
                End If

            'se for um produto pai de grade
            Else
            
                'Verifica se há filhos válidos da grade pai
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 126588
                
                'Se nao existir, erro
                If colItensRomaneioGrade.Count = 0 Then gError 126589
                
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
                If giRetornoTela <> vbOK Then gError 126590
        
                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
                If lErro <> SUCESSO Then gError 126591
        
                lErro = Preenche_ContaContabilEst(GridMovs.Row, Produto.Text)
                If lErro <> SUCESSO Then gError 126592
        
                Call Atualiza_Grid_Movimentos(objItemMovEst)
            
            End If

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34785

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 34781, 34784, 34785, 52214, 52232, 52823, 78762, 126588, 126590, 126591, 126592
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 34782
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 34783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52215
             lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126589
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165427)

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
        If lErro <> SUCESSO Then gError 34786

        If gobjMovEst.colItens(GridMovs.Row).lNumIntDoc = 0 Then

            If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then

                lErro = Testa_QuantRequisitada(StrParaDbl(Quantidade.Text))
                If lErro <> SUCESSO Then gError 34909

            End If

        End If

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    '#################################################################
    'Inserido por Wagner 13/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(Quantidade.Text), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
    If lErro <> SUCESSO Then gError 141911
    '#################################################################

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34787

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 34786, 34787, 34909, 141911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165428)
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
Dim objProduto As New ClassProduto 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34788

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 34789

        If lErro = 25157 Then gError 34790

        If lErro = 25162 Then gError 34791

        GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141948
        '###########################################################

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 34901
        Else
            lErro = QuantLote_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 78764
        End If
    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34792
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    Call Preenche_ContaContabilEst(GridMovs.Row, GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))
            
    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 34788, 34789, 34792, 34901, 141948
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 34790

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 34791

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165429)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_OPCodigo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String, sProdutoOPEnxuto As String
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOP, objItemOPUnico As ClassItemOP
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim iCont As Integer

On Error GoTo Erro_Saida_Celula_OPCodigo

    Set objGridInt.objControle = OPCodigo

    '################################
    'Inserido por Wagner
    lErro = Valida_Repeticao(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), OPCodigo.Text)
    If lErro <> SUCESSO Then gError 132510
    '################################
    
    If Len(Trim(OPCodigo.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = OPCodigo.Text

        lErro = CF("OrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 34793

        If lErro = 30368 Then gError 34794

        'ordem de producao baixada
        If lErro = 55316 Then gError 55318

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34795

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = OPCodigo.Text
            objItemOP.sProduto = sProdutoFormatado

            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 34796

            If lErro = 34711 Then gError 34797
            
        Else
        
            lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 30401 Then gError 90001
            
            For Each objItemOP In objOrdemProducao.colItens
                If objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
                    If objItemOPUnico Is Nothing Then Set objItemOPUnico = objItemOP
                    iCont = iCont + 1
                End If
            Next
            
            If iCont = 1 Then
                'Mascara o ProdutoOP
                lErro = Mascara_RetornaProdutoEnxuto(objItemOPUnico.sProduto, sProdutoOPEnxuto)
                If lErro <> SUCESSO Then gError 34903
        
                ProdutoOP.PromptInclude = False
                ProdutoOP.Text = sProdutoOPEnxuto
                ProdutoOP.PromptInclude = True
        
                GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = ProdutoOP.Text
            End If
        End If
          
    Else
        
        ProdutoOP.PromptInclude = False
        ProdutoOP.Text = ""
        ProdutoOP.PromptInclude = True

        GridMovs.TextMatrix(GridMovs.Row, iGrid_ProdutoOP_Col) = ""

    End If

    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34798

    Saida_Celula_OPCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_OPCodigo:

    Saida_Celula_OPCodigo = gErr

    Select Case gErr

        Case 34793, 34795, 34796, 34798, 78506, 78510, 78511, 90001, 132510 'Inserido por Wagner
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 34794
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OPCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("OrdemProducao", objOrdemProducao)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 34797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78507
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsg = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 78512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165430)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoOP(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objItemOP As New ClassItemOP
Dim objKit As New ClassKit
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iAchou As Integer, iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ProdutoOP

    Set objGridInt.objControle = ProdutoOP
    iAchou = 0

    If Len(Trim(ProdutoOP.ClipText)) > 0 Then

        lErro = CF("Produto_Critica", ProdutoOP.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 34799

        If lErro = 25041 Then gError 34800

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col))) > 0 Then

            objItemOP.sCodigo = GridMovs.TextMatrix(GridMovs.Row, iGrid_OPCodigo_Col)
            objItemOP.iFilialEmpresa = giFilialEmpresa

            objItemOP.sProduto = objProduto.sCodigo

            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 34808
            
            If lErro = 34711 Then gError 34809

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34889

    Saida_Celula_ProdutoOP = SUCESSO

    Exit Function


Erro_Saida_Celula_ProdutoOP:

    Saida_Celula_ProdutoOP = gErr

    Select Case gErr

        Case 34799, 34808, 34889
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 34800
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoOP.Text)
            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdutoOP.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 34809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165431)


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
        If lErro <> SUCESSO And lErro <> 5703 Then gError 34810

        If lErro = 5703 Then gError 34811

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34812

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 34810, 34812
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 34811
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165432)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

'#################################
'Inserido por Wagner 15/03/2006
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
'#################################

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    If Len(UnidadeMed.Text) > 0 Then

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141949
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141950
        '###########################################################

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 55220
        Else
            lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78765
        End If
        
        If gobjMovEst.colItens(GridMovs.Row).lNumIntDoc = 0 Then
    
            'Se a quantidade está preenchida e não se trata de estorno
            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then
    
                dQuantidade = CDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col))
    
                'Testa a Quantidade requisitada
                lErro = Testa_QuantRequisitada(dQuantidade)
                If lErro <> SUCESSO Then gError 55213
    
            End If
    
        End If

    Else
    
        QuantDisponivel.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34837

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 34837, 55213, 55220
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165433)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Estorno(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Estorno

    Set objGridInt.objControle = Estorno

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34838

    Saida_Celula_Estorno = SUCESSO

    Exit Function

Erro_Saida_Celula_Estorno:

    Saida_Celula_Estorno = gErr

    Select Case gErr

        Case 34838
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165434)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 34813

    Call Limpa_Tela_ProducaoSaida

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 34813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165435)

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
    
'    'Verifica se o código foi preenchido
'    If Len(Trim(Codigo.ClipText)) = 0 Then gError 34815
'
'    'Verifica se a Data foi preenchida
'    If Len(Trim(Data.ClipText)) = 0 Then gError 34820
    
    lErro = CF2(Me, "ReqProd_Gravar_Registro2")
    If lErro <> SUCESSO Then gError 198833

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovs
    If objGrid.iLinhasExistentes = 0 Then gError 34821

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 34822

        'Verifica se o Almoxarifado foi informado
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 And left(GridMovs.TextMatrix(iIndice, 0), 1) <> "#" Then gError 34823

        'Verifica se OPCodigo foi preenchido
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_OPCodigo_Col))) = 0 And left(GridMovs.TextMatrix(iIndice, 0), 1) <> "#" Then gError 34824

        'Verifica se ProdutoOP foi preenchido
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_ProdutoOP_Col))) = 0 And left(GridMovs.TextMatrix(iIndice, 0), 1) <> "#" Then gError 34825

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55221

    Next
    
    objMovEstoque.lCodigo = StrParaLong(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 34817

    If lErro = SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2")
        If vbMsgRes = vbNo Then gError 78792
    
    End If
    
    dtData = StrParaDate(Data.Text)
    If dtData = DATA_NULA Then dtData = Date

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(dtData)
    If lErro <> SUCESSO Then gError 92029

    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError 34826

    lErro = CF("Produto_Quant_Requerida_Superior", objMovEstoque, sReqMaiorOP)
    If lErro <> SUCESSO Then gError 126627

    'Se tem aviso => Exibe
    If Len(Trim(sReqMaiorOP)) > 0 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_REQUISICAO_MAIOR_OP", sReqMaiorOP)
        If vbMsgRes = vbNo Then gError 126628
    End If
    
    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 185847
    '###############################################
    
    lErro = CF2(Me, "ReqProd_Gravar_Registro", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 34827

'    'Grava no BD(inclusive dados contábeis)
'    lErro = CF("MovEstoque_ReqProducao_Grava", objMovEstoque, objContabil)
'    If lErro <> SUCESSO Then gError 34827

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
            gobjAnotacao.sID = CStr(objMovEstoque.iFilialEmpresa) & "," & CStr(objMovEstoque.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 34827
            
        End If
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 34815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 34816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case 34817, 34826, 34827, 78792, 92029, 126627, 126628, 185847, 198833

        Case 34818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_MOVTO_ESTOQUE_NAO_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 34819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVTO_ESTOQUE_CADASTRADO", gErr, objMovEstoque.lCodigo)

        Case 34820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
            Data.SetFocus

        Case 34821
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 34822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 34823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 34824
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPCODIGO_NAO_PREENCHIDO", gErr, iIndice)

        Case 34825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO", gErr, iIndice)

        Case 55221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165436)

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

    'Não preenche Tipo de Movimento a nível de Objeto
    If Len(Trim(Codigo.Text)) > 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

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
        If lErro <> SUCESSO Then gError 34829

    Next
    
    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_REQPROD)
    If lErro <> SUCESSO Then gError 185848
    
    Set objMovEstoque.objProjetoInfo = objProjetoInfo
    '##################################################

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 34828, 34829, 185848

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165437)

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
            If gobjMovEst.colItens(iIndice).lNumIntDoc = 0 Then gError 126675
            
            'Seta o Tipo como estorno
            If lBenef = BENEF_COMBO_BENEF Then
                iTipoMovEstoque = MOV_EST_ESTORNO_REQ_PRODUCAO_BENEF3
            ElseIf lBenef = BENEF_COMBO_DISP Then
                iTipoMovEstoque = MOV_EST_ESTORNO_REQ_PRODUCAO
            ElseIf lBenef = BENEF_COMBO_OUTROS Then
                iTipoMovEstoque = MOV_EST_ESTORNO_REQ_PRODUCAO_OUTROS
            End If

        Else
            
            If lBenef = BENEF_COMBO_BENEF Then
                iTipoMovEstoque = MOV_EST_REQ_PRODUCAO_BENEF3
            ElseIf lBenef = BENEF_COMBO_DISP Then
                iTipoMovEstoque = MOV_EST_REQ_PRODUCAO
            ElseIf lBenef = BENEF_COMBO_OUTROS Then
                iTipoMovEstoque = MOV_EST_REQ_PRODUCAO_OUTROS
            End If
            
        End If
        
        'Pega o Código do Almoxarifado
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = .TextMatrix(iIndice, iGrid_Almoxarifado_Col)

        If Len(Trim(objAlmoxarifado.sNomeReduzido)) <> 0 Then

            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 34831
    
            If lErro = 25056 Then gError 34832

        End If

        sProdutoFormatado = ""
        sProdutoOPFormatado = ""

        If Len(Trim(.TextMatrix(iIndice, iGrid_Ccl_Col))) > 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", .TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 34833

        Else

            sCclFormatada = ""

        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilEst_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilEst_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49660
        
        Else
            sContaFormatadaEst = ""
        End If
                
        If .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then gError 52054
        
        Else
            sContaFormatadaProducao = ""
        End If
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34834

        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_ProdutoOP_Col), sProdutoOPFormatado, iProdutoOPPreenchido)
        If lErro <> SUCESSO Then gError 34892
        
        'Move os dados do rastreamento para a Memória
        lErro = Move_RastroEstoque_Memoria(iIndice, colRastreamentoMovto)
        If lErro <> SUCESSO Then gError 78246
        
        Set objItemMovEst = objMovEstoque.colItens.Add(gobjMovEst.colItens(iIndice).lNumIntDoc, iTipoMovEstoque, 0, 0, sProdutoFormatado, .TextMatrix(iIndice, iGrid_Descricao_Col), .TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, .TextMatrix(iIndice, iGrid_Almoxarifado_Col), 0, sCclFormatada, CLng(.TextMatrix(iIndice, iGrid_Estorno_Col)), .TextMatrix(iIndice, iGrid_OPCodigo_Col), sProdutoOPFormatado, sContaFormatadaProducao, sContaFormatadaEst, 0, colRastreamentoMovto, Nothing, DATA_NULA)
        
        Set objItemMovEst.colItensRomaneioGrade = gobjMovEst.colItens(iIndice).colItensRomaneioGrade

    End With

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 34831, 34833, 34834, 49660, 53054, 78246

        Case 34832
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 126675
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165438)

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
    If lErro <> SUCESSO Then gError 78240
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78241

    If lErro = 28030 Then gError 78242
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            
                objRastreamentoMovto.sLote = GridMovs.TextMatrix(iLinha, iGrid_Lote_Col)
                
            Else
                
                lErro = CF2(Me, "Avisa_Nao_Preenchimento_Lote", iLinha)
                If lErro <> SUCESSO Then gError 199168
            
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            
            'Se colocou o Número do Lote e não colocou a filial ---> Erro
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                
                If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 78771
            
                objRastreamentoMovto.sLote = GridMovs.TextMatrix(iLinha, iGrid_Lote_Col)
                objRastreamentoMovto.iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))
            
            End If
            
            
            'Se colocou a Filial da O.P. e não colocou o lote ---> Erro
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 Then
            
                'Se não colocou o Número do Lote ---> Erro
                If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 78243
            
            End If
           
        '###########################################################
        'Inserido por Wagner 13/03/2006
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
           
            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iLinha)
                objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
        '###########################################################
            
        End If
        
        'Se o Número do Lote
        If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
        
            If Len(Trim(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then objRastreamentoMovto.dQuantidade = CDbl(GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col))
            objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
            objRastreamentoMovto.sProduto = sProdutoFormatado
            
            '######################################################
            'Alterado por Wagner 13/03/2006
            If objProduto.iRastro <> PRODUTO_RASTRO_NUM_SERIE Then
                colRastreamentoMovto.Add objRastreamentoMovto
            End If
            '######################################################
        
        End If
        
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case 78240, 78241, 199168
        
        Case 78242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 78770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 78771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165439)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 34835

    Call Limpa_Tela_ProducaoSaida

    'Coloca a Data atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 34835

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165440)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ProducaoSaida()

Dim lErro As Long
Dim iIndice As Integer

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)
    
    LblUM.Caption = ""
    
    For iIndice = 1 To objGrid.iLinhasExistentes 'm
        GridMovs.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'prepara o Grid para permitir inserir e excluir Linhas
    objGrid.iProibidoIncluir = 0
    objGrid.iProibidoExcluir = 0
    Call Grid_Inicializa(objGrid)
    
    Set gobjMovEst = New ClassMovEstoque

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    lCodigoAntigo = 0
    
    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 13/03/2006

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    Call CF2(Me, "ReqProd_Limpa_Tela")
    
    Set gobjAnotacao = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

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
    Set objEventoRastroLote = Nothing 'Inserido por Wagner
    Set objEventoBotaoOP = Nothing
    
    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set gcolcolRastreamentoSerie = Nothing 'Inserido por Wagner 13/03/2006
    
    Set gobjTelaProjetoInfo = Nothing
    
    Set gobjAnotacao = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
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

    '#################################################################
    'Inserido por Wagner 13/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
    If lErro <> SUCESSO Then gError 177299
    '#################################################################

    Exit Sub
    
Erro_Estorno_Click:

    Select Case gErr
    
        Case 177299
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177300)
    
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
        If lErro <> SUCESSO Then gError 39626
        
        If objGridInt.objGrid Is GridMovs Then
        
            Select Case GridMovs.Col
    
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 34839
    
                Case iGrid_OPCodigo_Col
    
                    lErro = Saida_Celula_OPCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 34840
    
                Case iGrid_Benef_Col
                    lErro = Saida_Celula_Benef(objGridInt)
                    If lErro <> SUCESSO Then gError 91272

                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 34841
    
                Case iGrid_Almoxarifado_Col
    
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 34842
                    
                Case iGrid_Ccl_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 34843
    
                Case iGrid_Estorno_Col
                    lErro = Saida_Celula_Estorno(objGridInt)
                    If lErro <> SUCESSO Then gError 34844
    
                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 34845
    
                Case iGrid_ProdutoOP_Col
                    lErro = Saida_Celula_ProdutoOP(objGridInt)
                    If lErro <> SUCESSO Then gError 34846
                
                Case iGrid_ContaContabilEst_Col
                    lErro = Saida_Celula_ContaContabilEst(objGridInt)
                    If lErro <> SUCESSO Then gError 49656
                    
                Case iGrid_ContaContabilProducao_Col
                    lErro = Saida_Celula_ContaContabilProducao(objGridInt)
                    If lErro <> SUCESSO Then gError 52200
                    
                Case iGrid_Lote_Col
                    lErro = Saida_Celula_Lote(objGridInt)
                    If lErro <> SUCESSO Then gError 78223
                
                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError 78754
                
                Case Else
                    lErro = Saida_Celula_Grid(objGridInt)
                    If lErro <> SUCESSO Then gError 34847
    
            End Select
        
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 34848

        iAlterado = REGISTRO_ALTERADO

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 34839, 34840, 34841, 34842, 34843, 34844, 34845, 34846, 34847, 49656, 52200, 78223, 78754

        Case 34848
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 39626

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165441)

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
Dim objProduto As New ClassProduto 'Inserido por Wagner 15/03/2006

On Error GoTo Erro_Saida_Celula_Benef

    Set objGridInt.objControle = Benef
    
    If Len(Trim(Benef.Text)) > 0 Then

        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 199182

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 199183
        '###########################################################

        GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col) = Benef.Text

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))) <> 0 Then

            If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then   'Alterado por Wagner 15/03/2006
                lErro = QuantDisponivel_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError 199184
            Else
                lErro = QuantLote_Calcula(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError 199185
            End If
            
        Else
            QuantDisponivel.Caption = ""
        End If
        
        
    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If
    

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91278
    
    Saida_Celula_Benef = SUCESSO

    Exit Function

Erro_Saida_Celula_Benef:

   Saida_Celula_Benef = gErr

    Select Case gErr

        Case 91278, 199182 To 199185
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165442)

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
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
    
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78775
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78776
            
        If lErro = 28030 Then gError 78777
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then

                objRastroLote.sCodigo = Lote.Text
                objRastroLote.sProduto = sProdutoFormatado

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78778

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78779

                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78780
                    
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then

                If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col))) > 0 Then


                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col))

                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78784

                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78785

                    'Preenche a Quantidade do Lote
                    lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78786
                
                Else
        
                    'Preenche a Quantidade do Lote
                    lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                    If lErro <> SUCESSO Then gError 78787
                
                End If
                            
            '###############################################################
            'Inserido por Wagner 15/03/2006
            'Se for rastro por série
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                'Preenche a Quantidade do Lote
                lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
                If lErro <> SUCESSO Then gError 78787
            '###############################################################
                            
            End If
            
        End If
    
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 78788

    End If
            
    If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then

        lErro = Testa_QuantRequisitada(StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 78776

    End If

    '###################################################################
    'Inserido por Wagner 13/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
    If lErro <> SUCESSO Then gError 141912
    '###################################################################

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78790

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 78775, 78776, 78778, 78780, 78784, 78786, 78787, 78788, 78789, 78790, 78776, 141912
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78779, 78785
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165443)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
Dim iEntradaSaida As Integer

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78740
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78741
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 78742
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78743
                    
        End If
        
        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) > 0 Then
                    
            lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78747
                                
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objRastroLote.sCodigo = GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
            
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78748
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78749
                            
                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError 78750
                
            End If
            
        End If
        
    Else
    
        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 78751
    
    End If
        
    If Len(Trim(QuantDisponivel.Caption)) <> 0 And GridMovs.TextMatrix(GridMovs.Row, iGrid_Estorno_Col) <> "1" Then

        lErro = Testa_QuantRequisitada(StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 78777

    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78753

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 78740, 78741, 78744, 78747, 78748, 78750, 78751, 78752, 78753, 78777
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78745
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78749
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165444)

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
    If lErro <> SUCESSO Then gError 34849

    Saida_Celula_Grid = SUCESSO

    Exit Function

Erro_Saida_Celula_Grid:

    Saida_Celula_Grid = gErr

    Select Case gErr

        Case 34849

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165445)

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

'#################################
'Inserido por Wagner 15/03/2006
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
'#################################

On Error GoTo Erro_GridMovs_RowColChange

    Call Grid_RowColChange(objGrid)

    If (GridMovs.Row <> iLinhaAntiga) Then

        'Guarda a Linha usada
        iLinhaAntiga = GridMovs.Row

        '###########################################################
        'Inserido por Wagner 15/03/2006
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141946
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141947
        '###########################################################

        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then  'Alterado por Wagner 15/03/2006
            lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 34850
        Else
            lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78766
        End If

    End If

    Exit Sub

Erro_GridMovs_RowColChange:

    Select Case gErr

        Case 34850, 78766

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165446)

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

Dim lErro As Long

    Call Grid_Campo_Recebe_Foco(objGrid)

    If gobjEST.iInventarioCodBarrAuto = 1 Then

        If objGrid.lErroSaidaCelula = 0 Then

            lErro = Trata_CodigoBarras1

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
        If lErro <> SUCESSO Then gError 55215
        
    Else
    
        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado, objProduto)
        If lErro <> SUCESSO Then gError 55216

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 55217

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case 55215, 55216, 55217

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165447)

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
        If lErro <> SUCESSO Then gError 30275
    
        If GridMovs.Row >= GridMovs.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 30276
    
                If lErro = 28030 Then gError 30277
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 30278
    
            If lErro = 25056 Then gError 30279
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 30280
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 30281
                
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

        Case 30275, 30276, 30278, 30280, 30281

        Case 30277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 30279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165448)

    End Select

    Exit Function

End Function

Private Function QuantLote_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantLote_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantLote_Calcula1(sProduto, sAlmoxarifado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78755
        
    Else
    
        lErro = QuantLote_Calcula1(sProduto, sAlmoxarifado, GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)), objProduto)
        If lErro <> SUCESSO Then gError 78756

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 78757
    QuantLote_Calcula = SUCESSO

    Exit Function

Erro_QuantLote_Calcula:

    QuantLote_Calcula = gErr

    Select Case gErr

        Case 78755, 78756, 78757

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165449)

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
        If lErro <> SUCESSO Then gError 78758
    
        If GridMovs.Row >= GridMovs.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78759
    
                If lErro = 28030 Then gError 78760
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 78761
    
            If lErro = 25056 Then gError 78762
    
            objRastreamentoLoteSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
            objRastreamentoLoteSaldo.sProduto = sProdutoFormatado
            objRastreamentoLoteSaldo.sLote = sLote
            objRastreamentoLoteSaldo.iFilialOP = iFilialOP
            
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoLoteSaldo)
            If lErro <> SUCESSO And lErro <> 78633 Then gError 78763
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 78633 Then
            
                 QuantDisponivel.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovs.TextMatrix(GridMovs.Row, iGrid_UnidadeMed_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 78764
                
                Call Combo_Obtem_ItemData(Benef, GridMovs.TextMatrix(GridMovs.Row, iGrid_Benef_Col), lBenef)
        
                If lBenef = BENEF_COMBO_BENEF Then
                    QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantBenef3 * dFator)
                ElseIf lBenef = BENEF_COMBO_OUTROS Then
                    QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantOutras3 * dFator)
                Else
                    QuantDisponivel.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDispNossa * dFator)
                End If
    
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

        Case 78758, 78759, 78761, 78763, 78764

        Case 78760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 78762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165450)

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
                If lErro <> SUCESSO Then gError 30403
    
            End If
    
        End If

    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = gErr

    Select Case gErr

        Case 30403

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165451)

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
        If lErro <> SUCESSO Then gError 30233

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 30209
    
        If lErro = 28030 Then gError 30210
    
        For iIndice = 1 To objGrid.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If GridMovs.Row <> iIndice Then
    
                sCodProduto = GridMovs.TextMatrix(iIndice, iGrid_Produto_Col)
                sAlmoxarifado = GridMovs.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
                iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(iIndice, iGrid_FilialOP_Col))
                sLote = GridMovs.TextMatrix(iIndice, iGrid_Lote_Col)
    
                lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 55346
    
                'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) And iFilialOPAtual = iFilialOP And UCase(sLoteAtual) = UCase(sLote) Then
    
                    'Verifica se há alguma QuanTidade informada
                    If Len(Trim(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 Then
    
                        sUnidadeProd = GridMovs.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
                        dQuantidadeProd = CDbl(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))
    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                        If lErro <> SUCESSO Then gError 30211
    
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

        Case 30209, 30211, 30233, 55346

        Case 30210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165452)

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
                        If lErro <> SUCESSO Then gError 64206

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
                    If lErro <> SUCESSO Then gError 79943
                    
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
            Error 39651

        End Select

        Calcula_Mnemonico = SUCESSO

        Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39651
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64206
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165453)

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
        If lErro <> SUCESSO And lErro <> 25060 Then gError 49690
        
        If lErro = 25060 Then gError 52005
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 49691
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError 49692
        
        If lErro = SUCESSO Then
            
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49693
        
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
            
            Case 49690, 49691, 49692
            
            Case 49693
                lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
                
            Case 52005
                 lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
        
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165454)
        
        End Select
        
        Exit Function
        
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REQUISICAO_MATERIAL_PRODUCAO_MOVIMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Requisição de Material Para Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ProducaoSaida"
    
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
        ElseIf Me.ActiveControl Is Produto Or Me.ActiveControl Is ProdutoOP Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is OPCodigo Then
            Call BotaoOP_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is ContaContabilEst Or Me.ActiveControl Is ContaContabilProducao Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Lote Then 'Inserido por Wagner
            Call BotaoLote_Click
        Else
            Call CF2(Me, "UserControl_KeyDown_CTProducaoSaida", KeyCode, Shift)
        End If
    
        ElseIf KeyCode = KEYCODE_CODBARRAS Then
            Call Trata_CodigoBarras1
    
    End If

End Sub


Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165455)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Já foi utilizada em outras telas do est
Function RastreamentoMovto_Le_DocOrigem(lNumIntDocOrigem As Long, iTipoDocOrigem As Integer, colRastreamentoMovto As Collection) As Long
'Lê a tabela de RastreamentoMovto através do Movimento de Estoque

Dim lErro As Long
Dim tRastreamentoMovto As typeRastreamentoMovto
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim lComando As Long

On Error GoTo Erro_RastreamentoMovto_Le_DocOrigem
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 78411
    
    tRastreamentoMovto.sProduto = String(STRING_PRODUTO, 0)
    tRastreamentoMovto.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
    'Lê o Rastreamento Movto
    lErro = Comando_Executar(lComando, "SELECT RastreamentoMovto.NumIntDoc, RastreamentoMovto.TipoDocOrigem, RastreamentoMovto.NumIntDocOrigem, RastreamentoMovto.Produto, RastreamentoMovto.Quantidade, RastreamentoLote.Lote, RastreamentoLote.FilialOP FROM RastreamentoMovto, RastreamentoLote WHERE RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND TipoDocOrigem = ? AND NumIntDocOrigem = ?" _
    , tRastreamentoMovto.lNumIntDoc, tRastreamentoMovto.iTipoDocOrigem, tRastreamentoMovto.lNumIntDocOrigem, tRastreamentoMovto.sProduto, tRastreamentoMovto.dQuantidade, tRastreamentoMovto.sLote, tRastreamentoMovto.iFilialOP, iTipoDocOrigem, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 78412
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78413

    Do While lErro = AD_SQL_SUCESSO
        
        Set objRastreamentoMovto = New ClassRastreamentoMovto
        
        'passa para o objeto
        objRastreamentoMovto.lNumIntDoc = tRastreamentoMovto.lNumIntDoc
        objRastreamentoMovto.iTipoDocOrigem = tRastreamentoMovto.iTipoDocOrigem
        objRastreamentoMovto.lNumIntDocOrigem = tRastreamentoMovto.lNumIntDocOrigem
        objRastreamentoMovto.sProduto = tRastreamentoMovto.sProduto
        objRastreamentoMovto.dQuantidade = tRastreamentoMovto.dQuantidade
        objRastreamentoMovto.sLote = tRastreamentoMovto.sLote
        objRastreamentoMovto.iFilialOP = tRastreamentoMovto.iFilialOP
        
        colRastreamentoMovto.Add objRastreamentoMovto
                
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78421
    
    Loop
            
    Call Comando_Fechar(lComando)
    
    RastreamentoMovto_Le_DocOrigem = SUCESSO
        
    Exit Function
    
Erro_RastreamentoMovto_Le_DocOrigem:

    RastreamentoMovto_Le_DocOrigem = gErr
    
    Select Case gErr
            
        Case 78411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 78412, 78413, 78421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165456)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Copiada da class CTInventario de GlobaisTelasEst
Function RastreamentoLoteSaldo_Le(objRastroLoteSaldo As ClassRastreamentoLoteSaldo) As Long
'Lê a tabela de Rastreamento Lote Saldo

Dim lErro As Long
Dim lComando As Long
Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo

On Error GoTo Erro_RastreamentoLoteSaldo_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 78630

    tRastroLoteSaldo.sProduto = String(STRING_PRODUTO, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT RastreamentoLoteSaldo.Produto, Almoxarifado, NumIntDocLote, QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3 FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND RastreamentoLote.Lote = ? AND RastreamentoLote.FilialOP = ?", _
    tRastroLoteSaldo.sProduto, tRastroLoteSaldo.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote, tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantReservada, tRastroLoteSaldo.dQuantReservadaConsig, tRastroLoteSaldo.dQuantEmpenhada, tRastroLoteSaldo.dQuantPedida, tRastroLoteSaldo.dQuantRecIndl, tRastroLoteSaldo.dQuantIndOutras, tRastroLoteSaldo.dQuantDefeituosa, tRastroLoteSaldo.dQuantConsig3, tRastroLoteSaldo.dQuantConsig, tRastroLoteSaldo.dQuantDemo3, tRastroLoteSaldo.dQuantDemo, tRastroLoteSaldo.dQuantConserto3, tRastroLoteSaldo.dQuantConserto, tRastroLoteSaldo.dQuantOutras3, tRastroLoteSaldo.dQuantOutras, tRastroLoteSaldo.dQuantOP, tRastroLoteSaldo.dQuantBenef, tRastroLoteSaldo.dQuantBenef3, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote, objRastroLoteSaldo.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 78631

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78632

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 78633

    objRastroLoteSaldo.sProduto = tRastroLoteSaldo.sProduto
    objRastroLoteSaldo.iAlmoxarifado = tRastroLoteSaldo.iAlmoxarifado
    objRastroLoteSaldo.lNumIntDocLote = tRastroLoteSaldo.lNumIntDocLote
    objRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa
    objRastroLoteSaldo.dQuantReservada = tRastroLoteSaldo.dQuantReservada
    objRastroLoteSaldo.dQuantReservadaConsig = tRastroLoteSaldo.dQuantReservadaConsig
    objRastroLoteSaldo.dQuantEmpenhada = tRastroLoteSaldo.dQuantEmpenhada
    objRastroLoteSaldo.dQuantPedida = tRastroLoteSaldo.dQuantPedida
    objRastroLoteSaldo.dQuantRecIndl = tRastroLoteSaldo.dQuantRecIndl
    objRastroLoteSaldo.dQuantIndOutras = tRastroLoteSaldo.dQuantIndOutras
    objRastroLoteSaldo.dQuantDefeituosa = tRastroLoteSaldo.dQuantDefeituosa
    objRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3
    objRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig
    objRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3
    objRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo
    objRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3
    objRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto
    objRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3
    objRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras
    objRastroLoteSaldo.dQuantOP = tRastroLoteSaldo.dQuantOP
    objRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef
    objRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLoteSaldo_Le = SUCESSO

    Exit Function

Erro_RastreamentoLoteSaldo_Le:

    RastreamentoLoteSaldo_Le = gErr

    Select Case gErr

        Case 78630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78631, 78632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO", gErr, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote)

        Case 78633 'RastreamentoLoteSaldo não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165457)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

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

'############################################
'Inserido por Wagner
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
                If sProduto = sProdutoLinha And sCodOP = sOPLinha Then gError 132430
            End If
        Next
    End If

    Valida_Repeticao = SUCESSO

    Exit Function

Erro_Valida_Repeticao:

    Valida_Repeticao = gErr

    Select Case gErr

         Case 132430
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE2", gErr, sProdutoLinha, sOPLinha, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165458)

    End Select

    Exit Function
    
End Function
'############################################

'################################################################
'Inserido por Wagner 04/10/2005
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

    If (GridMovs.Row = 0) Then gError 140223

    sCodProduto = GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col)
    sLote = GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 140224

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 140225
    
    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))) = 0 Then gError 177296
    
    Set objAlmoxarifado = New ClassAlmoxarifado
    
    objAlmoxarifado.sNomeReduzido = GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 177297
    
    If Len(Trim(sLote)) > 0 Then
        objRastroLoteSaldo.sLote = sLote
    End If

    colSelecao.Add sProdutoFormatado
    colSelecao.Add objAlmoxarifado.iCodigo

    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoRastroLote, "Produto = ? AND Almoxarifado = ?")

    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr

        Case 140224, 177297
        
        Case 140223
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 140225
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 177296
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, GridMovs.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165459)

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
        
        '############################################################
        'Inserido por Wagner 13/03/2006
        'Carrega as séries na coleção global
        lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovs.Row), StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)), GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col), GridMovs.Row)
        If lErro <> SUCESSO Then gError 141913
        '############################################################
        
        GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col) = objRastroLoteSaldo.sLote
    End If

    objProduto.sCodigo = objRastroLoteSaldo.sProduto
            
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 140226
    
    'Se for rastro por lote
     If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
        'Preenche a Quantidade do Lote
        lErro = QuantLote_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col), Lote.Text, Codigo_Extrai(GridMovs.TextMatrix(GridMovs.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 140227

    Else

        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), GridMovs.TextMatrix(GridMovs.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 140228
    
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoRastroLote_evSelecao:

    Select Case gErr
    
        Case 140226 To 140228, 141913
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165460)

    End Select

    Exit Sub

End Sub
'######################################################################

'#####################################################
'Inserido por Wagner 13/03/2006
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
    If GridMovs.Row = 0 Then gError 141914
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))) = 0 Then gError 141915
    If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Lote_Col))) = 0 Then gError 177303
    If StrParaDbl(GridMovs.TextMatrix(GridMovs.Row, iGrid_Quantidade_Col)) = 0 Then gError 177304
        
    'Formata o produto
    lErro = CF("Produto_Formata", GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141916
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141917
       
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
    If lErro <> SUCESSO Then gError 141918
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
    
        Case 141914
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 141915 To 141918
        
        Case 177303
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, GridMovs.Row)
        
        Case 177304
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, GridMovs.Row)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141919)
    
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
    If lErro <> SUCESSO Then gError 141921
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141922
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 141923

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
                If vbResult = vbNo Then gError 141920
            Else
                vbResult = vbNo
        
            End If
            
        Else
            vbResult = vbYes
            iTipoMovtoAnt = iTipoMovto
        End If
                    
        If vbResult = vbYes Then
                            
            If Len(Trim(sLoteIni)) <> 0 Then
            
                If Not IsNumeric(right(sLoteIni, objProduto.iSerieParteNum)) Then gError 141924
                
                Set objRastreamentoSerieIni = New ClassRastreamentoLote
                
                objRastreamentoSerieIni.sProduto = objProduto.sCodigo
                objRastreamentoSerieIni.iFilialOP = Codigo_Extrai(GridMovs.TextMatrix(iLinha, iGrid_FilialOP_Col))
                objRastreamentoSerieIni.sCodigo = sLoteIni
                
                lErro = CF("RastreamentoLote_Le", objRastreamentoSerieIni)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 141925
                
                If lErro <> SUCESSO Then gError 141926
                              
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
                If lErro <> SUCESSO And lErro <> 25056 Then gError 177237
                
                objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                
                lErro = CF("Rastreamento_Serie_Gera", objItemMovEstoque, objProduto, sLoteIni, colRastreamentoMovtoAux)
                If lErro <> SUCESSO Then gError 177240

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
    
        Case 141921, 141922, 141925, 141927, 141920, 141929, 177237, 177240
        
        Case 141923
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 141924
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, right(sLoteIni, objProduto.iSerieParteNum))
        
        Case 141926
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo, objRastreamentoSerieIni.iFilialOP)
        
        Case 141928
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerie.sProduto, objRastreamentoSerie.sCodigo, objRastreamentoSerie.iFilialOP)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141930)

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
            If lErro <> SUCESSO Then gError 141932

        End If

    End If
        
    GridMovs.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    Acerta_Quantidade_Rastreada = SUCESSO
    
    Exit Function

Erro_Acerta_Quantidade_Rastreada:

    Acerta_Quantidade_Rastreada = gErr

    Select Case gErr
    
        Case 141932

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141931)

    End Select

    Exit Function

End Function
'#####################################################

'#####################################
'Inserido por Wagner 03/08/2006
Sub BotaoProjetos_Click()
    Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Sub Projeto_GotFocus()
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Projeto_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR
        
        objItensPRJCR.sItem = CStr(iIndice)
        objItensPRJCR.sDescricao = GridMovs.TextMatrix(iIndice, iGrid_Produto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridMovs.TextMatrix(iIndice, iGrid_Quantidade_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_REQPROD
        
        colItensPRJ.Add objItensPRJCR

    Next
    
    If colItensPRJ.Count = 0 Then
        Set objItensPRJCR = New ClassItensPRJCR
        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_REQPROD
        colItensPRJ.Add objItensPRJCR
    End If
    
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

Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub
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

Public Function ReqProd_OP_Validate(ByVal objCT As Object, ByVal objOP As ClassOrdemDeProducao)
    ReqProd_OP_Validate = SUCESSO
End Function

Public Function ReqProd_BotaoGeraReq_Preenche_Grid(ByVal objCT As Object, ByVal colItensMovEstoque As colItensMovEstoque) As Long
    ReqProd_BotaoGeraReq_Preenche_Grid = SUCESSO
End Function

Public Function ReqProd_Gravar_Registro(ByVal objCT As Object, ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
    ReqProd_Gravar_Registro = CF("MovEstoque_ReqProducao_Grava", objMovEstoque, objContabil)
End Function

Public Function ReqProd_Preenche_Tela(ByVal objCT As Object, ByVal objMovEstoque As ClassMovEstoque) As Long
    ReqProd_Preenche_Tela = SUCESSO
End Function

Public Function ReqProd_Limpa_Tela(ByVal objCT As Object) As Long
    ReqProd_Limpa_Tela = SUCESSO
End Function

Public Function ReqProd_Gravar_Registro2(ByVal objCT As Object) As Long

On Error GoTo Erro_ReqProd_Gravar_Registro2

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 34815
    
    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 34820

    ReqProd_Gravar_Registro2 = SUCESSO

    Exit Function

Erro_ReqProd_Gravar_Registro2:

    ReqProd_Gravar_Registro2 = gErr
    
    Select Case gErr

        Case 34815
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 34820
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165436)

    End Select

    Exit Function
    
End Function

Public Function Avisa_Nao_Preenchimento_Lote(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    Avisa_Nao_Preenchimento_Lote = SUCESSO
End Function

Public Sub UserControl_KeyDown_CTProducaoSaida(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
'
End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
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

Public Sub BotaoRatear_Click()

Dim lErro  As Long
Dim iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim colItens As New colItensMovEstoque

On Error GoTo Erro_BotaoRatear_Click

    For iIndice = 1 To objGrid.iLinhasExistentes
        lErro = Move_Itens_Memoria(iIndice, objMovEstoque)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next
    
    lErro = CF("ReqProducao_Rateia_Insumos", objMovEstoque, colItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = objGrid.iLinhasExistentes To 1 Step -1
        Call gobjTelaProjetoInfo.Remove_Linha(iIndice)
    Next

    Call Grid_Limpa(objGrid)

    'Remove os ítens de gobjMovEst
    Set gobjMovEst.colItens = New colItensMovEstoque

    'preenche o grid de itens requisitados com os itens dos kits que compõem a ordem de produção
    lErro = BotaoGeraReq_Preenche_Grid(colItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_BotaoRatear_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211680)
            
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGerarReqRateio_Click()

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao
Dim colItensMovEstoque As New colItensMovEstoque
Dim colItensMovEstoqueAux As New colItensMovEstoque
Dim vbMsg As VbMsgBoxResult
Dim iIndice As Integer
Dim objItemMovEst As ClassItemMovEstoque
Dim objItemMovEstAux As ClassItemMovEstoque
Dim bAchou As Boolean

On Error GoTo Erro_BotaoGerarReqRateio_Click
          
    'se já existem linhas preenchidas no grid
    If objGrid.iLinhasExistentes > 0 Then

        DoEvents
        
        'pergunta se quer apagá-las
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_APAGAR_GRID")
        
        'se a resposta for afirmativa ==> remove as linhas do grid
        If vbMsg = vbYes Then
        
            For iIndice = objGrid.iLinhasExistentes To 1 Step -1
                Call gobjTelaProjetoInfo.Remove_Linha(iIndice)
            Next
        
            Call Grid_Limpa(objGrid)

            'Remove os ítens de gobjMovEst
            Set gobjMovEst.colItens = New colItensMovEstoque

        End If
        
    End If

    'se a ordem de produção não foi preenchida ==> erro
    If Len(Trim(OP.Text)) = 0 Then gError 34768

    objOP.sCodigo = OP.Text
    objOP.iFilialEmpresa = giFilialEmpresa

    'trata de gerar as requisições para todos os itens da OP
    lErro = OP_Requisicoes_Gera(objOP, colItensMovEstoque)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objItemMovEst In colItensMovEstoque
        bAchou = False
        For Each objItemMovEstAux In colItensMovEstoqueAux
            If objItemMovEst.sProduto = objItemMovEstAux.sProduto Then
                bAchou = True
                objItemMovEstAux.dQuantidade = objItemMovEstAux.dQuantidade + objItemMovEst.dQuantidade
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objItemMovEstAux = colItensMovEstoqueAux.Add(0, 0, 0, 0, objItemMovEst.sProduto, objItemMovEst.sProdutoDesc, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, 0, "", 0, "", 0, objItemMovEst.sOPCodigo, "", objItemMovEst.sContaContabilAplic, "", 0, Nothing, Nothing, DATA_NULA)
            objItemMovEstAux.iBenef = objItemMovEst.iBenef
        End If
    Next

    'preenche o grid de itens requisitados com os itens dos kits que compõem a ordem de produção
    lErro = BotaoGeraReq_Preenche_Grid(colItensMovEstoqueAux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_BotaoGerarReqRateio_Click:

    Select Case gErr

        Case 34768
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREECHIDA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165413)

    End Select

    Exit Sub
    
End Sub


Public Function Trata_CodigoBarras1() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String
Dim sCodBarras As String
Dim sCodBarrasOriginal As String
Dim dCusto As Double
Dim objItemMovEst As New ClassItemMovEstoque

On Error GoTo Erro_Trata_CodigoBarras1

    If objGrid.iLinhasExistentes + 1 = GridMovs.Row Then
    
        'Verifica se o Produto está preenchido
        If Len(Trim(GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col))) = 0 Then
            
            If ActiveControl Is Produto Then
                    
                    Set objGrid.objControle = Produto
            
                    lErro = Grid_Abandona_Celula(objGrid)
                    If lErro <> SUCESSO Then gError 210821
                    
            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 210822
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 210823
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 210826
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 210824
        
                Me.Show
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = Produto.Text
                
                gError 210866
                
'                If Not ActiveControl Is Produto Then
'                    GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = Produto.Text
'
'                    'Preenche a Linha do Grid
'                    lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
'                    If lErro <> SUCESSO Then gError 210825
'
'
'                End If
'
'                GridMovs.SetFocus
'                GridMovs.FocusRect = flexFocusHeavy
    
            Else
            
                gError 210828
    
    
            End If
    
        End If
    
    End If

    Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_Trata_CodigoBarras1:

    Trata_CodigoBarras1 = gErr

'    GridMovs.TextMatrix(GridMovs.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 210821, 210822, 210823, 210825, 210828, 210866

        Case 210826
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 210824
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210827)

    End Select

    Exit Function

End Function

