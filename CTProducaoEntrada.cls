VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProducaoEntrada"
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

Public gobjMovEst As ClassMovEstoque

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

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
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"

Dim gcolcolRastreamentoSerie As Collection 'Inserido por Wagner 14/03/2006
Dim gcolProdutos As Collection

'Declaração das Variáveis Globais
Public iAlterado As Integer
Dim iCodigoAlterado As Integer
Dim iFrameAtual As Integer
Dim lCodigoAntigo As Long

'GRID
Public objGrid As AdmGrid
Public iGrid_Sequencial_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Benef_Col  As Integer
Public iGrid_HorasMaquina_Col As Integer
Public iGrid_DataProducao_Col As Integer
Public iGrid_Ccl_Col As Integer
Public iGrid_Descricao_Col As Integer
Public iGrid_OP_Col As Integer
Public iGrid_Lote_Col As Integer
Public iGrid_Estorno_Col As Integer
Public iGrid_ContaContabilEst_Col As Integer
Public iGrid_ContaContabilProducao_Col As Integer

'GRID APROPRIACAO
Dim objGridApropriacao As AdmGrid
Dim iGrid_ProdutoApr_Col As Integer
Dim iGrid_DescricaoApr_Col As Integer
Dim iGrid_ApropItemProd_Col As Integer
Dim iGrid_ProdutoIns_Col As Integer
Dim iGrid_DescricaoIns_Col As Integer
Dim iGrid_UnidadeMedApr_Col As Integer
Dim iGrid_QuantidadeApr_Col As Integer
Dim iGrid_RequisicaoApr_Col As Integer
Dim iGrid_LoteApr_Col As Integer
Dim iGrid_FilialOPApr_Col As Integer
Dim iGrid_ApropAutomatica_Col As Integer

'BROWSERS
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoOPPadrao As AdmEvento
Attribute objEventoOPPadrao.VB_VarHelpID = -1
Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoOP As AdmEvento
Attribute objEventoOP.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAprop As AdmEvento
Attribute objEventoProdutoAprop.VB_VarHelpID = -1
Private WithEvents objEventoRequisicoes As AdmEvento
Attribute objEventoRequisicoes.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento '
Attribute objEventoEstoque.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Movimentos = 1
Private Const TAB_Contabilizacao = 3

'--- inicio dos properties get dos controles da tela

Public Property Get ApropAutomatica() As Object
     Set ApropAutomatica = objUserControl.Controls("ApropAutomatica")
End Property

Public Property Get ApropItemProd() As Object
     Set ApropItemProd = objUserControl.Controls("ApropItemProd")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get CTBGerencial() As Object
     Set CTBGerencial = objUserControl.Controls("CTBGerencial")
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

Public Property Get BotaoProjetos() As Object
     Set BotaoProjetos = objUserControl.Controls("BotaoProjetos")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get BotaoSerie() As Object
     Set BotaoSerie = objUserControl.Controls("BotaoSerie")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get BotaoImprimirRotulos() As Object
     Set BotaoImprimirRotulos = objUserControl.Controls("BotaoImprimirRotulos")
End Property

Public Property Get botaoOPCodigoPadrao() As Object
     Set botaoOPCodigoPadrao = objUserControl.Controls("botaoOPCodigoPadrao")
End Property

Public Property Get OPCodigoPadrao() As Object
     Set OPCodigoPadrao = objUserControl.Controls("OPCodigoPadrao")
End Property

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get Benef() As Object
     Set Benef = objUserControl.Controls("Benef")
End Property

Public Property Get BotaoEstoque() As Object
     Set BotaoEstoque = objUserControl.Controls("BotaoEstoque")
End Property

Public Property Get DataProducao() As Object
     Set DataProducao = objUserControl.Controls("DataProducao")
End Property

Public Property Get Estorno() As Object
     Set Estorno = objUserControl.Controls("Estorno")
End Property

Public Property Get OPCodigo() As Object
     Set OPCodigo = objUserControl.Controls("OPCodigo")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get ContaContabilProducao() As Object
     Set ContaContabilProducao = objUserControl.Controls("ContaContabilProducao")
End Property

Public Property Get ContaContabilEst() As Object
     Set ContaContabilEst = objUserControl.Controls("ContaContabilEst")
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

Public Property Get Lote() As Object
     Set Lote = objUserControl.Controls("Lote")
End Property

Public Property Get HorasMaq() As Object
     Set HorasMaq = objUserControl.Controls("HorasMaq")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get BotaoOP() As Object
     Set BotaoOP = objUserControl.Controls("BotaoOP")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get BotaoCcls() As Object
     Set BotaoCcls = objUserControl.Controls("BotaoCcls")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get CclPadrao() As Object
     Set CclPadrao = objUserControl.Controls("CclPadrao")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get GridMovimentos() As Object
     Set GridMovimentos = objUserControl.Controls("GridMovimentos")
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

Public Property Get CodigoLabel() As Object
     Set CodigoLabel = objUserControl.Controls("CodigoLabel")
End Property

Public Property Get CclPadraoLabel() As Object
     Set CclPadraoLabel = objUserControl.Controls("CclPadraoLabel")
End Property

Public Property Get OPPadraoLabel() As Object
     Set OPPadraoLabel = objUserControl.Controls("OPPadraoLabel")
End Property

Public Property Get LoteApr() As Object
     Set LoteApr = objUserControl.Controls("LoteApr")
End Property

Public Property Get FilialOpApr() As Object
     Set FilialOpApr = objUserControl.Controls("FilialOpApr")
End Property

Public Property Get BotaoRequisicao() As Object
     Set BotaoRequisicao = objUserControl.Controls("BotaoRequisicao")
End Property

Public Property Get BotaoProdutoApror() As Object
     Set BotaoProdutoApror = objUserControl.Controls("BotaoProdutoApror")
End Property

Public Property Get RequisicaoApr() As Object
     Set RequisicaoApr = objUserControl.Controls("RequisicaoApr")
End Property

Public Property Get DescricaoItemIns() As Object
     Set DescricaoItemIns = objUserControl.Controls("DescricaoItemIns")
End Property

Public Property Get UnidadeMedApr() As Object
     Set UnidadeMedApr = objUserControl.Controls("UnidadeMedApr")
End Property

Public Property Get DescricaoItemApr() As Object
     Set DescricaoItemApr = objUserControl.Controls("DescricaoItemApr")
End Property

Public Property Get QuantidadeApr() As Object
     Set QuantidadeApr = objUserControl.Controls("QuantidadeApr")
End Property

Public Property Get ProdutoApr() As Object
     Set ProdutoApr = objUserControl.Controls("ProdutoApr")
End Property

Public Property Get ProdutoIns() As Object
     Set ProdutoIns = objUserControl.Controls("ProdutoIns")
End Property

Public Property Get GridApropriacao() As Object
     Set GridApropriacao = objUserControl.Controls("GridApropriacao")
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

Public Property Get HorasMaquina() As Object
     Set HorasMaquina = objUserControl.Controls("HorasMaquina")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Benef_Click()

    iAlterado = REGISTRO_ALTERADO
    
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

Public Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 83497

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83498

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        
        objEstoqueProduto.sAlmoxarifadoNomeReduzido = Almoxarifado.Text

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 83499
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 83498
        
        Case 83497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 83499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165300)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 89902

    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    '''03/09/01 - Marcelo inclusao da pergunta se deseja excluir a Entrada de Producao
    'Envia aviso perguntando se realmente deseja excluir a Entrada de Producao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_ENTRADAPRODUCAO", objMovEstoque.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a producao
        lErro = CF("MovEstoque_Producao_Exclui", objMovEstoque, objContabil)
        If lErro <> SUCESSO Then gError 89903

        Call Limpa_Tela_ProducaoEntrada
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 89902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 89903
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165301)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83500

    'Verifica se o produto está preenchido e se a linha corrente é diferente da linha fixa
    If iProdutoPreenchido = PRODUTO_PREENCHIDO And GridMovimentos.Row <> 0 Then

        'Preenche o Nome do Almoxarifado
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
        lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
        If lErro <> SUCESSO Then gError 83503

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 83500, 83503

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165302)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutoApror_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoProdutoApror_Click

    If GridApropriacao.Row = 0 Then gError 78160
    
    'Lista de produtos produzidos e inventariados
    Call Chama_Tela("ProdutoAtivoLista", colSelecao, objProduto, objEventoProdutoAprop)
    
   Exit Sub
   
Erro_BotaoProdutoApror_Click:

    Select Case gErr
        
        Case 78160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165303)
        
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("MovEstoque_Automatico", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then gError 57522

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    lCodigoAntigo = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 57522
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165304)
    
    End Select

    Exit Sub

End Sub

'Private Sub Limpa_colItensNumIntDoc(colItensNumIntDoc As Collection)
'Dim lErro As Long
'Dim iCount As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_Limpa_colItensNumIntDoc
'
'    iCount = colItensNumIntDoc.Count
'    Set colItensNumIntDoc = New Collection
'
'    For iIndice = 1 To iCount
'
'        colItensNumIntDoc.Add 0
'        GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = "0"
'
'    Next
'
'    Exit Sub
'
'Erro_Limpa_colItensNumIntDoc:
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165305)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then gError 52277
    
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = "" Then gError 52278

    sConta = String(STRING_CONTA, 0)

    If GridMovimentos.Col = iGrid_ContaContabilEst_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52279
    
    ElseIf GridMovimentos.Col = iGrid_ContaContabilProducao_Col Then

        lErro = CF("Conta_Formata", ContaContabilProducao.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52280

    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 52277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 52278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 52279, 52280

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165306)

    End Select

    Exit Sub

End Sub

Public Sub BotaoRequisicao_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoRequisicao_Click

    If GridApropriacao.Row = 0 Then gError 78166

    If GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoApr_Col) = "" Then gError 78167

    If Len(Trim(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_RequisicaoApr_Col))) <> 0 Then objMovEstoque.lCodigo = CLng(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_RequisicaoApr_Col))
    
    'Janaina
    colSelecao.Add MOV_EST_REQ_PRODUCAO
    colSelecao.Add MOV_EST_REQ_PRODUCAO_BENEF3
    colSelecao.Add MOV_EST_REQ_PRODUCAO_OUTROS
    'Janaina
    Call Chama_Tela("MovEstoqueLista1", colSelecao, objMovEstoque, objEventoRequisicoes)

    Exit Sub
   
Erro_BotaoRequisicao_Click:

    Select Case gErr
        
        Case 78166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 78167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165307)
        
    End Select
    
    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub DataProducao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub DataProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub DataProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DataProducao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FilialOpApr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOpApr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub FilialOpApr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub FilialOpApr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = FilialOpApr
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If GridMovimentos.Col <> iGrid_ContaContabilEst_Col And GridMovimentos.Col <> iGrid_ContaContabilProducao_Col Then
        Me.Show
        Exit Sub
    End If
        
    If objPlanoConta.sConta <> "" Then
   
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 52281
               
        If GridMovimentos.Col = iGrid_ContaContabilEst_Col Then
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = ContaContabilEst.Text
        Else
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxuta
            ContaContabilProducao.PromptInclude = True
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = ContaContabilProducao.Text
        End If
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 52281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165308)

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
    
        If Len(Trim(Codigo.ClipText)) > 0 Then
        
            objMovEstoque.lCodigo = Codigo.Text
            
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_PRODUCAO)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 34911
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 78889
            
            If lErro = SUCESSO Then
            
                If objMovEstoque.iTipoMov <> MOV_EST_PRODUCAO Then gError 34899
                
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")
                
                If vbMsg = vbNo Then gError 34912
                
                lErro = Preenche_Tela(objMovEstoque)
                If lErro <> SUCESSO Then gError 34913
                      
            End If
        
        End If
      
        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
      
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
            
        Case 34899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_PENTRADA", gErr, objMovEstoque.lCodigo)
            lCodigoAntigo = 0
            
        Case 34911, 34913
        
        Case 34912
            
        Case 78889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165309)
    
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

Public Sub Estorno_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao As String

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 14/03/2006
    Set gcolProdutos = New Collection
    
    'Inicializa todos os objeventos
    Set gobjMovEst = New ClassMovEstoque
    
    Set objEventoCodigo = New AdmEvento
    Set objEventoOPPadrao = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoOP = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoProdutoAprop = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoRequisicoes = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    
    'Mostra a Data Atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 30918
    
    'Inicializa Máscara de ProdutoApr
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoApr)
    If lErro <> SUCESSO Then gError 78153
    
    'Inicializa Máscara de ProdutoIns
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoIns)
    If lErro <> SUCESSO Then gError 78154
    
    'Inicializa mascara de contaContabilEst
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
    If lErro <> SUCESSO Then gError 49636
    
    'Inicializa a mascara de ContaContabilProducao
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducao)
    If lErro <> SUCESSO Then gError 52037

    'Inicializa Máscara para CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 30919

    Ccl.Mask = sMascaraCclPadrao
    CclPadrao.Mask = sMascaraCclPadrao

    'Formata Quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    QuantidadeApr.Format = FORMATO_ESTOQUE
    
    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78290
    
    'Inicialização do GridMovimentos
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridMovimentos(objGrid)
    If lErro <> SUCESSO Then gError 30920
    
    'Inicialização do GridApropriacao
    Set objGridApropriacao = New AdmGrid
    
    lErro = Inicializa_GridApropriacao(objGridApropriacao)
    If lErro <> SUCESSO Then gError 78135
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then gError 39623
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Inserido por Jorge Specian - 12/07/2005 - Tamanho do campo Codigo configurável
    '---------------------------------------
    OPCodigoPadrao.MaxLength = STRING_ORDEM_DE_PRODUCAO
    OPCodigo.MaxLength = STRING_ORDEM_DE_PRODUCAO
    Lote.MaxLength = STRING_LOTE_RASTREAMENTO
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 30918, 30919, 30920, 39623, 49636, 52037, 78135, 78153, 78154, 78290

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165310)

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
    If lErro <> SUCESSO Then gError 78291

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOpApr.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOpApr.ItemData(FilialOpApr.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 78291
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165311)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridMovimentos(objGridInt As AdmGrid) As Long
    Inicializa_GridMovimentos = CF2(Me, "ProdEnt_Inicializa_GridMovimentos", objGridInt)
End Function

Public Function ProdEnt_Inicializa_GridMovimentos(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Ordem Produçao")
    objGridInt.colColuna.Add ("Lote/ Série Inicial") 'Alterado por Wagner 14/03/2006
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Escaninho")
    objGridInt.colColuna.Add ("Horas Máquina")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil de Estoque")
    objGridInt.colColuna.Add ("Conta Contabil de Produção")
    objGridInt.colColuna.Add ("Estorno")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (OPCodigo.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Benef.Name)
    objGridInt.colCampo.Add (HorasMaq.Name)
    objGridInt.colCampo.Add (DataProducao.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ContaContabilEst.Name)
    objGridInt.colCampo.Add (ContaContabilProducao.Name)
    objGridInt.colCampo.Add (Estorno.Name)
    
    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_OP_Col = 5
    iGrid_Lote_Col = 6
    iGrid_Almoxarifado_Col = 7
    iGrid_Benef_Col = 8
    iGrid_HorasMaquina_Col = 9
    iGrid_DataProducao_Col = 10
    iGrid_Ccl_Col = 11
    iGrid_ContaContabilEst_Col = 12
    iGrid_ContaContabilProducao_Col = 13
    iGrid_Estorno_Col = 14
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    ProdEnt_Inicializa_GridMovimentos = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridApropriacao(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Insumo")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto Produzido")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Requisição")
    objGridInt.colColuna.Add ("Lote / O.P.")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Automática")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (ProdutoApr.Name)
    objGridInt.colCampo.Add (DescricaoItemApr.Name)
    objGridInt.colCampo.Add (ApropItemProd.Name)
    objGridInt.colCampo.Add (ProdutoIns.Name)
    objGridInt.colCampo.Add (DescricaoItemIns.Name)
    objGridInt.colCampo.Add (UnidadeMedApr.Name)
    objGridInt.colCampo.Add (QuantidadeApr.Name)
    objGridInt.colCampo.Add (RequisicaoApr.Name)
    objGridInt.colCampo.Add (LoteApr.Name)
    objGridInt.colCampo.Add (FilialOpApr.Name)
    objGridInt.colCampo.Add (ApropAutomatica.Name)
    
    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_ProdutoApr_Col = 1
    iGrid_DescricaoApr_Col = 2
    iGrid_ApropItemProd_Col = 3
    iGrid_ProdutoIns_Col = 4
    iGrid_DescricaoIns_Col = 5
    iGrid_UnidadeMedApr_Col = 6
    iGrid_QuantidadeApr_Col = 7
    iGrid_RequisicaoApr_Col = 8
    iGrid_LoteApr_Col = 9
    iGrid_FilialOPApr_Col = 10
    iGrid_ApropAutomatica_Col = 11
    
    'Grid do GridInterno
    objGridInt.objGrid = GridApropriacao

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridApropriacao.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridApropriacao = SUCESSO

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no Banco de Dados
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objTipoMovEst As ClassTipoMovEst
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Tela_Extrai

    sTabela = "MovEstProd"
    
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
    'colSelecao.Add "TipoMov", OP_IGUAL, MOV_EST_PRODUCAO
    colSelecao.Add "NumIntDocEst", OP_IGUAL, 0

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165312)

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
    If lErro <> SUCESSO Then gError 30921

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 30921

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165313)

    End Select

    Exit Sub

End Sub

Function Preenche_Tela(objMovEstoque As ClassMovEstoque) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_Tela

    'Limpa a tela sem Fechar o Comando de setas
    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridApropriacao)
    
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
    lErro = CF("MovEstoqueItens_Le1", objMovEstoque, MOV_EST_PRODUCAO)
    If lErro <> SUCESSO And lErro <> 55387 Then gError 30922

    If lErro = 55387 Then gError 55403

    'Coloca os Dados na Tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    Codigo.PromptInclude = True

    Call DateParaMasked(Data, objMovEstoque.dtData)

    HORA.PromptInclude = False
    'este teste está correto
    If objMovEstoque.dtData <> DATA_NULA Then HORA.Text = Format(objMovEstoque.dtHora, "hh:mm:ss")
    HORA.PromptInclude = True

    lErro = Preenche_GridMovimentos(objMovEstoque.colItens)
    If lErro <> SUCESSO Then gError 30923
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovEstoque.colItens(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39634
    
    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objMovEstoque.objProjetoInfo, objMovEstoque.lCodigo, PRJ_CR_TIPO_PRODENTRADA, "", objMovEstoque.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 185844
    '###################################################

    lErro = CF2(Me, "ProdEnt_Preenche_Tela", objMovEstoque)
    If lErro <> SUCESSO Then gError 185844
    
    iAlterado = 0
    lCodigoAntigo = objMovEstoque.lCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 30922, 30923, 39634, 185844

        Case 55403
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_PRODUCAO", gErr, objMovEstoque.lCodigo)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165314)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMovimentos(colItens As colItensMovEstoque) As Long

Dim iIndice As Integer
Dim sProdutoMascarado As String, sCclMascarado As String
Dim lErro As Long
Dim objTipoMovEst As ClassTipoMovEst
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sContaEnxutaEst As String
Dim sContaEnxutaProducao As String
Dim colRatreamentoMovto As New Collection
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim objFilialEmpresa As New AdmFiliais
Dim colApropriacaoInsumo As Collection
Dim objProduto As New ClassProduto
Dim objItemMovEst As ClassItemMovEstoque
Dim colRastreamentoSerie As Collection
Dim iTipoMovto As Integer
Dim dSoma As Double

On Error GoTo Erro_Preenche_GridMovimentos

    Set gcolcolRastreamentoSerie = New Collection

    'Preenche GridMovimentos
    For Each objItemMovEstoque In colItens

        iIndice = iIndice + 1

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
            If lErro <> SUCESSO Then gError 49659
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
         'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then
        
            sContaEnxutaProducao = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then gError 52056
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxutaProducao
            ContaContabilProducao.PromptInclude = True
            
            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) = ContaContabilProducao.Text
            
        End If
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        GridMovimentos.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemMovEstoque.sProdutoDesc
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemMovEstoque.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed

        If objItemMovEstoque.iPossuiGrade = MARCADO Then GridMovimentos.TextMatrix(iIndice, 0) = "# " & GridMovimentos.TextMatrix(iIndice, 0)


        If objItemMovEstoque.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemMovEstoque.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 22921

        Else

            sCclMascarado = ""

        End If

        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado
        GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col) = objItemMovEstoque.sOPCodigo
        
        If objItemMovEstoque.lHorasMaquina > 0 Then
            GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col) = objItemMovEstoque.lHorasMaquina
            dSoma = dSoma + StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col))
        End If
        
        If objItemMovEstoque.dtDataInicioProducao <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataProducao_Col) = Format(objItemMovEstoque.dtDataInicioProducao, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataProducao_Col) = ""
        End If
        
        If objItemMovEstoque.lNumIntDocEst = 0 Then
            GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = CStr(DESMARCADO)
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = CStr(MARCADO)
        End If

        If objItemMovEstoque.iTipoMov = MOV_EST_PRODUCAO_BENEF3 Then
            iTipoMovto = BENEF_COMBO_BENEF
        ElseIf objItemMovEstoque.iTipoMov = MOV_EST_PRODUCAO_OUTROS Then
            iTipoMovto = BENEF_COMBO_OUTROS
        Else
            iTipoMovto = BENEF_COMBO_DISP
        End If
        
        Call Combo_Seleciona_ItemData(Benef, iTipoMovto)
        GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col) = Benef.Text
        
        Set colRatreamentoMovto = New Collection
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError 78422
        
        objProduto.sCodigo = objItemMovEstoque.sProduto
        
        'Lê o Produto para saber se é Lote ou se é Filial OP
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78621
        
        'Se existe rastreamento
        If colRatreamentoMovto.Count > 0 Then
                        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto(1)
            
            gcolcolRastreamentoSerie.Add objRatreamentoMovto.colRastreamentoSerie
            
            If objProduto.iRastro <> PRODUTO_RASTRO_OP Then
                If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridMovimentos.TextMatrix(iIndice, iGrid_Lote_Col) = objRatreamentoMovto.sLote
            End If
        
        Else
            Set colRastreamentoSerie = New Collection
            gcolcolRastreamentoSerie.Add colRastreamentoSerie
            
        End If
        
        Set colApropriacaoInsumo = New Collection
        
        'Le as Apriações do Item e preenche o Grid de Apropriacoes
        lErro = CF("ApropriacaoInsumo_Le_NumIntDocOrigem", objItemMovEstoque.lNumIntDoc, colApropriacaoInsumo)
        If lErro <> SUCESSO Then gError 78434
                        
        Set objItemMovEstoque.colApropriacaoInsumo = colApropriacaoInsumo
        
        lErro = Preenche_GridApropriacao(objItemMovEstoque, iIndice)
        If lErro <> SUCESSO Then gError 78435
        
        objItemMovEstoque.sSiglaUMEst = objProduto.sSiglaUMEstoque

        lErro = CF2(Me, "ProdEnt_Preenche_GridMovimentos", objItemMovEstoque, iIndice)
        If lErro <> SUCESSO Then gError 196433
        
    Next

    objGrid.iLinhasExistentes = colItens.Count

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 30925

    'Preenche Campo Horas Maquina
    If dSoma > 0 Then
        HorasMaquina.Text = Round(dSoma, 0)
    End If
    
    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 22921, 30925, 49659, 52056, 78422, 78423, 78434, 78435, 78621, 196433

        Case 30924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemMovEstoque.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165315)

    End Select

    Exit Function

End Function

Public Sub CodigoLabel_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    'Adiciona filtro
    'Janaina
    colSelecao.Add MOV_EST_PRODUCAO
    colSelecao.Add MOV_EST_PRODUCAO_BENEF3
    colSelecao.Add MOV_EST_PRODUCAO_OUTROS
    'Janaina
    
 'mario. colocar a selecao por BENEF3
    Call Chama_Tela("MovEstoqueLista1", colSelecao, objMovEstoque, objEventoCodigo)
   
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objMovEstoque As ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_objCodigoEvento_evSelecao

    Set objMovEstoque = obj1

    lErro = Preenche_Tela(objMovEstoque)
    If lErro <> SUCESSO Then gError 30926

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub

Erro_objCodigoEvento_evSelecao:

    Select Case gErr

        Case 30926

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165316)

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
                Parent.HelpContextID = IDH_ENTRADA_MATERIAL_PRODUZIDO_MOVIMENTOS
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ENTRADA_MATERIAL_PRODUZIDO_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Public Sub OPCodigoPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_OPCodigoPadrao_Validate

    'Verifica se foi Preenchido
    If Len(Trim(OPCodigoPadrao.Text)) = 0 Then Exit Sub

    objOrdemProducao.sCodigo = OPCodigoPadrao.Text
    objOrdemProducao.iFilialEmpresa = giFilialEmpresa

    lErro = CF("OrdemProducao_Le", objOrdemProducao)
    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 30944

    'se a OP não estiver cadastrada
    If lErro = 30368 Then gError 30945

    'se a ordem de producao está baixada
    If lErro = 55316 Then gError 55320
    
    'Alterado por Wagner 05/08/04
    'Alteração: Código comentado
    'Motivo: Função feita pelo botaoGeraGradeProducao
    'lErro = Traz_DadosOP_Tela(objOrdemProducao)
    'If lErro <> SUCESSO Then gError 30944
    'Fim da alteração


    Exit Sub

Erro_OPCodigoPadrao_Validate:

    Select Case gErr

        Case 30945
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemProducao.sCodigo)

        Case 55320
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165317)

    End Select

    Exit Sub

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
    
    Call OPCodigoPadrao_Validate(bSGECancelDummy)
    
    Me.Show

End Sub

Public Sub CclPadraoLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 30936

    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 30936
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165318)

    End Select

    Exit Sub

End Sub

Public Sub BotaoOP_Click()

Dim colSelecao As New Collection
Dim lErro As Long
Dim objItemOP As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoOP_Click

    If GridMovimentos.Row = 0 Then gError 54776
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 54777

    'Se o Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 54778

    colSelecao.Add sProdutoFormatado
                        
    Call Chama_Tela("ItemOP_ProdutoLista", colSelecao, objItemOP, objEventoOP)
    
    Exit Sub
    
Erro_BotaoOP_Click:

    Select Case gErr
        
        Case 54777
        
        Case 54778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
                         
        Case 54776
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165319)
        
    End Select
    
    Exit Sub
        
End Sub

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOP
Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_objEventoOP_evSelecao

    Set objItemOP = obj1

    If GridMovimentos.Row > 0 Then
    
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30927
    
        'Se Produto desta Linha Estiver Preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            'preenche o almoxarifado no grid a partir do item da OP
            lErro = Preenche_Almoxarifado(giFilialEmpresa, objItemOP.sCodigo, sProdutoFormatado)
            If lErro <> SUCESSO Then gError 55338
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col) = objItemOP.sCodigo
            
            OPCodigo.Text = objItemOP.sCodigo
            
            'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
            lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError 55337
    
        End If
    
    End If
    
    Me.Show
    
    Exit Sub

Erro_objEventoOP_evSelecao:

    Select Case gErr

        Case 30927, 55337, 55338

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165320)

    End Select

    Exit Sub

End Sub

Public Sub BotaoCcls_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) <> 0 And GridMovimentos.Row > 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 30935

        'Coloca o valor do Ccl na coluna correspondente
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 30935

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165321)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoProdutos_Click

    If GridMovimentos.Row = 0 Then gError 55341
    
    'Lista de produtos produzidos e inventariados
    Call Chama_Tela("ProdutoProduz_EstoqLista", colSelecao, objProduto, objEventoProduto)
    
   Exit Sub
   
Erro_BotaoProdutos_Click:

    Select Case gErr
        
        Case 55341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165322)
        
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
    
    If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 48725

    If GridMovimentos.Row = 0 Then gError 55339
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30930

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 55342

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 30931

    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 30932

    If lErro = 28030 Then gError 30933
    
    Produto.PromptInclude = False
    Produto.Text = sProdutoMascarado
    Produto.PromptInclude = True

    If Not (Me.ActiveControl Is Produto) Then

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = sProdutoMascarado
        
        'se a ContaContabilProdução não estiver preenchida em Produto procurar em TipoProduto
        If Trim(objProduto.sContaContabilProducao) = "" Then
            
            objTipoDeProduto.iTipo = objProduto.iTipo
            
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 52236
            
            If lErro = 22531 Then gError 52237
            
            objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
            
            'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
            If Trim(objProduto.sContaContabilProducao) = "" Then
                                
                objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
                objCTBConfig.iFilialEmpresa = giFilialEmpresa
                        
                lErro = CF("CTBConfig_Le", objCTBConfig)
                If lErro <> SUCESSO And lErro <> 9755 Then gError 52821
                
                If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
                
            End If
        
        End If
    
        'Preenche a Linha do Grid
        lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
        If lErro <> SUCESSO Then gError 30934
        
        'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
        lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
        If lErro <> SUCESSO Then gError 52234

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 30930, 30932, 30934, 52234, 52236, 52821

        Case 30931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
    
        Case 30933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 48725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, Produto.Text)
        
        Case 52237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
   
        Case 55339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
   
        Case 55342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, GridMovimentos.Row)
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165323)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRequisicoes_evSelecao(obj1 As Object)

Dim objMovEstoque As ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_objEventoRequisicoes_evSelecao

    Set objMovEstoque = obj1
    
    GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_RequisicaoApr_Col) = objMovEstoque.lCodigo

    Me.Show

    Exit Sub

Erro_objEventoRequisicoes_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165324)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAprop_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig
    
On Error GoTo Erro_objEventoProdutoAprop_evSelecao

    Set objProduto = obj1
    
    If GridApropriacao.Row = 0 Then gError 55339
    
    lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30930

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 30931

    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 30932

    If lErro = 28030 Then gError 30933
    
    If GridApropriacao.Col = iGrid_ProdutoApr_Col Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then gError 55342

        ProdutoApr.PromptInclude = False
        ProdutoApr.Text = sProdutoMascarado
        ProdutoApr.PromptInclude = True

        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoApr_Col) = sProdutoMascarado
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_DescricaoApr_Col) = objProduto.sDescricao
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_UnidadeMedApr_Col) = objProduto.sSiglaUMEstoque
    
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridApropriacao.Row - GridApropriacao.FixedRows) = objGridApropriacao.iLinhasExistentes Then
            objGridApropriacao.iLinhasExistentes = objGridApropriacao.iLinhasExistentes + 1
        End If
    
    ElseIf GridApropriacao.Col = iGrid_ProdutoIns_Col Then

        ProdutoIns.PromptInclude = False
        ProdutoIns.Text = sProdutoMascarado
        ProdutoIns.PromptInclude = True
        
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoIns_Col) = sProdutoMascarado
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_DescricaoIns_Col) = objProduto.sDescricao
        
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProdutoAprop_evSelecao:

    Select Case gErr

        Case 30930, 30932, 30934, 52234, 52236, 52821

        Case 30931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
    
        Case 30933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 48725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, Produto.Text)
        
        Case 52237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
   
        Case 55339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
   
        Case 55342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, GridApropriacao.Row)
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165325)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objMovEstoque Is Nothing Then

        'Lê MovEstoque no Banco de Dados
'hora

        objMovEstoque.iFilialEmpresa = giFilialEmpresa

        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 30939

        'Se o movimento existe
        If lErro = SUCESSO Then

            If objMovEstoque.iTipoMov <> MOV_EST_PRODUCAO And objMovEstoque.iTipoMov <> MOV_EST_PRODUCAO_BENEF3 And objMovEstoque.iTipoMov <> MOV_EST_PRODUCAO_OUTROS Then gError 30940

            lErro = Preenche_Tela(objMovEstoque)
            If lErro <> SUCESSO Then gError 30941

        Else
            'Se ele não existe exibe apenas o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objMovEstoque.lCodigo)
            Codigo.PromptInclude = True

            lCodigoAntigo = objMovEstoque.lCodigo

        End If

    Else

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 30939, 30941

        Case 30940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOV_EST_NAO_PRODUCAO", gErr, objMovEstoque.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165326)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub botaoOPCodigoPadrao_Click()
'Incluido por Wagner 05/08/04
'Alteração: Botão para trazer Ordens de Produção
'Motivo: Especificação do cliente

Dim lErro As Long
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_botaoOPCodigoPadrao_Click

    If Len(Trim(OPCodigoPadrao.Text)) <= 0 Then Exit Sub
    
    objOrdemProducao.sCodigo = OPCodigoPadrao.Text
    objOrdemProducao.iFilialEmpresa = giFilialEmpresa

    lErro = CF("OrdemProducao_Le", objOrdemProducao)
    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 30944

    lErro = Traz_DadosOP_Tela(objOrdemProducao)
    If lErro <> SUCESSO Then gError 30944

    Exit Sub

Erro_botaoOPCodigoPadrao_Click:

    Select Case gErr

        Case 30944

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165327)

    End Select

    Exit Sub

End Sub

Public Sub CclPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    'Verifica se o CclPadrao foi Preenchida
    If Len(Trim(CclPadrao.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 30943

        'se o ccl não estiver cadastrado
        If lErro = 5703 Then gError 30944

    End If

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 30943

        Case 30944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165328)

    End Select

    Exit Sub

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 30946

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 30946

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165329)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165330)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 30947

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 30947

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165331)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 30948

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 30948

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165332)

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
Dim iProdutoPreenchido As Integer, lNumIntDoc As Long
Dim sCodProduto2 As String
Dim sProdutoFormatado2 As String
Dim iProdutoPreenchido2 As Integer
Dim iApropAutomatica As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        'Verifica se produto está preenchido
        sCodProduto = GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
    
        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30949
    
        If gobjMovEst.colItens.Count >= GridMovimentos.Row And GridMovimentos.Row > 0 Then
            lNumIntDoc = gobjMovEst.colItens(GridMovimentos.Row).lNumIntDoc
        Else
            lNumIntDoc = 0
        End If
    
        'Verifica se produto está preenchido no Grid de Apropriações
        sCodProduto2 = GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoApr_Col)
    
        lErro = CF("Produto_Formata", sCodProduto2, sProdutoFormatado2, iProdutoPreenchido2)
        If lErro <> SUCESSO Then gError 78152
    
        If iProdutoPreenchido2 = PRODUTO_PREENCHIDO And GridApropriacao.Row > 0 Then
        
            iApropAutomatica = IIf(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ApropAutomatica_Col) = CStr(MARCADO), APROPINSUMOSPROD_AUTOMATICO, 0)
        
        Else
        
            iApropAutomatica = 0
            
        End If
        
        If objControl.Name = "Produto" Then
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
    
            End If
    
        ElseIf objControl.Name = "UnidadeMed" Then
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Or lNumIntDoc <> 0 Or left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) = "#" Then
                
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
    
                objProduto.sCodigo = sProdutoFormatado
    
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 30951
    
                If lErro = 28030 Then gError 30952
    
                objClasseUM.iClasse = objProduto.iClasseUM
    
                'Preenche a List da Combo UnidadeMed com as UM's do Produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 30953
    
                'Guardo o valor da Unidade de Medida da Linha
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
    
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
    
        ElseIf objControl.Name = "Quantidade" Or objControl.Name = "OPCodigo" Or objControl.Name = "Almoxarifado" Then
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 And left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        ElseIf objControl.Name = "Ccl" Or objControl.Name = "ContaContabilEst" Or objControl.Name = "ContaContabilProducao" Then
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        ElseIf objControl.Name = "HorasMaq" Then
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        ElseIf objControl.Name = "DataProducao" Then
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        ElseIf objControl.Name = "Lote" Then
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
            
                objProduto.sCodigo = sProdutoFormatado
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78213
        
                If lErro = 28030 Then gError 78214
            
                If objProduto.iRastro <> PRODUTO_RASTRO_OP Then 'Alterado por Wagner 14/03/2006
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            Else
                objControl.Enabled = False
            End If
        
        ElseIf objControl.Name = "Estorno" Then
        
            If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        ElseIf objControl.Name = "ProdutoApr" Then
    
            If iProdutoPreenchido2 = PRODUTO_PREENCHIDO Or iApropAutomatica <> 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
    
        ElseIf objControl.Name = "UnidadeMedApr" Then
    
            If iProdutoPreenchido2 <> PRODUTO_PREENCHIDO Or iApropAutomatica <> 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
    
                objProduto.sCodigo = sProdutoFormatado2
    
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO Then gError 78150
    
                objClasseUM.iClasse = objProduto.iClasseUM
    
                'Preenche a List da Combo UnidadeMed com as UM's do Produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 78151
    
                'Guardo o valor da Unidade de Medida da Linha
                sUnidadeMed = GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_UnidadeMedApr_Col)
    
                'Limpar as Unidades utilizadas anteriormente
                UnidadeMedApr.Clear
    
                For Each objUnidadeDeMedida In colSiglas
                    UnidadeMedApr.AddItem objUnidadeDeMedida.sSigla
                Next
    
                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMedApr.ListCount <> 0 Then
    
                    For iIndice = 0 To UnidadeMedApr.ListCount - 1
    
                        If UnidadeMedApr.List(iIndice) = sUnidadeMed Then
                            UnidadeMedApr.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
            End If
    
        ElseIf objControl.Name = "QuantidadeApr" Or objControl.Name = "RequisicaoApr" Then
    
            If iProdutoPreenchido2 = PRODUTO_PREENCHIDO And iApropAutomatica = 0 Then
                
                objControl.Enabled = True
    
            Else
                objControl.Enabled = False
    
            End If
            
        ElseIf objControl.Name = "ApropItemProd" Then
        
            If iProdutoPreenchido2 = PRODUTO_PREENCHIDO And iApropAutomatica = 0 Then
                
                objControl.Enabled = True
    
            Else
                objControl.Enabled = False
    
            End If
        
        ElseIf objControl.Name = "LoteApr" Then
    
            If iProdutoPreenchido2 = PRODUTO_PREENCHIDO And iApropAutomatica = 0 Then
            
                objProduto.sCodigo = sProdutoFormatado2
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78340
        
                If lErro = 28030 Then gError 78341
            
                If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
            Else
                objControl.Enabled = False
            End If
        
        ElseIf objControl.Name = "FilialOpApr" Then
    
            If iProdutoPreenchido2 = PRODUTO_PREENCHIDO And iApropAutomatica = 0 Then
            
                objProduto.sCodigo = sProdutoFormatado2
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 78342
        
                If lErro = 28030 Then gError 78343
            
                If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            Else
                objControl.Enabled = False
            End If
    
        ElseIf objControl.Name = "Benef" Then
        
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                
        ElseIf objControl.Name = "ApropAutomatica" Or objControl.Name = "ProdutoIns" Then
        
            objControl.Enabled = False
        
        Else
        
            lErro = CF2(Me, "ProdEnt_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada, iProdutoPreenchido, sProdutoFormatado, lNumIntDoc)
            If lErro <> SUCESSO Then gError 196432
    
        End If

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 30949, 30951, 30952, 30953, 78150, 78151, 78152, 78213, 78214, 78315, 78316, 78340, 78341, 78342, 78343, 196432

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165333)

    End Select

    Exit Sub

End Sub

Public Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer, lNumIntDoc As Long
Dim iLinhaAnterior As Integer
Dim iLinhasExistentes As Integer 'm
Dim lErro As Long

On Error GoTo Erro_GridMovimentos_KeyDown

    If gobjMovEst.colItens.Count >= GridMovimentos.Row Then
        lNumIntDoc = gobjMovEst.colItens(GridMovimentos.Row).lNumIntDoc
    Else
        lNumIntDoc = 0
    End If

    If lNumIntDoc = 0 Then

        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then
        
            'Guarda iLinhasExistentes
            iLinhasExistentesAnterior = objGrid.iLinhasExistentes
    
            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridMovimentos.Row
    
        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then
            
            'Recalcula as colunas de HorasMaquina do grid
            lErro = Calcula_Rateio_HorasMaquina()
            If lErro <> SUCESSO Then gError 87930
        
            'Exclui de colItens o Item correspondente, se houver
            gobjMovEst.colItens.Remove iLinhaAnterior
            gcolcolRastreamentoSerie.Remove iLinhaAnterior 'Inserido por Wagner 14/03/2006

            Call gobjTelaProjetoInfo.Remove_Linha(iLinhaAnterior)

            For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
                If gobjMovEst.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
                Else
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
                End If
                
            Next

            GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

            'ajustar grid de apropriacoes
            Call GridAprop_Exclui_ItemMov(iLinhaAnterior)

        End If

    End If
    
    Exit Sub
    
Erro_GridMovimentos_KeyDown:

    Select Case gErr
    
        Case 87930

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165334)

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
Dim colItensRomaneioGrade As New Collection
Dim objItemMovEst As New ClassItemMovEstoque
Dim objRomaneioGrade As New ClassRomaneioGrade

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 117643

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError 126490
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 30956
            
            lErro = Obtem_Conta_Producao(objProduto)
            If lErro <> SUCESSO Then gError 190829

'            'se o conta de produção não estiver associada ao produto
'            If Trim(objProduto.sContaContabilProducao) = "" Then
'
'                objTipoDeProduto.iTipo = objProduto.iTipo
'
'                'pesquisa a conta junto ao tipo do produto
'                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
'                If lErro <> SUCESSO And lErro <> 22531 Then gError 52238
'
'                If lErro = 22531 Then gError 52239
'
'                objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
'
'                'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
'                If Trim(objProduto.sContaContabilProducao) = "" Then
'
'                    objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
'                    objCTBConfig.iFilialEmpresa = giFilialEmpresa
'
'                    lErro = CF("CTBConfig_Le", objCTBConfig)
'                    If lErro <> SUCESSO And lErro <> 9755 Then gError 52822
'
'                    If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
'
'                End If
'
'            End If
                
            lErro = CF("Customiza_ProducaoEnt_SaidaCel", objProduto)
            If lErro <> SUCESSO Then gError 117694
                
            'se nao for um produto pai de grade
            If Len(Trim(objProduto.sGrade)) = 0 Then
                
                If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 30959
                    
                If objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 48724
                    
                ' preenche a linha do produto
                lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
                If lErro <> SUCESSO Then gError 30958
                
                'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
                lErro = Preenche_ContaContabilEst(Produto.Text)
                If lErro <> SUCESSO Then gError 52235
    
            'se for um produto pai de grade
            Else
            
                'Verifica se há filhos válidos da grade pai
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 117644
                
                'Se nao existir, erro
                If colItensRomaneioGrade.Count = 0 Then gError 117645
                
                objItemMovEst.sProduto = objProduto.sCodigo
                objItemMovEst.sDescricao = objProduto.sDescricao
                objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
                objItemMovEst.sOPCodigo = OPCodigoPadrao.Text
                        
                Set objRomaneioGrade = New ClassRomaneioGrade
                
                objRomaneioGrade.sNomeTela = Me.Name
                Set objRomaneioGrade.objTela = Me
                
                Set objRomaneioGrade.objObjetoTela = objItemMovEst
                            
                Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
                If giRetornoTela <> vbOK Then gError 117646
        
                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
                If lErro <> SUCESSO Then gError 117647
        
                lErro = Preenche_ContaContabilEst(Produto.Text)
                If lErro <> SUCESSO Then gError 117648
        
                Call Atualiza_Grid_Movimentos(objItemMovEst)
            
            End If
    
        End If

    Else
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEst_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducao_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30957

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 30955, 30957, 30958, 52235, 52238, 52822, 117643, 117644, 117646, 117647, 117648, 117694, 190829
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30956
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 30959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 48724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case 52239
             lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case 117645, 126490 'Alterado por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165335)

    End Select

    Exit Function

End Function

Public Sub BotaoGrade_Click()

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_BotaoGrade_Click

    If GridMovimentos.Row > 0 And GridMovimentos.Row <= objGrid.iLinhasExistentes Then
    
        Set objItemMovEst = gobjMovEst.colItens(GridMovimentos.Row)
        
        'If objItemMovEst.iPossuiGrade = DESMARCADO Then gError 117649
        
        If objItemMovEst.iPossuiGrade = MARCADO Then
        
            objItemMovEst.sOPCodigo = OPCodigoPadrao.Text
            objItemMovEst.sDescricao = objItemMovEst.sProdutoDesc
            objItemMovEst.sSiglaUMEst = objItemMovEst.sSiglaUM
    
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            Set objRomaneioGrade.objTela = Me
            
            Set objRomaneioGrade.objObjetoTela = objItemMovEst
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        
            Call Atualiza_Grid_Movimentos(objItemMovEst)
            
        End If
            
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
      
'        Case 117649
'            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_GRADE", gErr, GridMovimentos.Row)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165336)
            
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

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemMovEst.dQuantidade = dQuantidade

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemMovEst As ClassItemMovEstoque) As Long

Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sContaEnxuta As String
Dim sAlmoxarifadoPadrao As String
Dim objItemOP As New ClassItemOP
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colRastreamentoSerie As New Collection 'Inserido por Wagner 14/03/2006

On Error GoTo Erro_ProdutoLinha_Preenche
  
    If Trim(objProduto.sContaContabilProducao) <> "" Then
    
        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
        If lErro <> SUCESSO Then gError 52038
    
        'preenche  a ContaContabilProducao
        ContaContabilEst.PromptInclude = False
        ContaContabilEst.Text = sContaEnxuta
        ContaContabilEst.PromptInclude = True
    
        'preenche Conta De Producao
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducao_Col) = ContaContabilEst.Text
    
    End If
    
    Call Combo_Seleciona_ItemData(Benef, BENEF_COMBO_DISP)
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col) = Benef.Text
    
    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Descricao_Col) = objProduto.sDescricao
    
    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 30960

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = CclPadrao.Text

    If Len(Trim(objProduto.sGrade)) = 0 Then

        'se o codigo padrao estiver preenchido e nao for produto pai de grade
        If Len(Trim(OPCodigoPadrao.Text)) <> 0 And Len(Trim(objProduto.sGrade)) = 0 Then
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_OP_Col) = OPCodigoPadrao.Text
            
            'preenche o almoxarifado no grid a partir do item da OP
            lErro = Preenche_Almoxarifado(giFilialEmpresa, OPCodigoPadrao.Text, objProduto.sCodigo)
            If lErro <> SUCESSO Then gError 55333
    
        End If

    End If

    'Preenche Estorno com Valor 0 (Checked = False)
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Estorno_Col) = "0"

    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        gcolcolRastreamentoSerie.Add colRastreamentoSerie 'Inserido por Wagner 14/03/2006

        Call gobjTelaProjetoInfo.Adiciona_Linha(GridMovimentos.Row)

        gobjMovEst.colItens.Add1 objItemMovEst
    
        If Len(Trim(objProduto.sGrade)) = 0 Then

            objItemMovEst.iPossuiGrade = DESMARCADO
            
        Else
        
            objItemMovEst.iPossuiGrade = MARCADO
            GridMovimentos.TextMatrix(GridMovimentos.Row, 0) = "# " & GridMovimentos.TextMatrix(GridMovimentos.Row, 0)
        
        End If
        
        objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
        objItemMovEst.sProduto = objProduto.sCodigo
    
    End If
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 30960, 49674, 49675, 52038, 52233, 55333

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165337)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim vbMsg As VbMsgBoxResult
Dim sPercentual As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 30961

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 126801

    'William - 24/04/01
    If Len(Trim(OPCodigoPadrao.Text)) > 0 Then

        objOrdemProducao.sCodigo = OPCodigoPadrao.Text
        objOrdemProducao.iFilialEmpresa = giFilialEmpresa

        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> AD_SQL_SUCESSO Then gError 87562

        For Each objItemOP In objOrdemProducao.colItens

            If objItemOP.sProduto = sProdutoFormatado Then

                If StrParaDbl(Quantidade.Text) > objItemOP.dQuantidade Then
                
                    lErro = Retorna_Percentagem(objItemOP.dQuantidade, StrParaDbl(Quantidade.Text), sPercentual)
                    If lErro <> SUCESSO Then gError 87566
                            
                    vbMsg = Rotina_Aviso(vbOKOnly, "AVISO_QUANTID_MAIOR_OP", sPercentual)
                    
                End If
    
            End If
    
        Next
    
    End If

    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(Quantidade.Text), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141911

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30962
    
    'Verifica se Campo HorasMaquina foi preenchido
    If Len(Trim(HorasMaquina.Text)) > 0 Then
    
        'Preenche as colunas HorasMaquina do grid
        lErro = Calcula_Rateio_HorasMaquina()
        If lErro <> SUCESSO Then gError 87892
                                                
    End If

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 30961, 30962, 87562, 87566, 87892, 126801, 141911
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165338)
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

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 79761

        'Valida o ALmoxarifado
        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 79762
        
        'Se não for encontrado --> Erro
        If lErro = 25157 Then gError 79763
        If lErro = 25162 Then gError 79764

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 79765

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 79765, 79761, 79762
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 79763
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 79764

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(Almoxarifado.Text))
            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165339)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HorasMaq(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_HorasMaq

    Set objGridInt.objControle = HorasMaq

    If Len(Trim(HorasMaq.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(HorasMaq.Text)
        If lErro <> SUCESSO Then gError 78136

        HorasMaq.Text = Format(HorasMaq.Text, "Standard")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78137

    Saida_Celula_HorasMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_HorasMaq:

    Saida_Celula_HorasMaq = gErr

    Select Case gErr

        Case 78136, 78137
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165340)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote

    '##########################################################
    'Inserido por Wagner 14/03/2006
    'Carrega as séries na coleção global
    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(GridMovimentos.Row), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), Lote.Text, StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col), GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141912
    '##########################################################

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78148

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 78148, 141912
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165341)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_OP(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOP
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objProduto As New ClassProduto
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Saida_Celula_OP

    Set objGridInt.objControle = OPCodigo

    If Len(Trim(OPCodigo.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = OPCodigo.Text

        lErro = CF("OrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 30968

        If lErro = 30368 Then gError 30969

        'ordem de producao baixada
        If lErro = 55316 Then gError 55321

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30970

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            'preenche o almoxarifado no grid a partir do item da OP
            lErro = Preenche_Almoxarifado(giFilialEmpresa, OPCodigo.Text, sProdutoFormatado)
            If lErro <> SUCESSO Then gError 55330
        
            'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
            lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))
            If lErro <> SUCESSO Then gError 49707
            
            objItemOP.sProduto = sProdutoFormatado
            objItemOP.sCodigo = objOrdemProducao.sCodigo
            objItemOP.iFilialEmpresa = giFilialEmpresa
            
            'Le a data de Previsão do Inicio da Producao
            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 78450
                        
            If objItemOP.dtDataInicioProd <> DATA_NULA Then
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataProducao_Col) = objItemOP.dtDataInicioProd
            Else
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataProducao_Col) = ""
            End If
            
        End If
        
'        lErro = CF("Produto_Formata",GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 78518
'
'        objProduto.sCodigo = sProdutoFormatado
'
'        lErro = CF("Produto_Le",objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 78519
'
'        'Produto não encontrado
'        If lErro = 28030 Then gError 78520
'
'        If iProdutoPreenchido = PRODUTO_PREENCHIDO And objProduto.iRastro = PRODUTO_RASTRO_OP Then
'
'            objRastroLote.sCodigo = OPCodigo.Text
'            objRastroLote.sProduto = sProdutoFormatado
'            objRastroLote.iFilialOP = giFilialEmpresa
'
'            'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
'            lErro = CF("RastreamentoLote_Le",objRastroLote)
'            If lErro <> SUCESSO And lErro <> 75710 Then gError 78521
'
'            'Se não encontrou --> Erro
'            If lErro = 75710 Then gError 78522
'
'        End If

    Else

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30973

    Saida_Celula_OP = SUCESSO

    Exit Function

Erro_Saida_Celula_OP:

    Saida_Celula_OP = gErr

    Select Case gErr

        Case 30968, 30970, 30973, 49707, 55330, 78450, 78514, 78516, 78519, 78521
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30969
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OP_NAO_CADASTRADA", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("OrdemProducao", objOrdemProducao)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 55321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78520
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsg = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 78522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165342)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Trim(Ccl.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 30974

        If lErro = 5703 Then gError 30975

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30976

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 30974, 30976
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 30975
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            If vbMsg = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165343)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 30977

    Call Limpa_Tela_ProducaoEntrada

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 30977

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165344)

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
Dim sCodigoOP As String, bGravaSemReq As Boolean
Dim objItemMovEst As ClassItemMovEstoque
Dim aiItens() As Integer
Dim iCont As Integer
Dim sItensSemReq As String
Dim objItemOP As ClassItemOP
Dim sProdMaiorReq As String
Dim dSoma As Double
Dim dSomaTotal As Double

On Error GoTo Erro_Gravar_Registro

    ReDim aiItens(1 To NUM_MAX_ITENS_MOV_ESTOQUE)

    bGravaSemReq = False
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 30979

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 30984

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovimentos
    If objGrid.iLinhasExistentes = 0 Then gError 30985

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 30986

        'Verifica se OPCodigo foi preenchido
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col))) = 0 And left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then gError 30988

        If bGravaSemReq = False And left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then

            Set objItemMovEst = New ClassItemMovEstoque
    
            'Coloca o Produto no formato do BD
            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 138970
    
            objItemMovEst.sOPCodigo = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
            objItemMovEst.sProduto = sProdutoFormatado
    
            lErro = CF("ReqProducao_TestaExistencia", objItemMovEst)
            If lErro <> SUCESSO And lErro <> 71417 Then gError 71413
    
            'Acumula linhas sem Req. Prod
            If lErro = 71417 Then
                iCont = iCont + 1
                aiItens(iCont) = iIndice
            End If
        
            Set objItemOP = New ClassItemOP
            
            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sProduto = sProdutoFormatado
            objItemOP.sCodigo = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
            objItemOP.dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
            objItemOP.sSiglaUM = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Testa_Quant_Requerida", objItemOP, sProdMaiorReq)
            If lErro <> SUCESSO Then gError 126629
        
        End If
        
    Next

    For iIndice = 1 To iCont
        'Se é o primeito
        If iIndice = 1 Then
            sItensSemReq = CStr(aiItens(iIndice))
        Else
            'Se é o último
            If iIndice = iCont Then
                sItensSemReq = sItensSemReq & " e " & CStr(aiItens(iIndice))
            Else
                sItensSemReq = sItensSemReq & ", " & CStr(aiItens(iIndice))
            End If
        End If
    Next
    
    'Se tem aviso => Exibe
    If Len(Trim(sItensSemReq)) > 0 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NAO_EXISTE_REQPRODUCAO_OP", sItensSemReq, sCodigoOP)
        If vbMsgRes = vbNo Then gError 71418
    End If
    
    'Se tem aviso => Exibe
    If Len(Trim(sProdMaiorReq)) > 0 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUCAO_MAIOR_REQUISICAO", sProdMaiorReq)
        If vbMsgRes = vbNo Then gError 126629
    End If
    
    
    objMovEstoque.lCodigo = CLng(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa

    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 30981
    
    If lErro = SUCESSO Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS")
        If vbMsgRes = vbNo Then gError 78622
    
    End If
    
    'Verifica se HorasMaq é igual a HorasMaquina
    If Len(Trim(HorasMaquina.Text)) > 0 Then
    
        For iIndice = 1 To objGrid.iLinhasExistentes
            dSoma = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col))
            dSomaTotal = dSoma + dSomaTotal
        Next
    
        If StrParaDbl(HorasMaquina.Text) <> dSomaTotal Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HORASMAQUINA_DIF_HORASMAQ")
            If vbMsgRes = vbNo Then gError 87930
    
        End If
    
    End If
    
    'Verifica se todos os Produtos insumos do Grid de Apropriação estão no grid de Movimentos
    lErro = Apropriacao_Testa_Insumo()
    If lErro <> SUCESSO Then gError 78351
           
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 92028
           
    'Move os dados para o objMovimentoEstoque
    lErro = Move_Tela_Memoria(objMovEstoque)
    If lErro <> SUCESSO Then gError 30989
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 185842
    
    'Grava no BD(inclusive os dados contabeis)
    lErro = CF("MovEstoque_Producao_Grava", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 30990

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_MOVESTOQUE
            gobjAnotacao.sID = CStr(objMovEstoque.iFilialEmpresa) & "," & CStr(objMovEstoque.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 30990
            
        End If
        
    End If
    
    If gobjEST.iProdEntVerificaBxOP = MARCADO Then
        'Verifica a existencia no Grid de Itens de Ordem de Produção que não deram baixa
        lErro = ItensOrdemProducao_Verifica()
        If lErro <> SUCESSO Then gError 93649
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
       
        Case 30979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 30984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 30985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 30986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 30988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREENCHIDO1", gErr, iIndice)

        Case 30989, 30990, 30395, 30981, 78351, 87930, 185842
        
        Case 78622, 71413, 71418, 92028, 138970, 126629
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165345)

    End Select

    Exit Function

End Function

Function Apropriacao_Testa_Insumo() As Long
'Verifica se todos os Produtos insumos do Grid de Apropriação estão no grid de Movimentos

Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iEncontrou As Integer

On Error GoTo Erro_Apropriacao_Testa_Insumo

    For iIndice1 = 1 To objGridApropriacao.iLinhasExistentes
    
        iEncontrou = 0
        
        'Procura o Produto Insumo no Grid de Movimentos
        For iIndice2 = 1 To objGrid.iLinhasExistentes
            
            If GridMovimentos.TextMatrix(iIndice2, iGrid_Produto_Col) = GridApropriacao.TextMatrix(iIndice1, iGrid_ProdutoIns_Col) Then
                iEncontrou = 1
                Exit For
            End If
        
        Next
        
        'Se não encontrou --> Erro
        If iEncontrou = 0 Then gError 78350
        
    Next
    
    Exit Function
    
    Apropriacao_Testa_Insumo = SUCESSO
    
Erro_Apropriacao_Testa_Insumo:

    Apropriacao_Testa_Insumo = gErr
    
    Select Case gErr
        
        Case 78350
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOINSUMO_NAO_ENCONTRADO_GRIDMOVIMENTOS", gErr, GridApropriacao.TextMatrix(iIndice1, iGrid_ProdutoIns_Col), iIndice1)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165346)

    End Select

    Exit Function
        
End Function

Function Move_Tela_Memoria(objMovEstoque As ClassMovEstoque, Optional bRelRotulo As Boolean = False) As Long
'Preenche objMovEstoque (inclusive colItens)

Dim iIndice As Integer
Dim lCodigo As Long
Dim lErro As Long
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria
    
    'Carrega o Código
    If Len(Trim(Codigo.Text)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)
    
    'Carrega a Data
    If Len(Trim(Data.Text)) <> 0 Then
        objMovEstoque.dtData = CDate(Data.Text)
    Else
        objMovEstoque.dtData = DATA_NULA
    End If
    
'hora
    If Len(Trim(HORA.ClipText)) > 0 Then
        objMovEstoque.dtHora = CDate(HORA.Text)
    Else
        objMovEstoque.dtHora = Time
    End If
    
    'A Filial Empresa
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    'Varre o Grid de Itens de Movimentos
    For iIndice = 1 To objGrid.iLinhasExistentes
                
        'Pega todos os Itens mesmo que seja extorno
        lErro = Move_Itens_Memoria(iIndice, objMovEstoque, bRelRotulo)
        If lErro <> SUCESSO Then gError 30993

    Next
    
    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_PRODENTRADA)
    If lErro <> SUCESSO Then gError 185843
    
    Set objMovEstoque.objProjetoInfo = objProjetoInfo
    '##################################################

    lErro = CF2(Me, "ProdEnt_Move_Tela_Memoria", objMovEstoque)
    If lErro <> SUCESSO Then gError 185843
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 30992, 30993, 30991, 185843

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165347)

    End Select

    Exit Function

End Function

Function Move_Itens_Memoria(iIndice As Integer, objMovEstoque As ClassMovEstoque, Optional bRelRotulo As Boolean = False) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sCclFormatada As String, sCcl As String
Dim iCclPreenchida As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaFormatadaEst As String
Dim iContaPreenchida As Integer
Dim sContaFormatadaProducao As String
Dim colApropriacaoInsumos As New Collection
Dim colRastreamento As New Collection
Dim dtDataProducao As Date
Dim iTipoMovEstoque As Integer
Dim objItemMovEst As ClassItemMovEstoque
Dim lBenef As Long

On Error GoTo Erro_Move_Itens_Memoria

    With GridMovimentos
         
        Call Combo_Obtem_ItemData(Benef, .TextMatrix(iIndice, iGrid_Benef_Col), lBenef)
         
        'Se o Item for extorno
        If GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = CStr(MARCADO) Then
            
            'verifica se o lNumIntDoc do elemento correspondente em colItens está zerado
            If gobjMovEst.colItens(iIndice).lNumIntDoc = 0 Then gError 126676
            
            'Seta o Tipo como extorno
            If lBenef = BENEF_COMBO_BENEF Then
                iTipoMovEstoque = MOV_EST_ESTORNO_PRODUCAO_BENEF3
            ElseIf lBenef = BENEF_COMBO_OUTROS Then
                iTipoMovEstoque = MOV_EST_ESTORNO_PRODUCAO_OUTROS
            Else
                iTipoMovEstoque = MOV_EST_ESTORNO_PRODUCAO
            End If
         
        Else
            If lBenef = BENEF_COMBO_BENEF Then
                iTipoMovEstoque = MOV_EST_PRODUCAO_BENEF3
            ElseIf lBenef = BENEF_COMBO_OUTROS Then
                iTipoMovEstoque = MOV_EST_PRODUCAO_OUTROS
            Else
                iTipoMovEstoque = MOV_EST_PRODUCAO
            End If
        End If


        If left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then
        
            objAlmoxarifado.sNomeReduzido = .TextMatrix(iIndice, iGrid_Almoxarifado_Col)
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25060 Then gError 30994
    
            If lErro = 25060 Then gError 30995
       
        End If
       
        sCcl = .TextMatrix(iIndice, iGrid_Ccl_Col)
        
        If Len(Trim(sCcl)) <> 0 Then
        
            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", .TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 30996
        
        Else
        
            sCclFormatada = ""

        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilEst_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilEst_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 49657
        
        Else
            sContaFormatadaEst = ""
        End If
        
        If .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", .TextMatrix(iIndice, iGrid_ContaContabilProducao_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then gError 52053
        
        Else
            sContaFormatadaProducao = ""
        End If
        
        'Formata o Produto para BD
        sProdutoFormatado = ""
        lErro = CF("Produto_Formata", .TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 30997
        
        Set colApropriacaoInsumos = New Collection
        
        If Not bRelRotulo Then
        
            'Move os dados do Grid de Insumos para a memória
            lErro = Move_ApropriacaoInsumos_Memoria2(sProdutoFormatado, colApropriacaoInsumos, iIndice)
            If lErro <> SUCESSO Then gError 78173
        
        End If
            
        Set colRastreamento = New Collection
        
        'Move o Rastro para a Memória
        lErro = Move_RastroEstoque_Memoria(iIndice, colRastreamento)
        If lErro <> SUCESSO Then gError 78174
        
        If Len(Trim(DataProducao.ClipText)) > 0 Then
            dtDataProducao = CDate(DataProducao.Text)
        Else
            dtDataProducao = DATA_NULA
        End If
                
        Set objItemMovEst = objMovEstoque.colItens.Add(gobjMovEst.colItens(iIndice).lNumIntDoc, iTipoMovEstoque, 0, 0, sProdutoFormatado, .TextMatrix(iIndice, iGrid_Descricao_Col), .TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(.TextMatrix(iIndice, iGrid_Quantidade_Col)), objAlmoxarifado.iCodigo, .TextMatrix(iIndice, iGrid_Almoxarifado_Col), 0, sCclFormatada, StrParaLong(.TextMatrix(iIndice, iGrid_Estorno_Col)), .TextMatrix(iIndice, iGrid_OP_Col), "", sContaFormatadaProducao, sContaFormatadaEst, StrParaLong(.TextMatrix(iIndice, iGrid_HorasMaquina_Col)), colRastreamento, colApropriacaoInsumos, dtDataProducao)
        
        Set objItemMovEst.colItensRomaneioGrade = gobjMovEst.colItens(iIndice).colItensRomaneioGrade
        
    End With

    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 30994, 30996, 30997, 49657, 52053, 78173, 78174

        Case 30995
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 126676
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTORNO_ITEM_NAO_CADASTRADO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165348)

    End Select

    Exit Function

End Function

Function Move_ApropriacaoInsumos_Memoria(sProdutoInsumo As String, colApropriacaoInsumos As Collection) As Long
'??? apenas para nao perder a compatibilidade binaria

    Move_ApropriacaoInsumos_Memoria = SUCESSO
    
End Function

Private Function Move_ApropriacaoInsumos_Memoria2(sProdutoInsumo As String, colApropriacaoInsumos As Collection, ByVal iItem As Integer) As Long
'Move o Grid de insumos para a memória

Dim lErro As Long
Dim iIndice As Integer, iItemGrid As Integer
Dim objApropriacaoInsumos As New ClassApropriacaoInsumosProd
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatadoGridInsumo As String
Dim iProdutoPreenchidoGridInsumo As Integer
Dim objRequisicao As New ClassItemMovEstoque

On Error GoTo Erro_Move_ApropriacaoInsumos_Memoria2
    
    'Varre o Grid de Apropriação pegando os Produtos que tem vinculo com o Produto Insumo passado
    For iIndice = 1 To objGridApropriacao.iLinhasExistentes
                
        iItemGrid = StrParaInt(GridApropriacao.TextMatrix(iIndice, iGrid_ApropItemProd_Col))
        
        If iItemGrid = iItem Then
        
            lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(iIndice, iGrid_ProdutoIns_Col), sProdutoFormatadoGridInsumo, iProdutoPreenchidoGridInsumo)
            If lErro <> SUCESSO Then gError 78168
            
            'Se o produto insumo não foi preenchido --> erro
            If iProdutoPreenchidoGridInsumo = 0 Then gError 78169
                 
            'Se for igual
            If sProdutoFormatadoGridInsumo = sProdutoInsumo Then
                        
                Set objApropriacaoInsumos = New ClassApropriacaoInsumosProd
                
                lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(iIndice, iGrid_ProdutoApr_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 78170
                
                'Se o produto não foi preenchido --> erro
                If iProdutoPreenchido = 0 Then gError 78171
                
                objApropriacaoInsumos.sProduto = sProdutoFormatado
                
                'Se houve Requisição pega o NumIntDoc da Requisicao
                If Len(Trim(GridApropriacao.TextMatrix(iIndice, iGrid_RequisicaoApr_Col))) > 0 Then
            
                    objRequisicao.lCodigo = CLng(GridApropriacao.TextMatrix(iIndice, iGrid_RequisicaoApr_Col))
                    objRequisicao.sProduto = sProdutoFormatado
                    objRequisicao.iFilialEmpresa = giFilialEmpresa
                    
                    'Le a Requisição para pegar o NumIntDoc da Requisição
                    lErro = CF("MovEstoque_Le_Produto", objRequisicao)
                    If lErro <> SUCESSO And lErro <> 78189 Then gError 78172
                    
                    'não encontrou o produto
                    If lErro = 78189 Then gError 78200
    
                    objApropriacaoInsumos.lNumIntReqProd = objRequisicao.lNumIntDoc
                    
                End If
                
                'Pega a Quantidade do Produto
                If Len(Trim(GridApropriacao.TextMatrix(iIndice, iGrid_QuantidadeApr_Col))) > 0 Then objApropriacaoInsumos.dQuantidade = CDbl(GridApropriacao.TextMatrix(iIndice, iGrid_QuantidadeApr_Col))
                
                'Adiciona na Apropriacao
                colApropriacaoInsumos.Add objApropriacaoInsumos
                
                Set objApropriacaoInsumos.colRastreamentoMovto = New Collection
                
                'Move o Rastro da Apropriação para Memória
                lErro = Move_RastroApropriacao_Memoria(iIndice, objApropriacaoInsumos.colRastreamentoMovto)
                If lErro <> SUCESSO Then gError 78183
                
                objApropriacaoInsumos.iAutomatico = IIf(GridApropriacao.TextMatrix(iIndice, iGrid_ApropAutomatica_Col) = CStr(MARCADO), APROPINSUMOSPROD_AUTOMATICO, 0)
            
            End If
        
        End If
        
    Next
    
    Move_ApropriacaoInsumos_Memoria2 = SUCESSO
    
    Exit Function
    
Erro_Move_ApropriacaoInsumos_Memoria2:

    Move_ApropriacaoInsumos_Memoria2 = gErr
    
    Select Case gErr
        
        Case 78168, 78170, 78172
            
        Case 78169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOINSUMO_NAO_PREENCHIDO", gErr)
        
        Case 78171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 78183
        
        Case 78200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_SEM_PRODUTO", gErr, objRequisicao.sProduto, objRequisicao.lCodigo, objRequisicao.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165349)
    
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
    
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 78175
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78176

    If lErro = 28030 Then gError 78177
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
            
                objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col)
                
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            objRastreamentoMovto.sLote = GridMovimentos.TextMatrix(iLinha, iGrid_OP_Col)
            objRastreamentoMovto.iFilialOP = giFilialEmpresa
            
            lErro = CF2(Me, "ProdEnt_Move_RastroEstoque_Memoria", iLinha, objRastreamentoMovto)
            If lErro <> SUCESSO Then gError 196431
        
        '################################################################
        'Inserido por Wagner 14/03/2006
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
           
            For Each objRastreamentoMovto In gcolcolRastreamentoSerie.Item(iLinha)
                objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
        '################################################################
            
        End If
        
        If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Or objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then objRastreamentoMovto.dQuantidade = CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))
            objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
            objRastreamentoMovto.sProduto = sProdutoFormatado
            
            '#################################################################
            'Alterado por Wagner 14/03/2006
            If objProduto.iRastro <> PRODUTO_RASTRO_NUM_SERIE Then
                colRastreamentoMovto.Add objRastreamentoMovto
            End If
            '#################################################################
        
        End If
        
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case 78175, 78176
        
        Case 78177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 78337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 196431
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165350)
    
    End Select
    
    Exit Function
    
End Function

Function Move_RastroApropriacao_Memoria(iLinha As Integer, colRastreamentoMovto As Collection) As Long
'Move o Rastro da Apropriacao

Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer, lErro As Long
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim dFator As Double

On Error GoTo Erro_Move_RastroApropriacao_Memoria
    
    lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(iLinha, iGrid_ProdutoApr_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 78179
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78180

    If lErro = 28030 Then gError 78181
    
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM And Len(Trim(GridApropriacao.TextMatrix(iLinha, iGrid_LoteApr_Col))) <> 0 Then
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
            objRastreamentoMovto.sLote = GridApropriacao.TextMatrix(iLinha, iGrid_LoteApr_Col)
        
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
            objRastreamentoMovto.sLote = GridApropriacao.TextMatrix(iLinha, iGrid_LoteApr_Col)
            If Len(Trim(GridApropriacao.TextMatrix(iLinha, iGrid_FilialOPApr_Col))) = 0 Then gError 78336
            objRastreamentoMovto.iFilialOP = Codigo_Extrai(GridApropriacao.TextMatrix(iLinha, iGrid_FilialOPApr_Col))
                
        End If
        
        If Len(Trim(GridApropriacao.TextMatrix(iLinha, iGrid_QuantidadeApr_Col))) = 0 Then gError 78352
        objRastreamentoMovto.dQuantidade = CDbl(GridApropriacao.TextMatrix(iLinha, iGrid_QuantidadeApr_Col))
        objRastreamentoMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_APROP_PRODUCAO_ENTRADA
        objRastreamentoMovto.sProduto = sProdutoFormatado
            
        colRastreamentoMovto.Add objRastreamentoMovto
        
    End If
    
    Move_RastroApropriacao_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroApropriacao_Memoria:

    Move_RastroApropriacao_Memoria = gErr
    
    Select Case gErr
        
        Case 78179, 78180, 78184
        
        Case 78181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 78352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iLinha)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165351)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoLimpar_Click()
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 30320

    Call Limpa_Tela_ProducaoEntrada

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 30320

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165352)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ProducaoEntrada()

Dim lErro As Long
Dim lCodigo As Long
On Error GoTo Erro_Limpa_Tela_ProducaoEntrada

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'prepara o Grid para permitir inserir e excluir Linhas
    objGrid.iProibidoIncluir = 0
    objGrid.iProibidoExcluir = 0
    Call Grid_Inicializa(objGrid)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridApropriacao)

    'Remove os ítens de colItensNumIntDoc
'    Set colItensNumIntDoc = New Collection

    Set gobjMovEst = New ClassMovEstoque

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    Set gcolcolRastreamentoSerie = New Collection 'Inserido por Wagner 14/03/2006
    Set gcolProdutos = New Collection
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjAnotacao = Nothing

    lCodigoAntigo = 0

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ProducaoEntrada:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165353)

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
    
    Set gobjInfoUsu = Nothing
    
    Set gobjMovEst = Nothing
    
    Set objEventoCodigo = Nothing
    Set objEventoOPPadrao = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoOP = Nothing
    Set objEventoCcl = Nothing
    Set objEventoProduto = Nothing
    Set objEventoProdutoAprop = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoRequisicoes = Nothing
    Set objEventoEstoque = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    Set objGridApropriacao = Nothing
    
    Set gcolProdutos = Nothing
    Set gcolcolRastreamentoSerie = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    
    Set gobjAnotacao = Nothing
   
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
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

Public Sub HorasMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()

    If UnidadeMed.ListIndex = -1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30998

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 30998
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165354)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Estorno(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Estorno

    Set objGridInt.objControle = Estorno

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 30999

    Saida_Celula_Estorno = SUCESSO

    Exit Function

Erro_Saida_Celula_Estorno:

    Saida_Celula_Estorno = gErr

    Select Case gErr

        Case 30999
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165355)

    End Select

    Exit Function

End Function

Public Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Public Sub GridMovimentos_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Public Sub GridMovimentos_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Public Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If


End Sub

Public Sub GridMovimentos_LeaveCell()

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
        If lErro <> SUCESSO Then gError 39624
        
        If objGridInt.objGrid Is GridMovimentos Then
        
            Select Case GridMovimentos.Col
    
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 34717
    
                Case iGrid_OP_Col
    
                    lErro = Saida_Celula_OP(objGridInt)
                    If lErro <> SUCESSO Then gError 34718
    
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 34719
                
                Case iGrid_Almoxarifado_Col
    
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 79760
                
                Case iGrid_Benef_Col
                    lErro = Saida_Celula_Benef(objGridInt)
                    If lErro <> SUCESSO Then gError 91272
                
                Case iGrid_Ccl_Col
    
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 34721
    
                Case iGrid_Estorno_Col
    
                    lErro = Saida_Celula_Estorno(objGridInt)
                    If lErro <> SUCESSO Then gError 34722
    
                Case iGrid_UnidadeMed_Col
    
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 34723
                    
                Case iGrid_ContaContabilEst_Col
                    lErro = Saida_Celula_ContaContabilEst(objGridInt)
                    If lErro <> SUCESSO Then gError 49638
                    
                Case iGrid_ContaContabilProducao_Col
                    lErro = Saida_Celula_ContaContabilProducao(objGridInt)
                    If lErro <> SUCESSO Then gError 52039
                     
                Case iGrid_HorasMaquina_Col
                    lErro = Saida_Celula_HorasMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 78140
                     
                Case iGrid_Lote_Col
                    lErro = Saida_Celula_Lote(objGridInt)
                    If lErro <> SUCESSO Then gError 78141
                
                Case iGrid_DataProducao_Col
                    lErro = Saida_Celula_DataProducao(objGridInt)
                    If lErro <> SUCESSO Then gError 78443
                
                Case Else
                    lErro = Saida_Celula_Grid(objGridInt)
                    If lErro <> SUCESSO Then gError 34724
    
            End Select
        
        ElseIf objGridInt.objGrid Is GridApropriacao Then
        
            Select Case GridApropriacao.Col
    
                Case iGrid_ProdutoApr_Col
                    lErro = Saida_Celula_ProdutoApr(objGridInt)
                    If lErro <> SUCESSO Then gError 78155
    
                Case iGrid_ApropItemProd_Col
                    lErro = Saida_Celula_ApropItemProd(objGridInt)
                    If lErro <> SUCESSO Then gError 78156
    
                Case iGrid_ProdutoIns_Col
                    lErro = Saida_Celula_ProdutoIns(objGridInt)
                    If lErro <> SUCESSO Then gError 78156
    
                Case iGrid_QuantidadeApr_Col
                    lErro = Saida_Celula_QuantidadeApr(objGridInt)
                    If lErro <> SUCESSO Then gError 78157
   
                Case iGrid_UnidadeMedApr_Col
                    lErro = Saida_Celula_UnidadeMedApr(objGridInt)
                    If lErro <> SUCESSO Then gError 78158
                    
                Case iGrid_LoteApr_Col
                    lErro = Saida_Celula_LoteApr(objGridInt)
                    If lErro <> SUCESSO Then gError 78159
                     
                Case iGrid_FilialOPApr_Col
                    lErro = Saida_Celula_FilialOpApr(objGridInt)
                    If lErro <> SUCESSO Then gError 78335
                
                Case iGrid_RequisicaoApr_Col
                    lErro = Saida_Celula_RequisicaoApr(objGridInt)
                    If lErro <> SUCESSO Then gError 78160
                     
                Case Else
                    lErro = Saida_Celula_Grid(objGridInt)
                    If lErro <> SUCESSO Then gError 78161
    
            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 34725

        iAlterado = REGISTRO_ALTERADO

    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = gErr
    
    Select Case gErr

        Case 34717, 34718, 34719, 34721, 34722, 34723, 34724, 49638, 52039, 78140, 78141, 78155 To 78161, 78335, 78443, 79760

        Case 34725
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39624
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165356)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Benef(objGridInt As AdmGrid) As Long
'faz a critica da celula de Benef do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Benef

    Set objGridInt.objControle = Benef

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91279
    
    Saida_Celula_Benef = SUCESSO

    Exit Function

Erro_Saida_Celula_Benef:

   Saida_Celula_Benef = gErr

    Select Case gErr

        Case 91279
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165357)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_FilialOpApr(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOpApr As New AdmFiliais
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_FilialOpApr

    Set objGridInt.objControle = FilialOpApr

    If Len(Trim(FilialOpApr.Text)) <> 0 Then
            
        'Verifica se é uma FilialOpApr selecionada
        If FilialOpApr.Text <> FilialOpApr.List(FilialOpApr.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOpApr, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78330
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOpApr.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOpApr com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOpApr)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78331
        
                'Se não encontrou a FilialOpApr
                If lErro = 27378 Then gError 78332
        
                'coloca na tela
                FilialOpApr.Text = iCodigo & SEPARADOR & objFilialOpApr.sNome
                        
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78333
        
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78334

    Saida_Celula_FilialOpApr = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOpApr:

    Saida_Celula_FilialOpApr = gErr

    Select Case gErr

        Case 78330, 78331, 78334
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOpApr.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOpApr.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165358)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO Then gError 49714
    
    Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = gErr

    Select Case gErr

        Case 49639, 49641, 49714
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165359)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 52040
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 52041
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaMascarada
            ContaContabilProducao.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 52042
    
            'conta não cadastrada
            If lErro = 5700 Then gError 52043
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52044
    
    Saida_Celula_ContaContabilProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducao:

    Saida_Celula_ContaContabilProducao = gErr

    Select Case gErr

        Case 52040, 52042, 52044
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52043
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducao.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165360)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Grid(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Grid

    Select Case GridMovimentos.Col

        Case iGrid_Produto_Col

            Set objGridInt.objControle = Produto

        Case iGrid_UnidadeMed_Col

            Set objGridInt.objControle = UnidadeMed

        Case iGrid_Quantidade_Col

            Set objGridInt.objControle = Quantidade

        Case iGrid_Almoxarifado_Col

            Set objGridInt.objControle = Almoxarifado

        Case iGrid_OP_Col

            Set objGridInt.objControle = OPCodigo

        Case iGrid_Ccl_Col

            Set objGridInt.objControle = Ccl

        Case iGrid_Estorno_Col

            Set objGridInt.objControle = Estorno
            
        Case iGrid_ContaContabilEst_Col
        
            Set objGridInt.objControle = ContaContabilEst

    End Select
    
    lErro = CF2(Me, "ProdEnt_Saida_Celula_Grid", objGridInt)
    If lErro <> SUCESSO Then gError 196434

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34726

    Saida_Celula_Grid = SUCESSO

    Exit Function

Erro_Saida_Celula_Grid:

    Saida_Celula_Grid = gErr

    Select Case gErr

        Case 34726, 196434

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165361)

    End Select

    Exit Function

End Function

Public Sub GridMovimentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Public Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Sub GridMovimentos_RowColChange()

    Call Grid_RowColChange(objGrid)

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

Public Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

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

Public Sub HorasMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub HorasMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub HorasMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = HorasMaq
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
            
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    lErro = CF("UMEstoque_Conversao", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 64210

                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
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
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case ALMOXARIFADO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case ESTORNO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                objMnemonicoValor.colValor.Add CInt(GridMovimentos.TextMatrix(iLinha, iGrid_Estorno_Col))
            Next
                            
        Case PRODUTO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case UNIDADE_MED
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case DESCRICAO_ITEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case OP_CODIGO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_OP_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_OP_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
                    
        Case QUANTIDADE1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
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
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEst_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEst_Col)
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case Else
            Error 39650

        End Select

        Calcula_Mnemonico = SUCESSO

        Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39650
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64210
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165362)

    End Select

    Exit Function

End Function

Private Function Preenche_ContaContabilEst(ByVal sProduto As String, Optional ByVal sAlmoxarifado As String = "", Optional ByVal iLinha As Integer = 0) As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabilEst

    If Len(Trim(sAlmoxarifado)) = 0 Then sAlmoxarifado = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
    If iLinha = 0 Then iLinha = GridMovimentos.Row
        
    If Len(Trim(sAlmoxarifado)) > 0 And Len(Trim(sProduto)) > 0 Then
    
        'preenche o objEstoqueProduto
        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 49686
        
        If lErro = 25060 Then gError 52004
        
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 49687
        
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError 49688
        
        If lErro = SUCESSO Then
        
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49689
        
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxuta
            ContaContabilEst.PromptInclude = True
            
            GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEst_Col) = ContaContabilEst.Text
            
        End If
        
    End If
    
    Preenche_ContaContabilEst = SUCESSO
    
    Exit Function
    
Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = gErr
    
    Select Case gErr
        
        Case 49686, 49687, 49688
        
        Case 49689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
             
        Case 52004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165363)
    
    End Select
    
    Exit Function
        
End Function

Private Function Preenche_Almoxarifado(iFilialEmpresa As Integer, sOPCodigo As String, sProduto As String) As Long
'preenche o almoxarifado no grid a partir do item da OP

Dim objItemOP As New ClassItemOP
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long

On Error GoTo Erro_Preenche_Almoxarifado

    objItemOP.iFilialEmpresa = giFilialEmpresa
    objItemOP.sCodigo = sOPCodigo
    objItemOP.sProduto = sProduto

    lErro = CF("ItemOP_Le", objItemOP)
    If lErro <> SUCESSO And lErro <> 34711 Then gError 55333

    If lErro = 34711 Then gError 55332
    
    objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
    
    'le o nome reduzido do almoxarifado associado ao itemop
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 55334
    
    If lErro = 25056 Then gError 55335
    
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

    Call Combo_Seleciona_ItemData(Benef, objItemOP.iBeneficiamento)
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col) = Benef.Text

    Preenche_Almoxarifado = SUCESSO

    Exit Function

Erro_Preenche_Almoxarifado:

    Preenche_Almoxarifado = gErr
    
    Select Case gErr
    
        Case 55332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 55333, 55334

        Case 55335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objItemOP.iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165364)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ENTRADA_MATERIAL_PRODUZIDO_MOVIMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Entrada de Material Produzido"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ProducaoEntrada"
    
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
        ElseIf Me.ActiveControl Is OPCodigoPadrao Then
            Call OPPadraoLabel_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is OPCodigo Then
            Call BotaoOP_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is ContaContabilEst Or Me.ActiveControl Is ContaContabilProducao Then
            Call BotaoPlanoConta_Click
        End If
    
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        Call Trata_CodigoBarras1
    
    End If

End Sub























































Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub





'Raphael 03/10/2000

Private Function Saida_Celula_ProdutoApr(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ProdutoApr

    Set objGridInt.objControle = ProdutoApr

    'Verifica se o produto existe e foi preenchido
    lErro = CF("Produto_Critica", ProdutoApr.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 78142

    'se o produto não estiver cadastrado
    If lErro = 25041 Then gError 78143
            
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_DescricaoApr_Col) = objProduto.sDescricao
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_UnidadeMedApr_Col) = objProduto.sSiglaUMEstoque
            
        If objProduto.iRastro = PRODUTO_RASTRO_OP Then
            GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_FilialOPApr_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
        End If
            
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridApropriacao.Row - GridApropriacao.FixedRows) = objGridApropriacao.iLinhasExistentes Then
            objGridApropriacao.iLinhasExistentes = objGridApropriacao.iLinhasExistentes + 1
        End If
            
    Else
        
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_UnidadeMedApr_Col) = ""
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_DescricaoApr_Col) = ""
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_QuantidadeApr_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78144

    Saida_Celula_ProdutoApr = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoApr:

    Saida_Celula_ProdutoApr = gErr

    Select Case gErr

        Case 78142, 78144
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78143
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoApr.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdutoApr.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165365)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoIns(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ProdutoIns

    Set objGridInt.objControle = ProdutoIns

    'Verifica se o produto existe e foi preenhido
    lErro = CF("Produto_Critica", ProdutoIns.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 78142

    'se o produto não estiver cadastrado
    If lErro = 25041 Then gError 78143
            
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_DescricaoIns_Col) = objProduto.sDescricao
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78144

    Saida_Celula_ProdutoIns = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoIns:

    Saida_Celula_ProdutoIns = gErr

    Select Case gErr

        Case 78142, 78144
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78143
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165366)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeApr(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objItensMovimento As New ClassItemMovEstoque
Dim dQuantidade As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_QuantidadeApr

    Set objGridInt.objControle = QuantidadeApr

    If Len(Trim(QuantidadeApr.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeApr.Text)
        If lErro <> SUCESSO Then gError 78145

        If Len(Trim(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_RequisicaoApr_Col))) > 0 Then
        
            'Formata o produto para ler
            lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoApr_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78205
            
            objItensMovimento.sProduto = sProdutoFormatado
            objItensMovimento.lCodigo = GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_RequisicaoApr_Col)
            objItensMovimento.iFilialEmpresa = giFilialEmpresa
                            
            'Lê o Produto no Movimento de Estoque
            lErro = CF("MovEstoque_Le_Produto", objItensMovimento)
            If lErro <> SUCESSO And lErro <> 78189 Then gError 78197
            
            'Se não encontrar o Movimento de Estoque --> Erro
            If lErro = 78189 Then gError 78198
            
            dQuantidade = CDbl(QuantidadeApr.ClipText)
                
            'Verifica se a quantidade é maior que a quantidade Requisitada
            If dQuantidade > objItensMovimento.dQuantidade Then gError 78199
        
        End If
        
        QuantidadeApr.Text = Formata_Estoque(QuantidadeApr.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78146

    Saida_Celula_QuantidadeApr = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeApr:

    Saida_Celula_QuantidadeApr = gErr

    Select Case gErr

        Case 78145, 78146, 78197
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_SEM_PRODUTO", gErr, objItensMovimento.sProduto, objItensMovimento.lCodigo, objItensMovimento.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_PRODUTO_SEM_QUANTIDADE", gErr, objItensMovimento.sProduto, objItensMovimento.lCodigo, objItensMovimento.dQuantidade)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165367)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMedApr(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMedApr

    Set objGridInt.objControle = UnidadeMedApr

    GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_UnidadeMedApr_Col) = UnidadeMedApr.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78147

    Saida_Celula_UnidadeMedApr = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMedApr:

    Saida_Celula_UnidadeMedApr = gErr

    Select Case gErr

        Case 78147
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165368)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_LoteApr(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Saida_Celula_LoteApr

    Set objGridInt.objControle = LoteApr
    
    If Len(Trim(LoteApr.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoApr_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78460
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78461
            
        If lErro = 28030 Then gError 78462
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                objRastroLote.sCodigo = LoteApr.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78463
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78464
                
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
                If Len(Trim(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_FilialOPApr_Col))) > 0 Then
'??? lixo ?
''                    objOrdemProducao.iFilialEmpresa = Codigo_Extrai(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_FilialOPApr_Col))
''                    objOrdemProducao.sCodigo = LoteApr.Text
''
''                    'Verifica se existe a OP
''                    lErro = CF("OrdemProducao_Le",objOrdemProducao)
''                    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 78467
''
''                    If lErro = 30368 Then gError 78468
''
''                    If lErro = 55316 Then gError 78469
''
                    objRastroLote.sCodigo = LoteApr.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_FilialOPApr_Col))
                    
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78535
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78465
                
                End If
                
            End If
        
        End If
    
    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78534

    Saida_Celula_LoteApr = SUCESSO

    Exit Function

Erro_Saida_Celula_LoteApr:

    Saida_Celula_LoteApr = gErr

    Select Case gErr

        Case 78460, 78461, 78463, 78467, 78534, 78535
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78464, 78465
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 78468
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165369)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RequisicaoApr(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objItensMovimento As New ClassItemMovEstoque
Dim dQuantidade As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_RequisicaoApr

    Set objGridInt.objControle = RequisicaoApr
    
    If Len(Trim(RequisicaoApr.Text)) > 0 Then
                
        'Formata o produto para ler
        lErro = CF("Produto_Formata", GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoApr_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78205
        
        objItensMovimento.sProduto = sProdutoFormatado
        objItensMovimento.lCodigo = CLng(RequisicaoApr.Text)
        objItensMovimento.iFilialEmpresa = giFilialEmpresa
                        
        'Lê o Produto no Movimento de Estoque
        lErro = CF("MovEstoque_Le_Produto", objItensMovimento)
        If lErro <> SUCESSO And lErro <> 78189 Then gError 78194
        
        'Se não encontrar o Movimento de Estoque --> Erro
        If lErro = 78189 Then gError 78195
        
        If Len(Trim(QuantidadeApr.Text)) > 0 Then
            
            dQuantidade = CDbl(GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_QuantidadeApr_Col))
            
            'Verifica se a quantidade é maior que a quantidade Requisitada
'            If dQuantidade > objItensMovimento.dQuantidade Then gError 78196
        
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78149

    Saida_Celula_RequisicaoApr = SUCESSO

    Exit Function

Erro_Saida_Celula_RequisicaoApr:

    Saida_Celula_RequisicaoApr = gErr

    Select Case gErr

        Case 78194, 78205
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_SEM_PRODUTO", gErr, objItensMovimento.sProduto, objItensMovimento.lCodigo, objItensMovimento.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_PRODUTO_SEM_QUANTIDADE", gErr, objItensMovimento.sProduto, objItensMovimento.lCodigo, objItensMovimento.dQuantidade)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165370)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataProducao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataProducao

    Set objGridInt.objControle = DataProducao

    'verifica se a data está preenchida
    If Len(Trim(DataProducao.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataProducao.Text)
        If lErro <> SUCESSO Then gError 78440

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78441

    Saida_Celula_DataProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_DataProducao:

    Saida_Celula_DataProducao = gErr

    Select Case gErr

        Case 78440, 78441
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165371)

    End Select

    Exit Function

End Function

Public Sub ProdutoApr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoApr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub ProdutoApr_KeyPress(KeyAscii As Integer)
        
        Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub ProdutoApr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = ProdutoApr
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ProdutoIns_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoIns_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub ProdutoIns_KeyPress(KeyAscii As Integer)
        
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub ProdutoIns_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = ProdutoIns
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMedApr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMedApr_Click()

    If UnidadeMedApr.ListIndex = -1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMedApr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub UnidadeMedApr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub UnidadeMedApr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = UnidadeMedApr
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantidadeApr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeApr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub QuantidadeApr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub QuantidadeApr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = QuantidadeApr
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub LoteApr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LoteApr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub LoteApr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub LoteApr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = LoteApr
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub RequisicaoApr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RequisicaoApr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub RequisicaoApr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub RequisicaoApr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = RequisicaoApr
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridApropriacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridApropriacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridApropriacao, iAlterado)

    End If

End Sub

Public Sub GridApropriacao_EnterCell()

    Call Grid_Entrada_Celula(objGridApropriacao, iAlterado)

End Sub

Public Sub GridApropriacao_GotFocus()

    Call Grid_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub GridApropriacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridApropriacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridApropriacao, iAlterado)
    End If


End Sub

Public Sub GridApropriacao_LeaveCell()

    Call Saida_Celula(objGridApropriacao)

End Sub

Public Sub GridApropriacao_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer, lNumIntDoc As Long
Dim iLinhaAnterior As Integer

'    If colItensNumIntDoc.Count >= GridApropriacao.Row Then
'        lNumIntDoc = colItensNumIntDoc.Item(GridApropriacao.Row)
'    Else
'        lNumIntDoc = 0
'    End If

    If lNumIntDoc = 0 Then
        
        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then
        
            'Guarda iLinhasExistentes
            iLinhasExistentesAnterior = objGridApropriacao.iLinhasExistentes
    
            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridApropriacao.Row
    
        End If
    
        Call Grid_Trata_Tecla1(KeyCode, objGridApropriacao)
    
    End If

End Sub

Public Sub GridApropriacao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridApropriacao)

End Sub

Public Sub GridApropriacao_Scroll()

    Call Grid_Scroll(objGridApropriacao)

End Sub

Public Sub GridApropriacao_RowColChange()

    Call Grid_RowColChange(objGridApropriacao)

End Sub

Private Function Preenche_GridApropriacao(ByVal objItens As ClassItemMovEstoque, ByVal iItem As Integer) As Long

Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim lErro As Long
Dim objApropriacao As New ClassApropriacaoInsumosProd
Dim objProduto As New ClassProduto
Dim objRequisicao As New ClassItemMovEstoque
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim colRatreamentoMovto As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Preenche_GridApropriacao
    
    iIndice = objGridApropriacao.iLinhasExistentes
    
    'Preenche GridApropriacao
    For Each objApropriacao In objItens.colApropriacaoInsumo

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)
        objProduto.sCodigo = objApropriacao.sProduto
        
        'Le os dados do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78201
        
        'Não encontrou o Produto --> Erro
        If lErro = 28030 Then gError 78202
        
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 34723

        Call StrParaMasked2(ProdutoApr, sProdutoMascarado)
        
'        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
'        If lErro <> SUCESSO Then gError 78203
'
'        ProdutoApr.PromptInclude = False
'        ProdutoApr.Text = sProdutoMascarado
'        ProdutoApr.PromptInclude = True
        
        'Preenche o Produto Insumo
        GridApropriacao.TextMatrix(iIndice, iGrid_ProdutoApr_Col) = sProdutoMascarado
        GridApropriacao.TextMatrix(iIndice, iGrid_DescricaoApr_Col) = objProduto.sDescricao
        GridApropriacao.TextMatrix(iIndice, iGrid_UnidadeMedApr_Col) = objProduto.sSiglaUMEstoque
        
        GridApropriacao.TextMatrix(iIndice, iGrid_ApropItemProd_Col) = CStr(iItem)
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)
        
        lErro = Mascara_RetornaProdutoTela(objItens.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 34723

        Call StrParaMasked2(ProdutoIns, sProdutoMascarado)

'        lErro = Mascara_RetornaProdutoEnxuto(objItens.sProduto, sProdutoMascarado)
'        If lErro <> SUCESSO Then gError 78204
'
'        Produto.PromptInclude = False
'        Produto.Text = sProdutoMascarado
'        Produto.PromptInclude = True
        
        GridApropriacao.TextMatrix(iIndice, iGrid_ProdutoIns_Col) = sProdutoMascarado
        GridApropriacao.TextMatrix(iIndice, iGrid_DescricaoIns_Col) = objItens.sProdutoDesc
        
        GridApropriacao.TextMatrix(iIndice, iGrid_QuantidadeApr_Col) = Formata_Estoque(objApropriacao.dQuantidade)
        
        If objApropriacao.lNumIntReqProd > 0 Then
        
            objRequisicao.lNumIntDoc = objApropriacao.lNumIntReqProd
        
            'Le o Movimeto de Estoque com o número interno vindo do número interno da Requisicao
            lErro = CF("MovEstoque_Le_NumIntDoc", objRequisicao)
            If lErro <> SUCESSO And lErro <> 78193 Then gError 78205
                        
            'não encontrou o movimento (Requisição)
            If lErro = 78193 Then gError 78206
                
            GridApropriacao.TextMatrix(iIndice, iGrid_RequisicaoApr_Col) = objRequisicao.lCodigo
        
        End If
        
        Set colRatreamentoMovto = New Collection
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objApropriacao.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_APROP_PRODUCAO_ENTRADA, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError 78424
        
        'Se existe rastreamento
        If colRatreamentoMovto.Count > 0 Then
                        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto.Item(1)
                                
            GridApropriacao.TextMatrix(iIndice, iGrid_LoteApr_Col) = objRatreamentoMovto.sLote
            
            If objRatreamentoMovto.iFilialOP > 0 Then
            
                objFilialEmpresa.iCodFilial = objRatreamentoMovto.iFilialOP
                
                'Lê o empresa para pegar o nome da Filial
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 78425
                
                GridApropriacao.TextMatrix(iIndice, iGrid_FilialOPApr_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

            End If
            
        End If
        
        If objApropriacao.iAutomatico = 0 Then
            GridApropriacao.TextMatrix(iIndice, iGrid_ApropAutomatica_Col) = CStr(DESMARCADO)
        Else
            GridApropriacao.TextMatrix(iIndice, iGrid_ApropAutomatica_Col) = CStr(MARCADO)
        End If
    
    Next

    objGridApropriacao.iLinhasExistentes = iIndice

    lErro = Grid_Refresh_Checkbox(objGridApropriacao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Preenche_GridApropriacao = SUCESSO

    Exit Function

Erro_Preenche_GridApropriacao:

    Preenche_GridApropriacao = gErr

    Select Case gErr

        Case 78201, 78424, 78425
        
        Case 78202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objApropriacao.sProduto)
        
        Case 78203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objApropriacao.sProduto)
        
        Case 78204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItens.sProduto)

        Case 78205, ERRO_SEM_MENSAGEM
        
        Case 78206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_NAO_CADASTRADA", gErr, objRequisicao.lCodigo, objRequisicao.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165372)

    End Select

    Exit Function

End Function

Function Retorna_Percentagem(dValorTotal As Double, dValorCalculo As Double, sPercentRetornado As String) As Long

Dim dDiferenca As Double
Dim dPercentual As Double
Dim lErro As Long

On Error GoTo Erro_Retorna_Percentagem
    
    lErro = Valor_Positivo_Critica_Double(CStr(dValorTotal))
    If lErro <> SUCESSO Then gError 87564
    
    lErro = Valor_Positivo_Critica_Double(CStr(dValorCalculo))
    If lErro <> SUCESSO Then gError 87565
    
    If dValorTotal = 0 Then gError 87563
    
    dDiferenca = dValorCalculo - dValorTotal
    dPercentual = dDiferenca / dValorTotal

    sPercentRetornado = Format(dPercentual, "Percent")
    
    Retorna_Percentagem = SUCESSO
        
    Exit Function
    
Erro_Retorna_Percentagem:
        
    Retorna_Percentagem = gErr
        
    Select Case gErr
    
        Case 87563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIVISAO_POR_0", gErr)
        
        Case 87564, 87565
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165373)
        
        End Select
        
        Exit Function
        
End Function

Function Traz_DadosOP_Tela(objOrdemProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long, iIndice As Integer, iLinha As Integer
Dim objItemOP As ClassItemOP
Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sCclMascarado As String, bOPJaIncluida As Boolean
Dim objItemMovEst As ClassItemMovEstoque
Dim colRastreamentoSerie As Collection 'Inserido por Wagner 14/03/2006
Dim objProdutoAux As New ClassProduto 'Inserido por Wagner 15/03/2006
Dim bAchouProduto As Boolean 'Inserido por Wagner 15/03/2006
Dim sContaEnxuta As String

On Error GoTo Erro_Traz_DadosOP_Tela

    'verifica se já há algum item da OP carregado
    bOPJaIncluida = False
    For iIndice = 1 To objGrid.iLinhasExistentes
    
        If GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col) = objOrdemProducao.sCodigo Then
            
            bOPJaIncluida = True
            Exit For
            
        End If
    
    Next
    
    'se a OP ainda nao está no grid
    If bOPJaIncluida = False Then
    
        lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO Then gError 86550
        
        For Each objItemOP In objOrdemProducao.colItens
        
            If objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA And (objItemOP.dQuantidade - objItemOP.dQuantidadeProd) > QTDE_ESTOQUE_DELTA2 Then
                
                objProduto.sCodigo = objItemOP.sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 86551
                If lErro <> SUCESSO Then gError 86552
                
                sProdutoMascarado = String(STRING_PRODUTO, 0)
        
                lErro = Mascara_RetornaProdutoEnxuto(objItemOP.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 86553
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoMascarado
                Produto.PromptInclude = True
                
                objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
                
                'le o nome reduzido do almoxarifado associado ao itemop
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 86554
                
                If lErro = 25056 Then gError 86555
                
                If Trim(objItemOP.sCcl) <> "" Then
                    sCclMascarado = String(STRING_CCL, 0)
            
                    lErro = Mascara_MascararCcl(objItemOP.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then gError 86556
                Else
                    sCclMascarado = ""
                End If
                
                iLinha = objGrid.iLinhasExistentes + 1
                
                Produto.PromptInclude = False
                Produto.Text = sProdutoMascarado
                Produto.PromptInclude = True
                
                GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
                GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
                
                GridMovimentos.TextMatrix(iLinha, iGrid_Descricao_Col) = objProduto.sDescricao
                GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = objItemOP.sSiglaUM
                GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objItemOP.dQuantidade - objItemOP.dQuantidadeProd)
                GridMovimentos.TextMatrix(iLinha, iGrid_OP_Col) = objOrdemProducao.sCodigo
                GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col) = sCclMascarado
                
                GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col) = "0"
        
                GridMovimentos.TextMatrix(iIndice, iGrid_DataProducao_Col) = ""
                
                GridMovimentos.TextMatrix(iIndice, iGrid_Estorno_Col) = CStr(DESMARCADO)
        
                'GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col) = objItemOP.iBeneficiamento
                
                Call Combo_Seleciona_ItemData(Benef, objItemOP.iBeneficiamento)
                GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col) = Benef.Text
        
                objGrid.iLinhasExistentes = iLinha
                
                Call gobjTelaProjetoInfo.Adiciona_Linha(iLinha)
            
                Set objItemMovEst = New ClassItemMovEstoque
            
                gobjMovEst.colItens.Add1 objItemMovEst
            
                lErro = CF("Customiza_ProducaoEnt_SaidaCel", objProduto)
                If lErro <> SUCESSO Then gError 117698
            
                If Len(Trim(objProduto.sGrade)) = 0 Then
        
                    objItemMovEst.iPossuiGrade = DESMARCADO
                    
                Else
                
                    objItemMovEst.iPossuiGrade = MARCADO
                    GridMovimentos.TextMatrix(iLinha, 0) = "# " & GridMovimentos.TextMatrix(iLinha, 0)
                
                End If
                
                objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
                objItemMovEst.sProduto = objProduto.sCodigo
                objItemMovEst.sProdutoDesc = objProduto.sDescricao
                objItemMovEst.sSiglaUM = objItemOP.sSiglaUM
                
                '##################################################################
                'Preenche as contas de produção e estoque
                lErro = Obtem_Conta_Producao(objProduto)
                If lErro <> SUCESSO Then gError 190834
                
                If Trim(objProduto.sContaContabilProducao) <> "" Then
                
                    lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
                    If lErro <> SUCESSO Then gError 190836
                
                    'preenche  a ContaContabilProducao
                    ContaContabilEst.PromptInclude = False
                    ContaContabilEst.Text = sContaEnxuta
                    ContaContabilEst.PromptInclude = True
                
                    'preenche Conta De Producao
                    GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilProducao_Col) = ContaContabilEst.Text
                
                End If
            
                'Preenche a conta contabil de estoque depois que o produto e o Almoxarifado já estão preenchidos
                lErro = Preenche_ContaContabilEst(Produto.Text, objAlmoxarifado.sNomeReduzido, iLinha)
                If lErro <> SUCESSO Then gError 190835
                '##################################################################
    
                '###################################################
                'Inserido por Wagner 14/03/2006
                bAchouProduto = False
                For Each objProdutoAux In gcolProdutos
                    If objProdutoAux.sCodigo = objProduto.sCodigo Then
                        bAchouProduto = True
                        Exit For
                    End If
                Next
                If Not bAchouProduto Then
                    gcolProdutos.Add objProduto
                Else
                    Set objProduto = objProdutoAux
                End If
                
                Set colRastreamentoSerie = New Collection
                
                gcolcolRastreamentoSerie.Add colRastreamentoSerie
    
                If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
                    GridMovimentos.TextMatrix(iLinha, iGrid_Lote_Col) = objProduto.sSerieProx
    
                    'Carrega as séries na coleção global
                    lErro = Carrega_Series(gcolcolRastreamentoSerie.Item(iLinha), objItemOP.dQuantidade - objItemOP.dQuantidadeProd, objProduto.sSerieProx, 0, "", iLinha, True)
                    If lErro <> SUCESSO Then gError 141911
                    
                End If
                '###################################################
    
            End If
        Next

        lErro = Grid_Refresh_Checkbox(objGrid)
        If lErro <> SUCESSO Then gError 30925
    
    End If
    
    lErro = Calcula_Rateio_HorasMaquina()
    If lErro <> SUCESSO Then gError 86550
    
    Traz_DadosOP_Tela = SUCESSO

    Exit Function
    
Erro_Traz_DadosOP_Tela:

    Traz_DadosOP_Tela = gErr
    
    Select Case gErr
    
        Case 30925, 86550, 86551, 86553, 86554, 86556, 117698, 141911, 190834, 190835, 190836
        
        Case 86552
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 86555
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165374)
            
    End Select
    
    Exit Function

End Function

Public Sub BotaoImprimirRotulos_Click()

Dim objRelatorio As New AdmRelatorio
Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsgRes As VbMsgBoxResult
Dim sCodigoOP As String, objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_BotaoImprimirRotulos_Click

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridMovimentos
    If objGrid.iLinhasExistentes = 0 Then gError 30985

    'Para cada MovEstoque
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 30986

        'Verifica se OPCodigo foi preenchido
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col))) = 0 Then gError 30988

        sCodigoOP = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
        
        Set objItemMovEst = New ClassItemMovEstoque

        'Coloca o Produto no formato do BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71413

        objItemMovEst.sOPCodigo = sCodigoOP
        objItemMovEst.sProduto = sProdutoFormatado
    
        lErro = CF("ReqProducao_TestaExistencia", objItemMovEst)
        If lErro <> SUCESSO And lErro <> 71417 Then gError 71413

        If lErro = 71417 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NAO_EXISTE_REQPRODUCAO_OP", sCodigoOP)
            If vbMsgRes = vbNo Then gError 71418
    
        End If
        
    Next
   
    'Move os dados para o objMovimentoEstoque
    lErro = Move_Tela_Memoria(objMovEstoque, True)
    If lErro <> SUCESSO Then gError 30989
    
    Call objRelatorio.Rel_Menu_Executar("Rótulos de Produção", objMovEstoque)

    Exit Sub

Erro_BotaoImprimirRotulos_Click:
    
    Select Case gErr
       
        Case 30979
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 30984
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 30985
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_INFORMADO", gErr)

        Case 30986
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 30988
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_PREENCHIDO1", gErr, iIndice)

        Case 30989, 30990, 30395, 30981, 78351
        
        Case 78622, 71413, 71418, 92028
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165375)

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
Dim lBenef As Long

On Error GoTo Erro_BotaoSerie_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 141914
    
    'Se o produto não foi preenchido, erro
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 141915
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Lote_Col))) = 0 Then gError 177303
    If StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) = 0 Then gError 177304
        
    'Formata o produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141916
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141917
       
    objItemMovEstoque.dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    objItemMovEstoque.iItemNF = GridMovimentos.Row
    objItemMovEstoque.sAlmoxarifadoNomeRed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col)
    objItemMovEstoque.sProduto = sProdutoFormatado
    objItemMovEstoque.sSiglaUM = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)
    
    Call Combo_Obtem_ItemData(Benef, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col), lBenef)
    
    If lBenef = BENEF_COMBO_BENEF Then
        objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_PRODUCAO_BENEF3
    ElseIf lBenef = BENEF_COMBO_OUTROS Then
        objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_PRODUCAO_OUTROS
    Else
        objItemMovEstoque.iTipoMov = MOV_EST_ESTORNO_PRODUCAO
    End If
            
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela_Modal("RastreamentoSerie", gcolcolRastreamentoSerie.Item(GridMovimentos.Row), objItemMovEstoque, Me.Name)
                    
    lErro = Acerta_Quantidade_Rastreada(GridMovimentos.Row)
    If lErro <> SUCESSO Then gError 141918
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
    
        Case 141914
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 141915 To 141918
        
        Case 177303
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, GridMovimentos.Row)
        
        Case 177304
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA", gErr, GridMovimentos.Row)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141919)
    
    End Select
    
    Exit Sub

End Sub

Public Function Carrega_Series(colRastreamentoMovto As Collection, ByVal dQuantidade As Double, ByVal sLoteIni As String, ByVal dQuantidadeAnterior As Double, ByVal sLoteIniAnterior As String, ByVal iLinha As Integer, Optional bAtualizaProximoLote As Boolean = False)
'Gera o rastreamento por número de série

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim iIndiceAux As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoSerie As ClassRastreamentoLote
Dim objRastreamentoSerieIni As ClassRastreamentoLote
Dim vbResult As VbMsgBoxResult
Dim colRastreamentoMovtoAux As New Collection
Dim sLote As String
Dim objProdutoAux As New ClassProduto
Dim dQuantidadeEst As Double
Dim dFator As Double

On Error GoTo Erro_Carrega_Series

    'Formata o produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141921
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141922
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 141923

    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
        lErro = CF("UM_Conversao", objProduto.iClasseUM, GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 177312
        
        dQuantidadeEst = dQuantidade * dFator
    
        lErro = Valor_Inteiro_Critica(CStr(dQuantidadeEst))
        If lErro <> SUCESSO Then gError 177308
    
        If dQuantidadeAnterior <> 0 And Len(Trim(sLoteIniAnterior)) <> 0 Then
            If Abs(dQuantidade - dQuantidadeAnterior) > QTDE_ESTOQUE_DELTA Or sLoteIni <> sLoteIniAnterior Then
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_MODIFICACAO_SERIES")
                If vbResult = vbNo Then gError 141920
            Else
                vbResult = vbNo
            End If
        Else
            vbResult = vbYes
        End If
                    
        If vbResult = vbYes Then
            
            If Len(Trim(sLoteIni)) <> 0 Then
            
                If Not IsNumeric(right(sLoteIni, objProduto.iSerieParteNum)) Then gError 141924
                
                Set objRastreamentoSerieIni = New ClassRastreamentoLote
                
                objRastreamentoSerieIni.sProduto = objProduto.sCodigo
                objRastreamentoSerieIni.sCodigo = sLoteIni
                
                lErro = CF("RastreamentoLote_Le", objRastreamentoSerieIni)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 141925
                
                If lErro = SUCESSO Then gError 177314
                        
                iIndiceAux = 0
                iIndice = 1
                        
                Do While iIndice <= CInt(dQuantidadeEst)
                
                    iIndiceAux = iIndiceAux + 1
                
                    Set objRastreamentoMovto = New ClassRastreamentoMovto
                    Set objRastreamentoSerie = New ClassRastreamentoLote
                
                    objRastreamentoMovto.sProduto = objProduto.sCodigo
                    objRastreamentoMovto.sLote = left(sLoteIni, Len(sLoteIni) - objProduto.iSerieParteNum) & FormataCpoNum(iIndiceAux - 1 + StrParaLong(right(sLoteIni, objProduto.iSerieParteNum)), objProduto.iSerieParteNum)
                
                    objRastreamentoSerie.sProduto = objProduto.sCodigo
                    objRastreamentoSerie.iFilialOP = objRastreamentoMovto.iFilialOP
                    objRastreamentoSerie.sCodigo = objRastreamentoMovto.sLote
                    
                    sLote = objRastreamentoMovto.sLote
                    
                    lErro = CF("RastreamentoLote_Le", objRastreamentoSerie)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 141927
                    
                    'Não tem esse número de série
                    If lErro <> SUCESSO Then
                    
                        iIndice = iIndice + 1
                            
                        objRastreamentoMovto.dQuantidadeEst = 1
                        objRastreamentoMovto.dQuantidadeSerie = dQuantidadeEst
                        objRastreamentoMovto.dQuantidade = 1 / dFator
                        objRastreamentoMovto.lNumIntDocLote = objRastreamentoSerie.lNumIntDoc
                        objRastreamentoMovto.lNumIntDocLoteSerieIni = objRastreamentoSerieIni.lNumIntDoc
                        objRastreamentoMovto.sSiglaUM = objProduto.sSiglaUMEstoque
                    
                        colRastreamentoMovtoAux.Add objRastreamentoMovto
                        
                    End If
                
                Loop
                
            End If
            
            'Remove os dados anteriores
            For iIndice = colRastreamentoMovto.Count To 1 Step -1
                colRastreamentoMovto.Remove iIndice
            Next
            
            'Coloca os novos dados
            For Each objRastreamentoMovto In colRastreamentoMovtoAux
                colRastreamentoMovto.Add objRastreamentoMovto
            Next
            
            'Se foi gerado algum lote tem que atualizar o próximo Lote
            If bAtualizaProximoLote Then
            
                If Len(Trim(sLote)) > 0 Then
                
                    For Each objProdutoAux In gcolProdutos
                    
                        If objProdutoAux.sCodigo = objProduto.sCodigo Then
                            
                            objProdutoAux.sSerieProx = sLote
                            Exit For
                        End If
                    
                    Next
                    
                End If
            End If
            
        End If
        
    End If

    Carrega_Series = SUCESSO
    
    Exit Function

Erro_Carrega_Series:

    Carrega_Series = gErr

    Select Case gErr
    
        Case 141921, 141922, 141925, 141927, 141920, 141929, 177308
        
        Case 141923
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 141924
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, right(sLoteIni, objProduto.iSerieParteNum))
        
        Case 141926
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo, objRastreamentoSerieIni.iFilialOP)
        
        Case 141928
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastreamentoSerie.sProduto, objRastreamentoSerie.sCodigo, objRastreamentoSerie.iFilialOP)
        
        Case 177314
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_CADASTRADO", gErr, objRastreamentoSerieIni.sProduto, objRastreamentoSerieIni.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141930)

    End Select

    Exit Function

End Function

Public Function Acerta_Quantidade_Rastreada(ByVal iLinha As Integer)
'Acerta a quantidade do lote no grid com base na coleção global

Dim lErro As Long
Dim dQuantidade As Double
Dim objRastreamentoSerie As ClassRastreamentoMovto

On Error GoTo Erro_Acerta_Quantidade_Rastreada

    For Each objRastreamentoSerie In gcolcolRastreamentoSerie.Item(iLinha)
    
        dQuantidade = dQuantidade + objRastreamentoSerie.dQuantidade
    
    Next
        
    GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

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
        objItensPRJCR.sDescricao = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)
        objItensPRJCR.sObservacao = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridMovimentos.TextMatrix(iIndice, iGrid_Descricao_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PRODENTRADA
        
        colItensPRJ.Add objItensPRJCR

    Next
    
    If colItensPRJ.Count = 0 Then
        Set objItensPRJCR = New ClassItensPRJCR
        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PRODENTRADA
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

Function Obtem_Conta_Producao(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objCTBConfig As New ClassCTBConfig

On Error GoTo Erro_Obtem_Conta_Producao

    If Trim(objProduto.sContaContabilProducao) = "" Then
            
        objTipoDeProduto.iTipo = objProduto.iTipo
        
        'pesquisa a conta junto ao tipo do produto
        lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 190830
        
        If lErro = 22531 Then gError 190831
        
        objProduto.sContaContabilProducao = objTipoDeProduto.sContaProducao
                                
        'se não encontrar a ContaContabilProducao em Produto e TipoProduto procurar em CTBConfig à nivel de filialEmpresa
        If Trim(objProduto.sContaContabilProducao) = "" Then
                            
            objCTBConfig.sCodigo = CONTA_PRODUCAO_FILIAL
            objCTBConfig.iFilialEmpresa = giFilialEmpresa
                    
            lErro = CF("CTBConfig_Le", objCTBConfig)
            If lErro <> SUCESSO And lErro <> 9755 Then gError 190832
            
            If lErro = SUCESSO Then objProduto.sContaContabilProducao = objCTBConfig.sConteudo
            
        End If
        
    End If
                
    Obtem_Conta_Producao = SUCESSO

    Exit Function

Erro_Obtem_Conta_Producao:

    Obtem_Conta_Producao = gErr

    Select Case gErr
    
        Case 190830, 190832
        
        Case 190831
             Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190833)

    End Select

    Exit Function
    
End Function
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

Public Function ProdEnt_Move_Tela_Memoria(ByVal objCT As Object, ByVal objMovEstoque As ClassMovEstoque) As Long
    ProdEnt_Move_Tela_Memoria = SUCESSO
End Function

Public Function ProdEnt_Preenche_Tela(ByVal objCT As Object, ByVal objMovEstoque As ClassMovEstoque) As Long
    ProdEnt_Preenche_Tela = SUCESSO
End Function

Public Function ProdEnt_Preenche_GridMovimentos(ByVal objCT As Object, ByVal objItemMovEstoque As ClassItemMovEstoque, ByVal iIndice As Integer) As Long
    ProdEnt_Preenche_GridMovimentos = SUCESSO
End Function

Public Function ProdEnt_Move_RastroEstoque_Memoria(ByVal objCT As Object, ByVal iLinha As Integer, ByVal objRastreamentoMovto As ClassRastreamentoMovto) As Long
    ProdEnt_Move_RastroEstoque_Memoria = SUCESSO
End Function

Public Function ProdEnt_Rotina_Grid_Enable(ByVal objCT As Object, ByVal iLinha As Integer, ByVal objControl As Object, ByVal iLocalChamada As Integer, ByVal iProdutoPreenchido As Integer, ByVal sProdutoFormatado As String, ByVal lNumIntDoc As Long) As Long
    ProdEnt_Rotina_Grid_Enable = SUCESSO
End Function

Public Function ProdEnt_Saida_Celula_Grid(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    ProdEnt_Saida_Celula_Grid = SUCESSO
End Function

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

Public Function Calcula_Rateio_HorasMaquina() As Long
'Preenche ou atualiza as colunas Horas Maquina do GridMovimentos de acordo com o total da quantidade
'e o valor informado em Horas Maquina

Dim lErro As Long
Dim dSomaTotal As Double
Dim dLinhaQuant As Double
Dim dHorasMaqAtual As Double, iIndice As Integer, dSoma As Double
Dim iUltimoPreenchido As Integer

On Error GoTo Erro_Calcula_Rateio_HorasMaquina
                            
    '---> For para recolher a soma de todas as quantidades do Grid
                            
    'Percorre todas as linhas do grid e armazena total da quantidade na variável dSomaTotal
    For iIndice = 1 To objGrid.iLinhasExistentes
        
        'Armazena a quantidade da linha selecionada
        dSoma = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        'Soma a quantidade total
        dSomaTotal = dSoma + dSomaTotal
                      
    Next
        
    '---> For para preencher ou atualizar o valor de Horas Maquina do Grid
    
    dSoma = 0
    
    If dSomaTotal <> 0 Then
    
        'Preenche cada linha do Grid
        For iIndice = 1 To objGrid.iLinhasExistentes
        
            'Armazena a quantidade da linha atual
            dLinhaQuant = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
                
            'Cálcula valor para coluna Horas máquina
            dHorasMaqAtual = (StrParaDbl(HorasMaquina.Text) * (dLinhaQuant / dSomaTotal))
            
            'Preenche coluna Horas Maquina com valor calculado
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
                
                iUltimoPreenchido = iIndice
                
                'Preenche coluna Horas Maquina com valor calculado
                GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col) = Round(dHorasMaqAtual, 0)
                dSoma = dSoma + StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col))
                
            Else
                GridMovimentos.TextMatrix(iIndice, iGrid_HorasMaquina_Col) = ""
                               
            End If
            
            'Tratamento para casos de arredondamento de valor
            If iIndice = objGrid.iLinhasExistentes Then
                If iUltimoPreenchido <> 0 Then
                    If dSoma <> StrParaDbl(HorasMaquina.Text) Then GridMovimentos.TextMatrix(iUltimoPreenchido, iGrid_HorasMaquina_Col) = Format(dHorasMaqAtual, "Standard") + (StrParaDbl(HorasMaquina.Text - dSoma))
                End If
            End If
            
        Next

    End If

    Calcula_Rateio_HorasMaquina = SUCESSO
        
    Exit Function
    
Erro_Calcula_Rateio_HorasMaquina:
    
    Calcula_Rateio_HorasMaquina = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    Exit Function

End Function
'02/10/01 Marcelo
Function ItensOrdemProducao_Verifica() As Long
'Verifica se os itens de Ordem de Produção que estao entrando estao baixados, se não pergunta se deseja baixa-los

Dim lErro As Long
Dim iIndice As Integer, sProdutoFormatado As String
Dim objItemOP As New ClassItemOP, iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ItensOrdemProducao_Verifica
              
    For iIndice = 1 To objGrid.iLinhasExistentes
        
        objItemOP.iFilialEmpresa = giFilialEmpresa
        objItemOP.sCodigo = GridMovimentos.TextMatrix(iIndice, iGrid_OP_Col)
        objItemOP.iItem = GridMovimentos.TextMatrix(iIndice, iGrid_Sequencial_Col)
        
        'Formata o Produto para BD
        sProdutoFormatado = ""
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        
        objItemOP.sProduto = sProdutoFormatado
        
        'verifica se o itemop está na tabela ItensOrdemProducao (nao baixada)
        lErro = CF("ItemOP_Le_Item", objItemOP)
        If lErro <> SUCESSO And lErro <> 33215 Then gError 93650
        
        'Inicio Alteracao Daniel em 14/11/2001
        If lErro = SUCESSO And objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
        'Fim Alteracao Daniel em 14/11/2001
        
            'Pergunta ao usuário se quer dar baixa do restante de itens da OP
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BAIXAR_ITEM_OP", objItemOP.sProduto)

            If vbMsgRes = vbYes Then

                'Baixa o item da OP
                lErro = CF("ItemOP_Baixa", objItemOP)
                If lErro <> SUCESSO Then gError 93648
    
            End If
        End If
        
        
    Next
          
    ItensOrdemProducao_Verifica = SUCESSO
     
    Exit Function
    
Erro_ItensOrdemProducao_Verifica:

    ItensOrdemProducao_Verifica = gErr
     
    Select Case gErr
      
        Case 63650
        
        Case 93648
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Sub HorasMaquina_GotFocus()

    Call MaskEdBox_TrataGotFocus(HorasMaquina, iAlterado)

End Sub

Public Sub HorasMaquina_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HorasMaquina_Validate

    If Len(Trim(HorasMaquina.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(HorasMaquina.Text)
        If lErro <> SUCESSO Then gError 87890
        
        'Preenche as colunas HorasMaquina do grid
        lErro = Calcula_Rateio_HorasMaquina()
        If lErro <> SUCESSO Then gError 87894
                
    End If
    
    Exit Sub
    
Erro_HorasMaquina_Validate:

    Cancel = True

    Select Case gErr
            
        Case 87890, 87894
            'Tratado na rotina chamada
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
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

    If objGrid.iLinhasExistentes + 1 = GridMovimentos.Row Then
    
        'Verifica se o Produto está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then
            
            If ActiveControl Is Produto Then
                    
                    Set objGrid.objControle = Produto
            
                    lErro = Grid_Abandona_Celula(objGrid)
                    If lErro <> SUCESSO Then gError 210869
                    
            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 210870
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 210871
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 210872
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 210873
        
                Me.Show
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
                
                gError 210874
                
'                If Not ActiveControl Is Produto Then
'                    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
'
'                    'Preenche a Linha do Grid
'                    lErro = ProdutoLinha_Preenche(objProduto, objItemMovEst)
'                    If lErro <> SUCESSO Then gError 210825
'
'
'                End If
'
'                GridMovimentos.SetFocus
'                GridMovimentos.FocusRect = flexFocusHeavy
    
            Else
            
                gError 210875
    
    
            End If
    
        End If
    
    End If

    Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_Trata_CodigoBarras1:

    Trata_CodigoBarras1 = gErr

'    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 210869, 210870, 210871, 210874, 210875

        Case 210872
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 210873
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210876)

    End Select

    Exit Function

End Function


Public Sub ApropItemProd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ApropItemProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApropriacao)

End Sub

Public Sub ApropItemProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApropriacao)

End Sub

Public Sub ApropItemProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApropriacao.objControle = ApropItemProd
    lErro = Grid_Campo_Libera_Foco(objGridApropriacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ApropItemProd(objGridInt As AdmGrid) As Long

Dim lErro As Long, iItem As Integer

On Error GoTo Erro_Saida_Celula_ApropItemProd

    Set objGridInt.objControle = ApropItemProd

    If Len(Trim(ApropItemProd.ClipText)) <> 0 Then
    
        iItem = StrParaInt(ApropItemProd.ClipText)
        
        If iItem <= 0 Or iItem > objGrid.iLinhasExistentes Then gError 201475
        
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_ProdutoIns_Col) = GridMovimentos.TextMatrix(iItem, iGrid_Produto_Col)
        GridApropriacao.TextMatrix(GridApropriacao.Row, iGrid_DescricaoIns_Col) = GridMovimentos.TextMatrix(iItem, iGrid_Descricao_Col)
        
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78144

    Saida_Celula_ApropItemProd = SUCESSO

    Exit Function

Erro_Saida_Celula_ApropItemProd:

    Saida_Celula_ApropItemProd = gErr

    Select Case gErr

        Case 78142, 78144
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 201475
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_FORA_GRID_MOVTOS", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165366)

    End Select

    Exit Function

End Function

Private Function GridAprop_Exclui_ItemMov(ByVal iItem As Integer) As Long

Dim iLinha As Integer, iLinhasExcluidas As Integer
Dim lErro As Long
Dim bExcluiu As Boolean

On Error GoTo Erro_GridAprop_Exclui_ItemMov

    bExcluiu = False
    
    For iLinha = objGridApropriacao.iLinhasExistentes To 1 Step -1
    
        If StrParaInt(GridApropriacao.TextMatrix(iLinha, iGrid_ApropItemProd_Col)) = iItem Then
        
            'exclui a linha em questão
            Call Grid_Exclui_Linha(objGridApropriacao, iLinha)
            bExcluiu = True
        
        End If
    
    Next
    
    If bExcluiu Then
        For iLinha = 1 To objGridApropriacao.iLinhasExistentes
            If Len(Trim(GridApropriacao.TextMatrix(iLinha, iGrid_ApropItemProd_Col))) > 0 Then
                If CInt(GridApropriacao.TextMatrix(iLinha, iGrid_ApropItemProd_Col)) > iItem Then
                    GridApropriacao.TextMatrix(iLinha, iGrid_ApropItemProd_Col) = CStr(CInt(GridApropriacao.TextMatrix(iLinha, iGrid_ApropItemProd_Col)) - 1)
                End If
            End If
        Next
    End If
    
    lErro = Grid_Refresh_Checkbox(objGridApropriacao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GridAprop_Exclui_ItemMov = SUCESSO
    
    Exit Function
    
Erro_GridAprop_Exclui_ItemMov:

    GridAprop_Exclui_ItemMov = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201476)

    End Select
    
    Exit Function

End Function
