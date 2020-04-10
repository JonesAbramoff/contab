VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

Public gobjInfoUsu As Object

Public gobjInfoAdic As ClassInfoAdic
Public gobjTelaProjetoInfo As ClassTelaPRJInfo

'Property Variables:
Dim m_Caption As String
Event Unload()

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoObservacao As AdmEvento
Attribute objEventoObservacao.VB_VarHelpID = -1
Private WithEvents objEventoBotaoPedidosAvulsos As AdmEvento
Attribute objEventoBotaoPedidosAvulsos.VB_VarHelpID = -1
Private WithEvents objEventoProdutos As AdmEvento
Attribute objEventoProdutos.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornec As AdmEvento
Attribute objEventoFornec.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxarifados As AdmEvento
Attribute objEventoAlmoxarifados.VB_VarHelpID = -1
Private WithEvents objEventoBotaoCcl As AdmEvento
Attribute objEventoBotaoCcl.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoNotaPC As AdmEvento
Attribute objEventoNotaPC.VB_VarHelpID = -1
Private WithEvents objEventoEmbalagens As AdmEvento 'leo
Attribute objEventoEmbalagens.VB_VarHelpID = -1

Dim iPedidoBrowse As Integer

Public iAlterado As Integer
Dim iFornecAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameTipoDestinoAtual As Integer
Dim iGridBloqueiosLinhasBloqueadas As Integer

Dim iFilialEmpresaEntregaAnt As Integer

Public objGridItens As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_PercentDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
'Alteracao Daniel
Public iGrid_PrecoUnitarioMoedaReal_Col As Integer
Public iGrid_TotalMoedaReal_Col As Integer
'Fim da Alteracao Daniel
Public iGrid_DataLimite_Col As Integer
Public iGrid_AliquotaIPI_Col As Integer
Public iGrid_ValorIPIItem_Col As Integer
Public iGrid_AliquotaICMS_Col As Integer
Public iGrid_PercentMaisReceb_Col As Integer
Public iGrid_RecebForaFaixa_Col As Integer
Public iGrid_Observacao_Col As Integer
Public iGrid_DescCompleta_Col As Integer 'leo
Public iGrid_DeliveryDate_Col As Integer
Public iGrid_TempoTransito_Col As Integer

Dim objGridDistribuicao As AdmGrid
Dim iGrid_Prod_Col As Integer
Public iGrid_DescProd_Col As Integer
Dim iGrid_CentroCusto_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_UnidMed_Col As Integer
Dim iGrid_Quant_Col As Integer
Dim iGrid_ContaContabil_Col As Integer

Dim gsCodigoAnterior As String

Dim objGridBloqueio As AdmGrid
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_DataBloqueio_Col As Integer
Dim iGrid_CodUsuario_Col As Integer
Dim iGrid_ResponsavelBL_Col As Integer
Dim iGrid_DataLiberacao_Col As Integer
Dim iGrid_ResponsavelLib_Col As Integer

'Grid de Notas '??? POR LEO
Dim objGridNotas As AdmGrid
Dim iGrid_NotaPC_Col As Integer

Dim gcolItemPedido As Collection
Dim gobjPC As ClassPedidoCompras

Dim iFrameAtual As Integer
Dim colAliqExternas As Collection
'Alteracao Daniel
Public bExibirColReal As Boolean

'--- inicio dos properties get dos controles da tela
Public Property Get DeliveryDate() As Object
     Set DeliveryDate = objUserControl.Controls("DeliveryDate")
End Property

Public Property Get TempoTransito() As Object
     Set TempoTransito = objUserControl.Controls("TempoTransito")
End Property

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoPedidosAvulsos() As Object
     Set BotaoPedidosAvulsos = objUserControl.Controls("BotaoPedidosAvulsos")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get ObsEmbalagem() As Object
     Set ObsEmbalagem = objUserControl.Controls("ObsEmbalagem")
End Property

Public Property Get BotaoTrazCotacao() As Object
     Set BotaoTrazCotacao = objUserControl.Controls("BotaoTrazCotacao")
End Property

Public Property Get Moeda() As Object
     Set Moeda = objUserControl.Controls("Moeda")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get CondPagto() As Object
     Set CondPagto = objUserControl.Controls("CondPagto")
End Property

Public Property Get ObservacaoPC() As Object
     Set ObservacaoPC = objUserControl.Controls("ObservacaoPC")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get Contato() As Object
     Set Contato = objUserControl.Controls("Contato")
End Property

Public Property Get Embalagem() As Object
     Set Embalagem = objUserControl.Controls("Embalagem")
End Property

Public Property Get Taxa() As Object
     Set Taxa = objUserControl.Controls("Taxa")
End Property

Public Property Get LabelEmbalagem() As Object
     Set LabelEmbalagem = objUserControl.Controls("LabelEmbalagem")
End Property

Public Property Get Comprador() As Object
     Set Comprador = objUserControl.Controls("Comprador")
End Property

Public Property Get ObsLabel() As Object
     Set ObsLabel = objUserControl.Controls("ObsLabel")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
End Property

Public Property Get CodigoLabel() As Object
     Set CodigoLabel = objUserControl.Controls("CodigoLabel")
End Property

Public Property Get Frame8() As Object
     Set Frame8 = objUserControl.Controls("Frame8")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get UpDownDataEnvio() As Object
     Set UpDownDataEnvio = objUserControl.Controls("UpDownDataEnvio")
End Property

Public Property Get DataEnvio() As Object
     Set DataEnvio = objUserControl.Controls("DataEnvio")
End Property

Public Property Get DataAlteracao() As Object
     Set DataAlteracao = objUserControl.Controls("DataAlteracao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get FrameTipo() As Object
     Set FrameTipo = objUserControl.Controls("FrameTipo")
End Property

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get FilialFornec() As Object
     Set FilialFornec = objUserControl.Controls("FilialFornec")
End Property

Public Property Get Fornec() As Object
     Set Fornec = objUserControl.Controls("Fornec")
End Property

Public Property Get FornLabel() As Object
     Set FornLabel = objUserControl.Controls("FornLabel")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get TipoDestino() As Object
     Set TipoDestino = objUserControl.Controls("TipoDestino")
End Property

Public Property Get Pais() As Object
     Set Pais = objUserControl.Controls("Pais")
End Property

Public Property Get Estado() As Object
     Set Estado = objUserControl.Controls("Estado")
End Property

Public Property Get CEP() As Object
     Set CEP = objUserControl.Controls("CEP")
End Property

Public Property Get Cidade() As Object
     Set Cidade = objUserControl.Controls("Cidade")
End Property

Public Property Get Bairro() As Object
     Set Bairro = objUserControl.Controls("Bairro")
End Property

Public Property Get Endereco() As Object
     Set Endereco = objUserControl.Controls("Endereco")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get TotalMoedaReal() As Object
     Set TotalMoedaReal = objUserControl.Controls("TotalMoedaReal")
End Property

Public Property Get PrecoUnitarioMoedaReal() As Object
     Set PrecoUnitarioMoedaReal = objUserControl.Controls("PrecoUnitarioMoedaReal")
End Property

Public Property Get DescCompleta() As Object
     Set DescCompleta = objUserControl.Controls("DescCompleta")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get DescProduto() As Object
     Set DescProduto = objUserControl.Controls("DescProduto")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get PrecoUnitario() As Object
     Set PrecoUnitario = objUserControl.Controls("PrecoUnitario")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get PercentDesc() As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")
End Property

Public Property Get RecebForaFaixa() As Object
     Set RecebForaFaixa = objUserControl.Controls("RecebForaFaixa")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get PrecoTotal() As Object
     Set PrecoTotal = objUserControl.Controls("PrecoTotal")
End Property

Public Property Get DataLimite() As Object
     Set DataLimite = objUserControl.Controls("DataLimite")
End Property

Public Property Get AliquotaICMS() As Object
     Set AliquotaICMS = objUserControl.Controls("AliquotaICMS")
End Property

Public Property Get ValorIPIItem() As Object
     Set ValorIPIItem = objUserControl.Controls("ValorIPIItem")
End Property

Public Property Get AliquotaIPI() As Object
     Set AliquotaIPI = objUserControl.Controls("AliquotaIPI")
End Property

Public Property Get PercentMaisReceb() As Object
     Set PercentMaisReceb = objUserControl.Controls("PercentMaisReceb")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get OutrasDespesas() As Object
     Set OutrasDespesas = objUserControl.Controls("OutrasDespesas")
End Property

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get ValorIPI() As Object
     Set ValorIPI = objUserControl.Controls("ValorIPI")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
End Property

Public Property Get BotaoEmail() As Object
     Set BotaoEmail = objUserControl.Controls("BotaoEmail")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
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

Public Property Get BotaoLiberaBloqueio() As Object
     Set BotaoLiberaBloqueio = objUserControl.Controls("BotaoLiberaBloqueio")
End Property

Public Property Get SSFrame1() As Object
     Set SSFrame1 = objUserControl.Controls("SSFrame1")
End Property

Public Property Get TipoBloqueio() As Object
     Set TipoBloqueio = objUserControl.Controls("TipoBloqueio")
End Property

Public Property Get CodUsuario() As Object
     Set CodUsuario = objUserControl.Controls("CodUsuario")
End Property

Public Property Get ResponsavelBL() As Object
     Set ResponsavelBL = objUserControl.Controls("ResponsavelBL")
End Property

Public Property Get DataBloqueio() As Object
     Set DataBloqueio = objUserControl.Controls("DataBloqueio")
End Property

Public Property Get ResponsavelLib() As Object
     Set ResponsavelLib = objUserControl.Controls("ResponsavelLib")
End Property

Public Property Get DataLiberacao() As Object
     Set DataLiberacao = objUserControl.Controls("DataLiberacao")
End Property

Public Property Get GridBloqueios() As Object
     Set GridBloqueios = objUserControl.Controls("GridBloqueios")
End Property

Public Property Get BotaoCcl() As Object
     Set BotaoCcl = objUserControl.Controls("BotaoCcl")
End Property

Public Property Get BotaoAlmoxarifados() As Object
     Set BotaoAlmoxarifados = objUserControl.Controls("BotaoAlmoxarifados")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get UnidMed() As Object
     Set UnidMed = objUserControl.Controls("UnidMed")
End Property

Public Property Get DescProd() As Object
     Set DescProd = objUserControl.Controls("DescProd")
End Property

Public Property Get Prod() As Object
     Set Prod = objUserControl.Controls("Prod")
End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get Quant() As Object
     Set Quant = objUserControl.Controls("Quant")
End Property

Public Property Get CentroCusto() As Object
     Set CentroCusto = objUserControl.Controls("CentroCusto")
End Property

Public Property Get Almoxarifado() As Object
     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
End Property

Public Property Get GridDistribuicao() As Object
     Set GridDistribuicao = objUserControl.Controls("GridDistribuicao")
End Property

Public Property Get BotaoContasContabeis() As Object
     Set BotaoContasContabeis = objUserControl.Controls("BotaoContasContabeis")
End Property

Public Property Get BotaoIncluirNota() As Object
     Set BotaoIncluirNota = objUserControl.Controls("BotaoIncluirNota")
End Property

Public Property Get Nota() As Object
     Set Nota = objUserControl.Controls("Nota")
End Property

Public Property Get Frame10() As Object
     Set Frame10 = objUserControl.Controls("Frame10")
End Property

Public Property Get NotaPC() As Object
     Set NotaPC = objUserControl.Controls("NotaPC")
End Property

Public Property Get GridNotas() As Object
     Set GridNotas = objUserControl.Controls("GridNotas")
End Property

Public Property Get LabelNota() As Object
     Set LabelNota = objUserControl.Controls("LabelNota")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get CodigoPV() As Object
     Set CodigoPV = objUserControl.Controls("CodigoPV")
End Property

Public Property Get UpDownDataRefFluxo() As Object
     Set UpDownDataRefFluxo = objUserControl.Controls("UpDownDataRefFluxo")
End Property

Public Property Get DataRefFluxo() As Object
     Set DataRefFluxo = objUserControl.Controls("DataRefFluxo")
End Property

Public Property Get ImprimePedido() As Object
     Set ImprimePedido = objUserControl.Controls("ImprimePedido")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Almoxarifado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Almoxarifado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)
End Sub

Public Sub Almoxarifado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)
End Sub

Public Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CentroCusto
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoAlmoxarifados_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colSelecao As New Collection
Dim iFilialEmpresa As Integer
Dim objEstoqueProdutoTerceiros As New ClassEstoqueProduto

On Error GoTo Erro_BotaoALmoxarifados_Click

    'Verifica se existe alguma linha selecionada
    If GridDistribuicao.Row = 0 Then gError 89434

    If iFrameTipoDestinoAtual <> TIPO_DESTINO_EMPRESA Then Exit Sub

    'Verifica se o produto da linha esta preenchido
    If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col))) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 56052
        
        iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)
        If Len(Trim(iFilialEmpresa)) = 0 Then gError 31152

        colSelecao.Add sProdutoFormatado
        colSelecao.Add iFilialEmpresa

        Call Chama_Tela("AlmoxarifadoFilialLista", colSelecao, objEstoqueProdutoTerceiros, objEventoAlmoxarifados)

    End If

    Exit Sub

Erro_BotaoALmoxarifados_Click:

    Select Case gErr

        Case 56052

        Case 31152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", gErr)

        Case 89434
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164556)

    End Select

    Exit Sub

End Sub

Public Sub BotaoCcl_Click()

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCcl_Click

    'Se alguma linha do Grid Distribuicao esta selecionada
    If (GridDistribuicao.Row <> 0) Then

        'Verifica se o Produto foi preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col))) = 0 Then gError 68336

        'Verifica se o Ccl esta preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_CentroCusto_Col))) > 0 Then

            'Formata o Ccl
            lErro = CF("Ccl_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_CentroCusto_Col), sCclFormata, iCclPreenchida)
            If lErro <> SUCESSO Then gError 56053

            'Coloca Ccl formatada em objCcl.sCcl
            objCcl.sCcl = sCclFormata

        End If

        'Chama tela CclLista
        Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoBotaoCcl)

    End If

    Exit Sub

Erro_BotaoCcl_Click:

    Select Case gErr

        Case 56053
        
        Case 68336
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164557)

    End Select

    Exit Sub

End Sub

Public Sub BotaoContasContabeis_Click()

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim sContaFormata As String
Dim iContaPreenchida As Integer
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoContasContabeis_Click

    'Verifica s alguma linha do grid foi selecionada
    If (GridDistribuicao.Row = 0) Then Exit Sub

    'Verifica se o Produto foi preenchido
    If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col))) = 0 Then gError 68337

    'Verifica se ContaContabil foi preenchida
    If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col))) > 0 Then

        'Formata a conta contabil
        lErro = CF("Conta_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col), sContaFormata, iContaPreenchida)
        If lErro <> SUCESSO Then gError 56055

        'Coloca contaFormatada em objPlanoConta
        objPlanoConta.sConta = sContaFormata

    End If

    'Chama tela PlanoContaEstLista
    Call Chama_Tela("PlanoContaEstLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoContasContabeis_Click:

    Select Case gErr

        Case 56055

        Case 68337
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164558)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEmail_Click()

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio
Dim sMailTo As String, sFiltro As String, sInfoEmail As String
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_BotaoEmail_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 86122
    
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56095

    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76036
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76037
    
    If objPedidoCompra.dtDataRegAprov = DATA_NULA Then
        If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
    End If
        
    'Verifica se existem bloqueios para o Pedido de compra em questao
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76055
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76049
    
    Next
    
    If objPedidoCompra.lFornecedor <> 0 And objPedidoCompra.iFilial <> 0 Then

        objFilialFornecedor.lCodFornecedor = objPedidoCompra.lFornecedor
        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 129314
         
        If lErro = SUCESSO Then
        
            objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError 129315
        
            sMailTo = objEndereco.sEmail
            
        End If
        
        sInfoEmail = "Fornecedor: " & CStr(objFilialFornecedor.lCodFornecedor) & " - " & Fornecedor.Text & " . Filial: " & Filial.Text
        
    End If
    
    If Len(Trim(sMailTo)) = 0 Then gError 129316
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yy")

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra, True)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 56096

    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra", sFiltro)
    If lErro <> SUCESSO Then gError 76038
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Compra", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Compra " & CStr(objPedidoCompra.lCodigo), "TALIASATTACH", "PedCompra" & CStr(objPedidoCompra.lCodigo), "TINFO_EMAIL", sInfoEmail)
    If lErro <> SUCESSO Then gError 76038
    
    Exit Sub

Erro_BotaoEmail_Click:

    Select Case gErr

        Case 56096, 76036, 76038, 76055, 129314, 129315
        
        Case 76037
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
                
        Case 76049
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
        
        Case 86122
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
        
        Case 129316
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_NAO_ENCONTRADO", gErr, objPedidoCompra.lCodigo)
            
        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164559)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 56088

    objPedidoCompra.iFilialEmpresa = giFilialEmpresa
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Text)

    'Busca no BD o Pedido de Compra Baixado
    lErro = CF("PedidoCompraBaixado_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56137 Then Error 56138
    If lErro = SUCESSO Then Error 56414

    'Busca no BD o Pedido de Compra
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then Error 56143
    If lErro = 56142 Then Error 56416

    'Confirma a exclusao do Pedido de Compra
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PEDIDO_COMPRA", Codigo.Text)

    'Se a resposta for negativa
    If vbMsgRes = vbYes Then

        'Exclui o Pedido de Compra
        lErro = CF("PedidoCompra_Exclui", objPedidoCompra)
        If lErro <> SUCESSO Then Error 56132
    
        'Limpa a tela
        Call Limpa_Tela_PedidoCompras
    
        'Fecha o comando das setas, se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)
    
        iAlterado = 0

    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 56088
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", Err)

        Case 56132, 56133, 56138
            'Erro tratado na rotina chamada

        Case 56143
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_JA_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case 56414
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_BAIXADO_EXCLUSAO", Err, objPedidoCompra.lCodigo)

        Case 56416
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164560)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Sub Limpa_Tela_PedidoCompras()

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridBloqueio)
    Call Grid_Limpa(objGridDistribuicao)
    Call Grid_Limpa(objGridNotas) '??? leo

    'Limpa os outros campos da tela
    DataAlteracao.Caption = ""
    DataEmissao.Caption = ""
    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    Filial.Text = ""
    CondPagto.Text = ""
    TabelaPreco.Text = ""
'leo
    Moeda.ListIndex = -1

    Codigo.promptinclude = False
    Codigo.Text = ""
    Codigo.promptinclude = True

    Data.promptinclude = False
    Data.Text = Format(gdtDataHoje, "dd/mm/yy")
    Data.promptinclude = True

    'Limpa  combobox
    FilialFornec.Clear
    TipoFrete.ListIndex = -1
    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True

    ImprimePedido.Value = vbUnchecked
    
    Set gcolItemPedido = New Collection
    Set gobjPC = New ClassPedidoCompras

    ObsEmbalagem.Text = ""

    CodigoPV.promptinclude = False
    CodigoPV.Text = ""
    CodigoPV.promptinclude = True
    
    Set gobjInfoAdic = Nothing

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    gbCarregandoTela = False
    
    Exit Sub

End Sub

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 56269

        'Descrição do Produto
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_Descricao(objGridInt)
            If lErro <> SUCESSO Then gError 86176
            
        
        'UnidadeMed
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UnidadeMed(objGridInt)
            If lErro <> SUCESSO Then gError 56270

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 56271

        'Preco Unitario
        Case iGrid_PrecoUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 56272

        'PercentDesc
        Case iGrid_PercentDesc_Col
            lErro = Saida_Celula_PercDesc(objGridInt)
            If lErro <> SUCESSO Then gError 56273

        'Desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 56274

        'DataLimite
        Case iGrid_DataLimite_Col
            lErro = Saida_Celula_DataLimite(objGridInt)
            If lErro <> SUCESSO Then gError 56277

        'AliquotaIPI
        Case iGrid_AliquotaIPI_Col
            lErro = Saida_Celula_AliquotaIPI(objGridInt)
            If lErro <> SUCESSO Then gError 56278

        'ValorIPIItem
        Case iGrid_ValorIPIItem_Col
            lErro = Saida_Celula_ValorIPI(objGridInt)
            If lErro <> SUCESSO Then gError 56279

        'AliquotaICMS
        Case iGrid_AliquotaICMS_Col
            lErro = Saida_Celula_AliquotaICMS(objGridInt)
            If lErro <> SUCESSO Then gError 56280

        'PercentMaisReceb
        Case iGrid_PercentMaisReceb_Col
            lErro = Saida_Celula_PercentMaisReceb(objGridInt)
            If lErro <> SUCESSO Then gError 56281

        'RecebForaFaixa
        Case iGrid_RecebForaFaixa_Col
            lErro = Saida_Celula_RecebForaFaixa(objGridInt)
            If lErro <> SUCESSO Then gError 56283

        'Observacao
        Case iGrid_Observacao_Col
            lErro = Saida_Celula_Observacao(objGridInt)
            If lErro <> SUCESSO Then gError 56284

        'DescProduto
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescProduto(objGridInt)
            If lErro <> SUCESSO Then gError 56373
            
        Case iGrid_DeliveryDate_Col
            lErro = Saida_Celula_DeliveryDate(objGridInt)
            If lErro <> SUCESSO Then gError 56373
            
        Case iGrid_TempoTransito_Col
            lErro = Saida_Celula_TempoTransito(objGridInt)
            If lErro <> SUCESSO Then gError 56373
            
        '###########################################################
        'Inserido por Wagner 17/07/2006
        Case Else
            lErro = CF2(Me, "PC_Saida_Celula_GridItens", objGridInt)
            If lErro <> SUCESSO Then gError 181213
        '###########################################################

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 56269 To 56274, 56277 To 56284, 56373, 86176, 181213

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164561)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridBloqueios(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridBloqueios

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'TipoBloqueio
        Case iGrid_TipoBloqueio_Col
            lErro = Saida_Celula_TipoBloqueio(objGridInt)
            If lErro <> SUCESSO Then Error 56285

        'ResponsavelBL
        Case iGrid_ResponsavelBL_Col
            lErro = Saida_Celula_Responsavel(objGridInt)
            If lErro <> SUCESSO Then Error 56364

    End Select

    Saida_Celula_GridBloqueios = SUCESSO

    Exit Function

Erro_Saida_Celula_GridBloqueios:

    Saida_Celula_GridBloqueios = Err

    Select Case Err

        Case 56285, 56364

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164562)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDistribuicao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDistribuicao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Prod
        Case iGrid_Prod_Col
            lErro = Saida_Celula_Prod(objGridInt)
            If lErro <> SUCESSO Then Error 56291

        'CentroCusto
        Case iGrid_CentroCusto_Col
            lErro = Saida_Celula_CentroCusto(objGridInt)
            If lErro <> SUCESSO Then Error 56293

        'Almoxarifado
        Case iGrid_Almoxarifado_Col
            lErro = Saida_Celula_Almoxarifado(objGridInt)
            If lErro <> SUCESSO Then Error 56294

        'Quant
        Case iGrid_Quant_Col
            lErro = Saida_Celula_Quant(objGridInt)
            If lErro <> SUCESSO Then Error 56296

        'ContaContabil
        Case iGrid_ContaContabil_Col
            lErro = Saida_Celula_ContaContabil(objGridInt)
            If lErro <> SUCESSO Then Error 56297

    End Select

    Saida_Celula_GridDistribuicao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDistribuicao:

    Saida_Celula_GridDistribuicao = Err

    Select Case Err

        Case 56291, 56293 To 56297

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164563)

    End Select

    Exit Function

End Function
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 56265

            'se for o GridBloqueios
            Case GridBloqueios.Name

                lErro = Saida_Celula_GridBloqueios(objGridInt)
                If lErro <> SUCESSO Then gError 56266

            'se for o GridDistribuicao
            Case GridDistribuicao.Name

                lErro = Saida_Celula_GridDistribuicao(objGridInt)
                If lErro <> SUCESSO Then gError 56267

            Case GridNotas.Name '??? por leo
                
                lErro = Saida_Celula_GridNotas(objGridInt)
                If lErro <> SUCESSO Then gError 103329
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 56268

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 56265, 56266, 56267, 56268, 103329
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164564)

    End Select

    Exit Function

End Function

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 49451

    'Limpa a tela
    Call Limpa_Tela_PedidoCompras

    'Fecha o comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 49451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164565)

    End Select

    Exit Sub

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio, sFiltro As String

On Error GoTo Erro_BotaoImprimir_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 86122
    
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56095

    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76036
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76037
    
'    If objPedidoCompra.dtDataRegAprov = DATA_NULA Then
'        If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
'    End If
        
    'Verifica se existem bloqueios para o Pedido de compra em questao
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76055
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76049
    
    Next
    
    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 56096

    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yy")

    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra", sFiltro)
    If lErro <> SUCESSO Then gError 76038
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Compra", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76038
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 56096, 76036, 76038, 76055
        
        Case 76037
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
                
        Case 76049
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
        
        Case 86122
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
        
        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164566)

    End Select

    Exit Sub

End Sub
'??? por leo
Public Sub BotaoIncluirNota_Click()

    If Len(Trim(Nota.Text)) > 0 Then
    
        objGridNotas.iLinhasExistentes = objGridNotas.iLinhasExistentes + 1
        
        GridNotas.TextMatrix(objGridNotas.iLinhasExistentes, iGrid_NotaPC_Col) = Trim(Nota.Text)
        
        Nota.Text = ""
        
    End If

End Sub

Public Sub BotaoLiberaBloqueio_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_BotaoLiberaBloqueio_Click

    'Verifica se o número do Pedido de Compra está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 49499

    'Chama tela LiberaBloqueioPC
    Call Chama_Tela("LiberaBloqueioPC", objPedidoCompra)

    Exit Sub

Erro_BotaoLiberaBloqueio_Click:

    Select Case Err

        Case 49499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164567)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_PercDesc(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim bPercentDescIgual As Boolean
Dim dValorIPI As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_PercDesc

    Set objGridInt.objControle = PercentDesc

    bPercentDescIgual = False

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then

        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then Error 56030

        dPercentDesc = StrParaDbl(PercentDesc.Text)

    End If

    'Verifica se percentual foi alterado
    If PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col)) = StrParaDbl(PercentDesc.Text) Then bPercentDescIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56032

    If bPercentDescIgual = False Then

        'Verifica se o percentual é de 100%
        If dPercentDesc = 100# Then Error 56031

        'Coloca PercentDesc no formato Fixed da tela
        PercentDesc.Text = Format(dPercentDesc, "Fixed")

        'Verifica se preco unitario e  quantidade foram preenchidos
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 Then
            
'            dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'            Call CF2(Me, "PC_PrecoTotal_Calcula", GridItens.Row, dQuantidade) 'Inserido por Wagner 17/07/2006
'
'            dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
'
'            'Calcula valor total
'            dValorTotal = dQuantidade * dPrecoUnitario
'            dDesconto = (dValorTotal * dPercentDesc) / 100
'            dValorTotal = dValorTotal - dDesconto
'
'            'Coloca valor do Desconto no GridItens
'            If dDesconto > 0 Then
'                GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dDesconto, "Fixed")
'            Else
'                GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
'            End If
            
            Call PrecoTotal_Calcula(GridItens.Row) 'Inserido por Wagner 17/07/2006
            
'            'coloca o Preco Total no grid
'            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal, PrecoTotal.Format) 'Alterado por Wagner
'
'            'Verifica se ALiquotaIPI foi preenchida
'            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col))) > 0 Then
'
'                'Calcula valorIPI
'                dValorIPI = (dValorTotal) * (PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col)))
'                If dValorIPI > 0 Then
'                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "Standard")
'                Else
'                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""
'                End If
'
'                dValorTotal = 0
'
'                'Calcula o valorIPI da tela
'                For iIndice = 1 To objGridItens.iLinhasExistentes
'                    dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
'                    dValorTotal = dValorTotal + dValorIPI
'                Next
'
'                ValorIPI.Text = Format(dValorTotal, "Standard")
'
'            End If

            'Calcula Valor Total
            lErro = ValorTotal_Calcula()
            If lErro <> SUCESSO Then Error 56126

        End If

    End If
    
    If Len(Trim(Taxa.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Text))

    Saida_Celula_PercDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercDesc:

    Saida_Celula_PercDesc = Err

    Select Case Err

        Case 56030, 56032, 56126
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56031
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164568)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaEnxuta As String

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    'Se o Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 56033

        'Valida o Almoxarifado
        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 56034

        'Se não for encontrado --> Erro
        If lErro = 25130 Then gError 56035
        If lErro = 25135 Then gError 56036
        If objAlmoxarifado.iFilialEmpresa <> Codigo_Extrai(FilialEmpresa.Text) Then gError 86120

        'Se a conta contábil do grid não estiver preenchida
        If Len(Trim(GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col))) = 0 Then
            'Coloca Conta Contábil no GridDistribuicao
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 49991 Then gError 68386
    
            If lErro <> 49991 Then
    
                lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
                If lErro <> SUCESSO Then gError 68387
    
                ContaContabil.promptinclude = False
                ContaContabil.Text = sContaEnxuta
                ContaContabil.promptinclude = True
    
                GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
            End If

        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 56037

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 56037, 56033, 56034, 68383, 68387
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56035
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

        Case 56036

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

        Case 86120
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_FILIALENTREGA", gErr, objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido, Codigo_Extrai(FilialEmpresa.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164569)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim sProduto As String
Dim bQuantidadeIgual As Boolean
Dim objProduto As New ClassProduto
Dim iPreenchido As Integer
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorIPI As Double
Dim dValorIPITotal As Double
Dim dPrecoTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    Quantidade = Trim(Quantidade)
    
    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 56038

        dQuantidade = StrParaDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)

    End If

    'Verifica se a quantidade foi alterada
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)) = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 56039

    'Se quantidade foi alterada e Preco Unitario foi preenchido
    If bQuantidadeIgual = False Then

        If iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 56205

            'Guarda o Codigo do Produto em objProduto
            objProduto.sCodigo = sProduto

            'Le o produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 56381
            'Se nao encontrou => erro
            If lErro = 28030 Then gError 56437

            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

                'Atualiza Grid Distribuicao
                lErro = Distribuicao_Atualiza(GridItens.Row, objProduto)
                If lErro <> SUCESSO Then gError 56206

            End If

        End If

        'Verifica se quantidade foi preenchida
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 Then

'            dDesconto = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col))
'
'            'Verifica se PercentDesc foi preenchido
'            If dDesconto <> 0 Then
'
'                'Calcula valor do desconto
'                dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'                Call CF2(Me, "PC_PrecoTotal_Calcula", GridItens.Row, dQuantidade) 'Inserido por Wagner 17/07/2006
'
'                dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
'                dDesconto = (dQuantidade * dPrecoUnitario) * (PercentParaDbl((GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col))))
'                GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
'
'            End If

            '################################################
            'Alterado por Wagner 14/07/2006
'            'Calcula o preco total
'            dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
'            dPrecoTotal = (StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) - dDesconto) * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
            Call PrecoTotal_Calcula(GridItens.Row)
            '################################################
                        
            If Len(Trim(Taxa.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Text))

'            'Verifica se AliquotaIPI foi preenchida
'            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col))) <> 0 Then
'
'                'Calcula valorIPI
'                dValorIPI = dPrecoTotal * (PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col)))
'                If dValorIPI > 0 Then
'                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "standard")
'                Else
'                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""
'                End If
'
'            End If

        Else 'se quantidade =0

            'Limpa as linhas do grid
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = ""
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""

        End If

        'Calcula o valorIPI da tela
        For iIndice = 1 To objGridItens.iLinhasExistentes

            dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
            dValorIPITotal = dValorIPITotal + dValorIPI

        Next

        ValorIPI.Text = Format(dValorIPITotal, "standard") 'Ok Você que colocr o valorIPI Total

        '#############################################################
        'Inserido por Wagner 23/03/2006
        lErro = CF2(Me, "PC_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177161
        '#############################################################

    End If

    'Calcula Valor Total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56207

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 56038, 56039, 56205, 56206, 56207, 56381, 177161
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56437
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164570)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataLimite(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtData As Date
Dim dtDataLimite As Date

On Error GoTo Erro_Saida_Celula_DataLimite

    Set objGridInt.objControle = DataLimite

    'Verifica se a data limite foi preenchida
    If Len(Trim(DataLimite.ClipText)) > 0 Then

        'Critica a data limite
        lErro = Data_Critica(DataLimite.Text)
        If lErro <> SUCESSO Then Error 56078

        dtDataLimite = CDate(DataLimite.Text)

        'Verifica se a data do pedido foi preenchida
        If Len(Trim(Data.ClipText)) > 0 Then

            dtData = CDate(Data.Text)

            'comparacao entre data limite e data do pedido
            If dtDataLimite < dtData Then Error 56079

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56080

    Saida_Celula_DataLimite = SUCESSO

    Exit Function

Erro_Saida_Celula_DataLimite:

    Saida_Celula_DataLimite = Err

    Select Case Err

        Case 56078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56079
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INFERIOR_DATAPEDIDO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56080
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164571)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoTotal As Double
Dim dValorDesconto As Double
Dim dPrecoUnitario As Double
Dim dValorIPI As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim bDescontoIgual As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    bDescontoIgual = False

    'Se estiver preenchido
    If Len(Trim(Desconto.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then Error 56000

        dValorDesconto = CDbl(Desconto.Text)

        'Coloca Desconto no formato Standard da tela
        Desconto.Text = Format(dValorDesconto, "standard")

    End If

    'Verifica se o desconto foi alterado
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) = StrParaDbl(Desconto.Text) Then bDescontoIgual = True


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56002

    'Se o Desconto foi alterado
    If bDescontoIgual = False Then

        'Verifica se quantidade e PrecoUnitario foram preenchidos
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 Then

            'Tira o percentual de desconto para não recalcular o desconto
            GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col) = ""
            Call PrecoTotal_Calcula(GridItens.Row) 'Inserido por Wagner 17/07/2006
'
'            dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'            Call CF2(Me, "PC_PrecoTotal_Calcula", GridItens.Row, dQuantidade) 'Inserido por Wagner 17/07/2006
'
'            dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))

            'Calcula Preco Total
            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col)) + dValorDesconto

            'Verifica se ValorDesconto é maior que Preco Total
            If dValorDesconto >= dPrecoTotal Then Error 56001

            If dPrecoTotal > 0 Then

                'Calcula PercentDesc
                dPercentDesc = dValorDesconto / (dPrecoTotal)
                GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col) = Format(dPercentDesc, "Percent")
                'Calcula Preco Total
'                dPrecoTotal = dPrecoTotal - dValorDesconto
'                GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
'
'                'Verifica se AliquotaIPI foi preenchida
'                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col))) > 0 Then
'
'                    'Calcula valorIPI
'                    dValorIPI = dPrecoTotal * (PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col)))
'                    If dValorIPI > 0 Then
'                        GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "Standard")
'                    Else
'                        GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""
'                    End If
'
'                    'Calcula o valorIPI da tela
'                    For iIndice = 1 To objGridItens.iLinhasExistentes
'
'                        dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
'                        dValorTotal = dValorTotal + dValorIPI
'
'                    Next
'
'                    ValorIPI.Text = Format(dValorTotal, "Standard")
'
'                End If

                'Calcula Valor Total
                lErro = ValorTotal_Calcula()
                If lErro <> SUCESSO Then Error 56120

            End If

        End If

    End If
    
    If Len(Trim(Taxa.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Text))

    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = Err

    Select Case Err

        Case 56000, 56120
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56001
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56002
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164572)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CentroCusto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_CentroCusto

    Set objGridInt.objControle = CentroCusto

    'Verifica se CentroCusto foi preenchido
    If Len(Trim(CentroCusto.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", CentroCusto, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then Error 56003

        'Se o Ccl nao existe
        If lErro = 5703 Then Error 56004

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56005

    Saida_Celula_CentroCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_CentroCusto:

    Saida_Celula_CentroCusto = Err

    Select Case Err

        Case 56003
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56004
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, CentroCusto.Text)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56005
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164573)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 49427

    'Limpa a tela
    Call Limpa_Tela_PedidoCompras

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iGridBloqueiosLinhasBloqueadas = 0

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 49427

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164574)

    End Select

    Exit Sub

End Sub

Public Sub BotaoPedidosAvulsos_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedidosAvulsos_Click

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 49493

    'Chama a tela PedComprasNaoEnvLista
    Call Chama_Tela("PedComprasNaoEnvLista", colSelecao, objPedidoCompra, objEventoBotaoPedidosAvulsos)

    Exit Sub

Erro_BotaoPedidosAvulsos_Click:

    Select Case Err

        Case 49493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164575)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sProduto1 As String
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_BotaoProdutos_Click

    If GridItens.Row = 0 Then gError 89435

    'Verifica se o Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 68393
        'Verifica se a Filial do Fornecedor está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 68392
            
    objFornecedor.sNomeReduzido = Fornecedor.Text

    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 68388
    If lErro = 6681 Then gError 68389 'Se nao encontrar --> erro

    colSelecao.Add objFornecedor.lCodigo

    objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Text)
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

    lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 12929 Then gError 68390

    'Se não encontrou a Filial do Fornecedor ==> erro
    If lErro = 12929 Then gError 68391

    colSelecao.Add objFilialFornecedor.iCodFilial
    
    '###############################################
    'Inserido por Wagner 05/05/06
    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 177415
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
    
    objProduto.sCodigo = sProduto
    '###############################################

    'Chama a tela ProdutoFilialFornLista
    Call Chama_Tela("ProdutoFilialFornLista", colSelecao, objProduto, objEventoProdutos)
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 49495, 68388, 68391, 177415

        Case 68389
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case 68390
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFornecedor.lCodigo)

        Case 68392
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 68393
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 89435
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164576)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()
'Gera o próximo número de Pedido de Compras

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera o próximo código para Pedido de Compras
    lErro = CF("PedidoCompra_Codigo_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 63614

    'Coloca o código gerado na tela
    Codigo.promptinclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.promptinclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 61614
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164577)

    End Select

    Exit Sub

End Sub

Public Sub BotaoTrazCotacao_Click()
'Traz a última cotação da moeda selecionada

Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoTrazCotacao_Click

    'Carrega objCotacao
    objCotacao.dtData = gdtDataAtual
    
    'Se a moeda não foi selecionada => Erro
    If Len(Trim(Moeda.Text)) = 0 Then gError 108943
        
    'Preeche com a Moeda selecionada
    objCotacao.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
    objCotacaoAnterior.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 108944
    
    'Se nao existe cotacao para a data informada => Mostra a última.
    Taxa.Text = IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000"))
    
    Call ComparativoMoedaReal_Calcula(StrParaDbl(Taxa.Text))

    Exit Sub
    
Erro_BotaoTrazCotacao_Click:

    Select Case gErr
    
        Case 108943
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            '??? Falta cadastrar: ERRO_MOEDA_NAO_PREENCHIDA - "Para trazer a cotação a moeda deve ser selecionada antes."
            
        Case 108944
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164578)
    
    End Select
    
End Sub

Public Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 26513

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 26513

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164579)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()

    gsCodigoAnterior = Codigo.Text
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Public Sub CodigoLabel_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_CodigoLabel_Click

    'Move os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 49441

    'Chama a Tela de browse
    Call Chama_Tela("PedComprasNaoEnvLista", colSelecao, objPedidoCompra, objEventoCodigo)

    Exit Sub

Erro_CodigoLabel_Click:

    Select Case Err

        Case 49441

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164580)

    End Select

    Exit Sub

End Sub

Public Sub CondPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondPagto.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondPagto.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56068

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 56069
        If lErro = 19205 Then Error 56070

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmPagamento = 0 Then Error 56071

        'Coloca na Tela
        CondPagto.Text = CondPagto_Traz(objCondicaoPagto)

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then Error 56072

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case Err

       Case 56068, 56069

       Case 56070
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)
            If vbMsgRes = vbYes Then Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

        Case 56071
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", Err, objCondicaoPagto.iCodigo)

        Case 56072
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164581)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Verifica se condicao de pagamento esta preenchida
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondPagto)

    'Chama tela de browse
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

    Exit Sub

End Sub

Public Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'Critica a data informada
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then Error 49442

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case Err

        Case 49442

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164582)

    End Select

    Exit Sub

End Sub

Public Sub DataEnvio_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)
    
End Sub

Public Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataEnvio As Date
Dim dtData As Date

On Error GoTo Erro_DataEnvio_Validate

    'Verifica se a data foi preenchida
    If Len(DataEnvio.ClipText) = 0 Then Exit Sub

    'Critica a data informada
    lErro = Data_Critica(DataEnvio.Text)
    If lErro <> SUCESSO Then Error 49445

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True

    Select Case Err

        Case 49445

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164583)

    End Select

    Exit Sub

End Sub

Private Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DescProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DescProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Embalagem_Change()
'leo
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Embalagem_Validate(Cancel As Boolean)
'leo
Dim objEmbalagem As New ClassEmbalagem
Dim lErro As Long

On Error GoTo Erro_Embalagem_Validate
        
    If Len(Trim(Embalagem.Text)) = 0 Then Exit Sub
    
    'se o texto é numérico ou se está no formato Código + Separador + Sigla
    If IsNumeric(Embalagem.Text) = True Or Codigo_Extrai(Embalagem.Text) <> 0 Then
        
        If IsNumeric(Embalagem.Text) Then
        
            objEmbalagem.iCodigo = StrParaInt(Embalagem.Text)
        
        Else
            
            objEmbalagem.iCodigo = Codigo_Extrai(Embalagem.Text)
            
        End If
            
        'Lê a embalagem através do Código
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 103363
        
        If lErro <> SUCESSO Then gError 103364
    
    'Se é uma String
    Else
        
        objEmbalagem.sSigla = Embalagem.Text
    
        'Lê a Embalagem através da sigla
        lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 95088 Then gError 103366
        
        If lErro <> SUCESSO Then gError 103365
        
    End If
    
    Embalagem.Text = objEmbalagem.sSigla
    
    Exit Sub
    
Erro_Embalagem_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 103363, 103366
            
        Case 103364, 103365
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_INEXISTENTE", gErr, Embalagem.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164584)
    
    End Select
    
    Exit Sub

End Sub

Public Sub FilialFornec_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornec_Click

    objFilialFornecedor.iCodFilial = FilialFornec.ItemData(FilialFornec.ListIndex)

    'Busca no BD a FilialFornecedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornec.Text, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then Error 57258

    If lErro = 18272 Then Error 57259

    objEndereco.lCodigo = objFilialFornecedor.lEndereco

    'Le o endereco da FilialFornecedor passada como parametro
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 57260
    'Se nao encontrou o endereco
    If lErro = 12309 Then Error 57261

    'Preenche endereco da FilialFornecedor
    Call Preenche_Endereco(objEndereco)

    Exit Sub

Erro_FilialFornec_Click:

    Select Case Err

        Case 57258, 57260

        Case 57259
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FilialFornec.Text)
            If vbMsgRes = vbYes Then Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)

        Case 57261
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164585)

    End Select

    Exit Sub

End Sub

Public Sub GridDistribuicao_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDistribuicao)

End Sub

Public Sub LabelEmbalagem_Click()
'leo
Dim objEmbalagem As New ClassEmbalagem
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_LabelEmbalagem_Click
        
    objEmbalagem.sSigla = Embalagem.Text
    
    lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 95088 Then gError 103362
    
    'chama a tela de browser
    Call Chama_Tela("EmbalagensLista", colSelecao, objEmbalagem, objEventoEmbalagens)
    
    Exit Sub
    
Erro_LabelEmbalagem_Click:
    
    Select Case gErr
    
        Case 103362
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164586)
    
    End Select
    
    Exit Sub

End Sub

Public Sub LabelNota_Click()

Dim objNotasPC As New ClassNotaPC
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objNotasPC.sNota = Nota.Text

    Call Chama_Tela("NotasPCLista", colSelecao, objNotasPC, objEventoNotaPC)

End Sub

Public Sub Moeda_Change()
'leo
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Moeda_Click()

Dim lErro As Long
Dim iLinhasGrid As Integer
Dim objPedidoCompras As New ClassPedidoCompras

On Error GoTo Erro_Moeda_Click
    
    'Se a moeda selecionada for = REAL
    If Codigo_Extrai(Moeda.List(Moeda.ListIndex)) = MOEDA_REAL Then
    
        'Limpa a cotacao
        Taxa.Text = ""
        
        'Desabilita a cotacao
        Taxa.Enabled = False
        BotaoTrazCotacao.Enabled = False
        
        bExibirColReal = False
        
    Else
            
        'Habilita a cotacao
        Taxa.Enabled = True
        BotaoTrazCotacao.Enabled = True
        
        bExibirColReal = True
    
    End If
    
    lErro = Move_GridItens_Memoria(objPedidoCompras)
    If lErro <> SUCESSO Then gError 114660
    
    Set objGridItens = New AdmGrid
    
    Call Inicializa_Grid_Itens(objGridItens)
    
    lErro = Preenche_Grid_Itens(objPedidoCompras)
    If lErro <> SUCESSO Then gError 114661
    
    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 114661
    
    Exit Sub
    
Erro_Moeda_Click:

    Select Case gErr
    
        Case 114660, 114661
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164587)
            
    End Select
    
End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoNotaPC_evSelecao(obj1 As Object)

Dim objNotasPC As ClassNotaPC

    Set objNotasPC = obj1

    'Coloca a Nota na tela
    Nota.Text = objNotasPC.sNota

    Me.Show

    Exit Sub

End Sub

Public Sub Nota_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Taxa_Change()
'leo
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Taxa_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Taxa_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(Taxa.Text)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(Taxa.Text)
        If lErro <> SUCESSO Then gError 103378
    
        'Põe o valor formatado na tela
        Taxa.Text = Format(Taxa.Text, FORMATO_TAXA_CONVERSAO_MOEDA)
        
        'Calcula o comparativo em real para o grid de itens
        Call ComparativoMoedaReal_Calcula(StrParaDbl(Taxa.Text))
    
    End If

    Exit Sub

Erro_Taxa_Validate:

    Cancel = True

    Select Case gErr

        Case 103378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164588)

    End Select

    Exit Sub

End Sub

Public Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub PercentDesc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub DataLimite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub DataLimite_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataLimite_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataLimite_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataLimite
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Filial_Click()

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer

On Error GoTo Erro_Filial_Click

    'Verifica se a filial foi selecionada
    If Filial.ListIndex <> -1 Then

        'Verifica se o Tipo Frete foi selecionado
        If TipoFrete.ListIndex = -1 Then

            sFornecedor = Fornecedor.Text

            objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Text)

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then Error 56298

            'Se nao encontrou
            If lErro = 18272 Then Error 56438

            'Coloca na tela o tipo frete default
            For iIndice = 0 To TipoFrete.ListCount - 1
                If TipoFrete.ItemData(iIndice) = objFilialFornecedor.iTipoFrete Then
                    TipoFrete.ListIndex = iIndice
                End If
            Next

        End If

    End If

    Exit Sub

Erro_Filial_Click:

    Select Case Err

        Case 56298

        Case 56438
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", objFilialFornecedor.iCodFilial, Fornecedor.Text)

            If vbMsgRes = vbYes Then Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164589)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Se a filial nao estiver preenchida => sai da rotina
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56063

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 56064

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 56065

        If lErro = 18272 Then Error 56066

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 56067

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err

        Case 56063, 56065

        Case 56066
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 56064
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 56067
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164590)

    End Select

    Exit Sub

End Sub

Public Sub FilialEmpresa_Click()

Dim objFilialEmpresa As New AdmFiliais
Dim lErro As Long
Dim iCodigo As Integer
Dim objPais As New ClassPais

On Error GoTo Erro_FilialEmpresa_Click

    'Verifica se a filial foi selecionada
    If FilialEmpresa.ListIndex = -1 Then Exit Sub

    'Verifica o tipo destino
    If TipoDestino(TIPO_DESTINO_EMPRESA).Value = True Then

        'Limpa o grid de Distribuicao
        Call Grid_Limpa(objGridDistribuicao)

        'preeenche Codigo da Filial em objFilialEmpresa
        objFilialEmpresa.iCodFilial = Codigo_Extrai(FilialEmpresa.Text)

        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 56386

        'Se nao encontrou => erro
        If lErro = 27378 Then Error 56387

        'Traz endereco para tela
        If objFilialEmpresa.objEnderecoEntrega.lCodigo <> 0 Then
            Call Preenche_Endereco(objFilialEmpresa.objEnderecoEntrega)
        Else
            Call Preenche_Endereco(objFilialEmpresa.objEndereco)
        End If

    End If

    'Verifica se o Grid Itens esta preenchido
    If objGridItens.iLinhasExistentes > 0 Then
    
        If iFilialEmpresaEntregaAnt <> Codigo_Extrai(FilialEmpresa.Text) Then

            Call Distribuicao_Grid_Atualiza
            
        End If

    End If

    iFilialEmpresaEntregaAnt = Codigo_Extrai(FilialEmpresa.Text)

    Exit Sub

Erro_FilialEmpresa_Click:

    Select Case Err

        Case 56386

        Case 56387
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, iCodigo)

        Case 57253
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, objFilialEmpresa.objEndereco.iCodigoPais)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164591)

    End Select

    Exit Sub

End Sub

Public Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se FilialEmpresa foi preenchida
    If Len(Trim(FilialEmpresa.Text)) > 0 Then

        If FilialEmpresa.ListIndex <> -1 Then Exit Sub

        lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56129

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'preeenche objFilialEmpresa
            objFilialEmpresa.iCodFilial = iCodigo

            'Le a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then Error 56195

            'Se nao encontrou => erro
            If lErro = 27378 Then Error 56196

            If lErro = SUCESSO Then

                'Coloca na tela o codigo e o nome da FilialEmpresa
                FilialEmpresa.Text = objFilialEmpresa.lCodEmpresa & SEPARADOR & objFilialEmpresa.sNome

                If objFilialEmpresa.objEnderecoEntrega.lCodigo <> 0 Then
                    Call Preenche_Endereco(objFilialEmpresa.objEnderecoEntrega)
                Else
                    Call Preenche_Endereco(objFilialEmpresa.objEndereco)
                End If

                If iFilialEmpresaEntregaAnt <> Codigo_Extrai(FilialEmpresa.Text) Then
                
                    lErro = Distribuicao_Grid_Atualiza()
                    If lErro <> SUCESSO Then Error 56198
                    
                End If

            End If

        End If

        'Se nao encontrou e nao era codigo
        If lErro = 6731 Then Error 56201

    Else

        'Limpa frame de Endereco
        Call Limpa_Frame_Endereco

    End If
    
    iFilialEmpresaEntregaAnt = Codigo_Extrai(FilialEmpresa.Text)

    Exit Sub

Erro_FilialEmpresa_Validate:

    Select Case Err

        Case 56196
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, iCodigo)

        Case 56129, 56195, 56198, 57256

        Case 56201
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164592)

    End Select

    Exit Sub

End Sub

Function Distribuicao_Grid_Atualiza() As Long

Dim lErro As Long
Dim sProduto As String
Dim iIndice As Integer
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Distribuicao_Grid_Atualiza

'    'Atualiza a Distribuicao
'    lErro = Distribuicao_Atualiza(iIndice, objProduto)
'    If lErro <> SUCESSO Then Error 56174

    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 56173

        objProduto.sCodigo = sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 56439

        'Se nao encontrou
        If lErro = 28030 Then Error 56440

        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            'Atualiza a Distribuicao
            lErro = Distribuicao_Atualiza(iIndice, objProduto)
            If lErro <> SUCESSO Then Error 56174

        End If

    Next

    Distribuicao_Grid_Atualiza = SUCESSO

    Exit Function

Erro_Distribuicao_Grid_Atualiza:

    Distribuicao_Grid_Atualiza = Err

    Select Case Err

        Case 56439, 56173, 56174

        Case 56440
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164593)

    End Select

    Exit Function

End Function

Public Sub FilialFornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim sNomeRed As String
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_FilialFornec_Validate

    'Verifica se FilialFornec esta preenchida
    If Len(Trim(FilialFornec.Text)) > 0 Then

        'Verifica se FilialFornec esta selecionada
        If FilialFornec.ListIndex <> -1 Then Exit Sub

        'Seleciona combo box de FilialFornec
        lErro = Combo_Seleciona(FilialFornec, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56091

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornec.ClipText)) = 0 Then Error 56092

            sNomeRed = Fornec.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornec.Text, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then Error 56175

            If lErro = 18272 Then Error 56177

            'Coloca na tela Codigo e Nome Reduzido de FilialFornec
            FilialFornec.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

            objEndereco.lCodigo = objFilialFornecedor.lEndereco

            'Le o endereco da FilialFornecedor passada como parametro
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then Error 56176
            'Se nao encontrou o endereco
            If lErro = 12309 Then Error 56441

            Call Preenche_Endereco(objEndereco)

        End If

        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then Error 56178

    Else

    'OK endentacao
        'Limpa Frame Endereco
        Call Limpa_Frame_Endereco

    End If

    Exit Sub

Erro_FilialFornec_Validate:

    Cancel = True

    Select Case Err

        Case 56092
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 56175, 57254

        Case 57255
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, objEndereco.iCodigoPais)

        Case 56177
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FilialFornec.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 56178
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, objFilialFornecedor.sNome)

        Case 56441
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164594)

    End Select

    Exit Sub

End Sub

'Esta funcao já existe na tela de PedComprasCons
Private Sub Preenche_Endereco(objEndereco As ClassEndereco)

Dim objPais As New ClassPais
Dim lErro As Long

On Error GoTo Erro_Preenche_Endereco

    Call Limpa_Frame_Endereco
    
    objPais.iCodigo = objEndereco.iCodigoPais

    lErro = CF("Paises_Le", objPais)
    If lErro <> SUCESSO And lErro <> 47876 Then Error 53088
    If lErro = 47876 Then Error 53089

    Endereco.Caption = objEndereco.sEndereco
    Bairro.Caption = objEndereco.sBairro
    Estado.Caption = objEndereco.sSiglaEstado
    Cidade.Caption = objEndereco.sCidade
    Pais.Caption = objPais.sNome
    CEP.Caption = objEndereco.sCEP

    Exit Sub

Erro_Preenche_Endereco:

    Select Case Err

        Case 53088

        Case 53089
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, objEndereco.iCodigoPais)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164595)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'''''FUNCAO JA EXISTENTE EM ROTINASMAT
'''''Quando uma função já está em Rotinas e você consegue enxergar pode usar sem copiar pra tela
''''Function Inicializa_Mascara_Produto_MaskEd(MaskEdProduto As Object) As Long
'''''inicializa a mascara de Produto
''''
''''Dim sMascaraProduto As String
''''Dim lErro As Long
''''
''''On Error GoTo Erro_Inicializa_Mascara_Produto_MaskEd
''''
''''    'Inicializa a máscara de Produto
''''    sMascaraProduto = String(STRING_PRODUTO, 0)
''''
''''    'le a mascara das Produtos
''''    lErro = MascaraProduto(sMascaraProduto)
''''    If lErro Then Error 49461
''''
''''    MaskEdProduto.Mask = sMascaraProduto
''''
''''    Inicializa_Mascara_Produto_MaskEd = SUCESSO
''''
''''    Exit Function
''''
''''Erro_Inicializa_Mascara_Produto_MaskEd:
''''
''''    Inicializa_Mascara_Produto_MaskEd = Err
''''
''''    Select Case Err
''''
''''        Case 49461
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164596)
''''
''''    End Select
''''
''''    Exit Function
''''
''''End Function

Private Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador
Dim objUsuarios As New ClassUsuarios
Dim objFilEmp As New AdmFiliais
Dim iFilialEmpresa As Integer
Dim objObsEmbalagem As Object

On Error GoTo Erro_Form_Load

    gbCarregandoTela = True
    
    bExibirColReal = True
    
    iFrameAtual = 1

    '##################################
    'Inserido por Wagner
    Call Formata_Controles
    '##################################

    objComprador.sCodUsuario = gsUsuario
    Set objEventoCodigo = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCondPagto = New AdmEvento
    Set objEventoObservacao = New AdmEvento
    Set objEventoBotaoPedidosAvulsos = New AdmEvento
    Set objEventoProdutos = New AdmEvento
    Set objEventoFornec = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoAlmoxarifados = New AdmEvento
    Set objEventoBotaoCcl = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objGridNotas = New AdmGrid
    Set objEventoNotaPC = New AdmEvento
    Set objEventoEmbalagens = New AdmEvento 'leo
        
    Set gcolItemPedido = New Collection
    Set gobjPC = New ClassPedidoCompras
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    'Verifica se gsUsuario e comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then gError 56024
    If lErro = 50059 Then gError 56025

    objUsuarios.sCodUsuario = objComprador.sCodUsuario

    'le  o usuário contido na tabela de Usuarios
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 56089
    If lErro <> SUCESSO Then gError 56090

    'Coloca nome reduzido do Comprador na tela
    Comprador.Caption = objUsuarios.sNomeReduzido

    'Carrefa a combo de Condicao de Pagamento
    lErro = CF("Carrega_CondicaoPagamento", CondPagto, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 49431

    'Carrega a combo combo de Tabela de Preços
    lErro = Carrega_TabelaPreco()
    If lErro <> SUCESSO Then gError 49431
    
    'Carrega a combo de Transportadora
    lErro = Carrega_Transportadora()
    If lErro <> SUCESSO Then gError 49452

    'Carrega a combo de Tipos de Bloqueio
    lErro = Carrega_TipoBloqueio()
    If lErro <> SUCESSO Then gError 49464

    'Carrega a combo de RecebForaFaixa
    lErro = Carrega_RecebForaFaixa()
    If lErro <> SUCESSO Then gError 49453

    'Carrega a combo de filial
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then gError 56379
    
    'carrega a combo de Moedas
    lErro = Carrega_Moeda() 'leo
    If lErro <> SUCESSO Then gError 103373
    
    objFilEmp.lCodEmpresa = glEmpresa
    objFilEmp.iCodFilial = giFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilEmp)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 62663
    If lErro <> SUCESSO Then gError 62664
    
    Set colAliqExternas = New Collection
    
    lErro = CF("ICMSAliqExternas_Le_EstadoDestino", objFilEmp.objEndereco.sSiglaEstado, colAliqExternas)
    If lErro <> SUCESSO Then gError 62665

    'Inicializa mascara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 49454

    'Inicializa mascara do Ccl
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 49455

    'Inicializa mascara da conta contabil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 49456

    Quantidade.Format = FORMATO_ESTOQUE

    'Preenche data com data atual
    Data.promptinclude = False
    Data.Text = Format(gdtDataHoje, "dd/mm/yy")
    Data.promptinclude = True

    Set objGridItens = New AdmGrid
    Set objGridBloqueio = New AdmGrid
    Set objGridDistribuicao = New AdmGrid

    'Faz a inicializacao do grid itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 49457

    'Faz a inicializacao do grid bloqueio
    lErro = Inicializa_Grid_Bloqueios(objGridBloqueio)
    If lErro <> SUCESSO Then gError 49459

    'Faz a inicializacao do grid distribuicao
    lErro = Inicializa_Grid_Distribuicao(objGridDistribuicao)
    If lErro <> SUCESSO Then gError 49458

    '??? por leo
    lErro = Inicializa_GridNotas(objGridNotas)
    If lErro <> SUCESSO Then gError 103327
    
    TipoDestino(TIPO_DESTINO_EMPRESA).Value = True
    If gobjCOM.iFilialCompra > 0 Then
        Call CF("Filial_Seleciona", FilialEmpresa, gobjCOM.iFilialCompra)
        iFilialEmpresa = gobjCOM.iFilialCompra
    Else
        Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)
        iFilialEmpresa = giFilialEmpresa
    End If

    lErro = CF("FilialEmpresa_Customiza", iFilialEmpresa)
    If lErro <> SUCESSO Then gError 126950
    
    FilialEmpresa.Text = iFilialEmpresa
    
    Call FilialEmpresa_Validate(bSGECancelDummy)

    Set objObsEmbalagem = ObsEmbalagem

    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPOEMBALAGEM, objObsEmbalagem, False, False)
    If lErro <> SUCESSO Then gError 178000

    gbCarregandoTela = False
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    gbCarregandoTela = False
    
    Select Case gErr

        Case 49431, 49452, 49453, 49454, 49455, 49456, 49457, 49458, 49459, 49464, 56024, 56089, 56379, 62663, 62665, 103373, 126950, 178000

        Case 56025
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", gErr, objComprador.sCodUsuario)

        Case 56090
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)

        Case 62664
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA1", gErr, gsNomeFilialEmpresa)
        
        Case 103327
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164597)

    End Select

    Exit Sub

End Sub

Private Function Carrega_RecebForaFaixa() As Long

    'Limpa a combo
    RecebForaFaixa.Clear

    RecebForaFaixa.AddItem MENSAGEM_NAO_AVISA_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = NAO_AVISA_E_ACEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_REJEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = ERRO_E_REJEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = AVISA_E_ACEITA_RECEBIMENTO

    Exit Function

End Function

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then Error 56380

    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = Err

    Select Case Err

        Case 56380

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164598)

    End Select

    Exit Function


End Function

Private Function Carrega_TipoBloqueio() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoBloqueio
    'Lê o Código e o NOme de Todas os Tipos de Bloqueio do BD
    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueioPC", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then Error 49491

    'Carrega a combo de Tipo de Bloqueio
    For Each objCodigoNome In colCodigoNome

        'Verifica se o Tipo de Bloqueio é diferente de "Alçada"
        If objCodigoNome.iCodigo <> BLOQUEIO_ALCADA Then

            TipoBloqueio.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodigoNome.iCodigo

        End If

    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = Err

    Select Case Err

        Case 49491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164599)

    End Select

    Exit Function

End Function
Private Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then Error 49462

    'Carrega a combo de Transportadoras
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = Err

    Select Case Err

        Case 49462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164600)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then Error 49463
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Adiciona novo ítem na List da Combo CondicaoPagamento
'        CondPagto.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        CondPagto.ItemData(CondPagto.NewIndex) = objCod_DescReduzida.iCodigo
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
'        Case 49463
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164601)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 49460

    CentroCusto.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 49460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164602)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
    Inicializa_Grid_Itens = CF2(Me, "PC_Inicializa_Grid_Itens", objGridInt)
End Function
 
Public Function PC_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens

Dim iIncremento As Integer

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    
    'Alteracao Daniel
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGridInt.colColuna.Add ("Preço (R$)")
        objGridInt.colColuna.Add ("Total (R$)")
    End If
    
    If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
        objGridInt.colColuna.Add ("Delivery Date")
        objGridInt.colColuna.Add ("Tempo de Trânsito")
    End If
    
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Valor IPI ")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("% a Mais Receb")
    objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Desc. Completa") 'leo
    
    ' campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    
    'Alteracao Daniel
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal Then
        objGridInt.colCampo.Add (PrecoUnitarioMoedaReal.Name)
        objGridInt.colCampo.Add (TotalMoedaReal.Name)
    Else
        PrecoUnitarioMoedaReal.left = POSICAO_FORA_TELA
        PrecoUnitarioMoedaReal.TabStop = False
    
        TotalMoedaReal.left = POSICAO_FORA_TELA
        TotalMoedaReal.TabStop = False
    End If
    
    If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
        objGridInt.colCampo.Add (DeliveryDate.Name)
        objGridInt.colCampo.Add (TempoTransito.Name)
    Else
        DeliveryDate.left = POSICAO_FORA_TELA
        DeliveryDate.TabStop = False
    
        TempoTransito.left = POSICAO_FORA_TELA
        TempoTransito.TabStop = False
    End If
    
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (ValorIPIItem.Name)
    objGridInt.colCampo.Add (AliquotaICMS.Name)
    objGridInt.colCampo.Add (PercentMaisReceb.Name)
    objGridInt.colCampo.Add (RecebForaFaixa.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    objGridInt.colCampo.Add (DescCompleta.Name) 'leo
    
    'indica onde estao situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_PrecoUnitario_Col = 5
    iGrid_PercentDesc_Col = 6
    iGrid_Desconto_Col = 7
    iGrid_PrecoTotal_Col = 8
    
    iIncremento = 0
    
    If bExibirColReal Then
        iGrid_PrecoUnitarioMoedaReal_Col = 9 + iIncremento
        iGrid_TotalMoedaReal_Col = 10 + iIncremento
        iIncremento = iIncremento + 2
    End If
    
    If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
        iGrid_DeliveryDate_Col = 9 + iIncremento
        iGrid_TempoTransito_Col = 10 + iIncremento
        iIncremento = iIncremento + 2
    End If
    
    iGrid_DataLimite_Col = 9 + iIncremento
    iGrid_AliquotaIPI_Col = 10 + iIncremento
    iGrid_ValorIPIItem_Col = 11 + iIncremento
    iGrid_AliquotaICMS_Col = 12 + iIncremento
    iGrid_PercentMaisReceb_Col = 13 + iIncremento
    iGrid_RecebForaFaixa_Col = 14 + iIncremento
    iGrid_Observacao_Col = 15 + iIncremento
    iGrid_DescCompleta_Col = 16 + iIncremento
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'linhas visiveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 16
    Else
        objGridInt.iLinhasVisiveis = 5
    End If

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    PC_Inicializa_Grid_Itens = SUCESSO

End Function

 Private Function Inicializa_Grid_Distribuicao(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Distribuicao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Centro de Custo")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Unidade Medida")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Conta Contábil")

    ' campos de edição do grid
    objGridInt.colCampo.Add (Prod.Name)
    objGridInt.colCampo.Add (DescProd.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (UnidMed.Name)
    objGridInt.colCampo.Add (Quant.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Prod_Col = 1
    iGrid_DescProd_Col = 2
    iGrid_CentroCusto_Col = 3
    iGrid_Almoxarifado_Col = 4
    iGrid_UnidMed_Col = 5
    iGrid_Quant_Col = 6
    iGrid_ContaContabil_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDistribuicao

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_DISTRIBUICAO + 1

    'linhas visiveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 18
    Else
        objGridInt.iLinhasVisiveis = 6
    End If

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    If m_objUserControl.Height <= 9000 Then
        GridDistribuicao.Width = 8295
    End If

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Distribuicao = SUCESSO

    Exit Function

End Function
Private Function Inicializa_Grid_Bloqueios(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Distribuicao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Tipo Bloqueio")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Responsável")
    objGridInt.colColuna.Add ("Data Liberação")
    objGridInt.colColuna.Add ("Resp. Liberação")
    
    ' campos de edição do grid
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (CodUsuario.Name)
    objGridInt.colCampo.Add (ResponsavelBL.Name)
    objGridInt.colCampo.Add (DataLiberacao.Name)
    objGridInt.colCampo.Add (ResponsavelLib.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_CodUsuario_Col = 3
    iGrid_ResponsavelBL_Col = 4
    iGrid_DataLiberacao_Col = 5
    iGrid_ResponsavelLib_Col = 6
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridBloqueios

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1

    GridBloqueios.ColWidth(0) = 300

    'linhas visiveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 18
    Else
        objGridInt.iLinhasVisiveis = 6
    End If

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Bloqueios = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSiglas As New Collection
Dim objClasseUM As New ClassClasseUM
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String, sProd As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 49435

    'Pesquisa controle da coluna em questão
    Select Case objControl.Name
        'Produto
        Case Produto.Name
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

'            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then Error 49435

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then Error 49436
                If lErro = 28030 Then Error 49437 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then Error 49438
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                If Len(Trim(sUM)) > 0 Then

                    For iIndice = 0 To UnidadeMed.ListCount - 1

                        If sUM = UnidadeMed.List(iIndice) Then
                            UnidadeMed.ListIndex = iIndice
                        End If

                    Next
                End If
            End If

        'Nas demais
        'Case Quantidade.Name, PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, DataLimite.Name, AliquotaIPI.Name, ValorIPIItem.Name, AliquotaICMS.Name, PercentMaisReceb.Name, RecebForaFaixa.Name, Observacao.Name, DescProduto.Name
        Case Quantidade.Name, PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, AliquotaIPI.Name, ValorIPIItem.Name, AliquotaICMS.Name, PercentMaisReceb.Name, RecebForaFaixa.Name, Observacao.Name, DescProduto.Name

'            'Verifica se o produto está preenchido
'            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then Error 49432

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case DeliveryDate.Name, TempoTransito.Name

            If iProdutoPreenchido = PRODUTO_VAZIO Or gobjCOM.iPCExibeDeliveryDate = DESMARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case DataLimite.Name
            If iProdutoPreenchido = PRODUTO_VAZIO Or gobjCOM.iPCExibeDeliveryDate = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Prod.Name

            If Len(Trim(GridDistribuicao.TextMatrix(iLinha, iGrid_Prod_Col))) <> 0 Then
                sProd = GridDistribuicao.TextMatrix(iLinha, iGrid_Prod_Col)
            Else
                sProd = Prod.Text
            End If
            
            Prod.Clear

            'carrega os produtos que estao no grid de itens
            For iIndice = 1 To objGridItens.iLinhasExistentes
                'se o elemento está marcado para ser liberado
                If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
                    Prod.AddItem (GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
                End If
            Next

            lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 49433

            For iIndice = 0 To Prod.ListCount - 1
                If sProd = Prod.List(iIndice) Then
                    Prod.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            If iProdutoPreenchido = PRODUTO_VAZIO Or Len(Trim(GridDistribuicao.TextMatrix(iLinha, iGrid_Prod_Col))) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case CentroCusto.Name, Almoxarifado.Name, Quant.Name, ContaContabil.Name

            lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 49434

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            ElseIf iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                objControl.Enabled = True
            End If

        Case TipoBloqueio.Name, ResponsavelBL.Name

            If iLinha <= iGridBloqueiosLinhasBloqueadas Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True

            End If
            
        Case Else
            objControl.Enabled = False

            '###############################################################
            'Inserido por Wagner 17/07/2006
            lErro = CF2(Me, "PC_Rotina_Grid_Enable", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124178
            '###############################################################


    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case Err

        Case 49432, 49433, 49434, 49435, 49436, 49438, 49439, 56396, 56397, 56399, 56400

        Case 49437, 56398
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164603)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'libera as variaveis globais
    Set objEventoAlmoxarifados = Nothing
    Set objEventoBotaoCcl = Nothing
    Set objEventoBotaoPedidosAvulsos = Nothing
    Set objEventoCliente = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoFornec = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoObservacao = Nothing
    Set objEventoProdutos = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoNotaPC = Nothing
    Set objEventoEmbalagens = Nothing 'leo
    Set gobjInfoAdic = Nothing

    Set objGridItens = Nothing
    Set objGridBloqueio = Nothing
    Set objGridDistribuicao = Nothing

    Set gcolItemPedido = Nothing
    Set colAliqExternas = Nothing
    Set gobjTelaProjetoInfo = Nothing
    Set gobjPC = Nothing
        
    '??? por leo
    Set objGridNotas = Nothing
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

End Sub

Public Sub Fornec_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Fornec_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Fornec_Validate

    If iFornecAlterado <> 0 Then

        'Verifica se Fornec esta preenchido
        If Len(Trim(Fornec.Text)) <> 0 Then

            'Le Fornec
            lErro = TP_Fornecedor_Le(Fornec, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 49496

            'Le as filiais do Fornec
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO And lErro <> 6698 Then Error 49497

            'Preenche a combo de filiais
            Call CF("Filial_Preenche", FilialFornec, colCodigoNome)

            'Seleciona a filial na combo de filiais
            Call CF("Filial_Seleciona", FilialFornec, iCodFilial)

        Else
            'Limpa a combobox
            FilialFornec.Clear

        End If

    End If

    iFornecAlterado = 0

    Exit Sub

Erro_Fornec_Validate:

    Cancel = True

    Select Case Err

        Case 49496, 49497

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164604)

    End Select

    Exit Sub

End Sub
Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO
    
    Call Fornecedor_Preenche

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iCondPagto As Integer
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub

    'Verifica se fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 56061

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then Error 56062

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

        If objFornecedor.iCondicaoPagto > 0 Then

            iCondPagto = objFornecedor.iCondicaoPagto

        Else

            If objFornecedor.iTipo > 0 Then

                objTipoFornecedor.iCodigo = objFornecedor.iTipo

                lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
                'OK não encontrar
                If lErro <> SUCESSO And lErro <> 12765 Then Error 57247
                'Se nao encontrou => errro
                If lErro = 12765 Then Error 57248

                If objTipoFornecedor.iCondicaoPagto <> 0 Then

                    iCondPagto = objFornecedor.iCondicaoPagto

                End If

            End If

        End If
        If iCondPagto <> 0 Then

            CondPagto.Text = iCondPagto
            CondPagto_Validate (Cancel)

        End If
        
    Else

        FilialFornec.Clear

    End If

    iFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case Err

        Case 56061, 56062, 57247

        Case 57248
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", Err, objTipoFornecedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164605)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objFornecedor.sNomeReduzido = FornecedorLabel.Caption

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Public Sub FornLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    objFornecedor.sNomeReduzido = Fornec.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornec)

End Sub

Public Sub GridBloqueios_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If GridBloqueios.Row <= iGridBloqueiosLinhasBloqueadas Then Exit Sub
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)

    Exit Sub

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAtual As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim dValorIPI As Double
Dim sProduto As String
Dim iItem As Long

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iLinhaAtual = GridItens.Row
    sProduto = GridItens.TextMatrix(iLinhaAtual, iGrid_Produto_Col)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        gcolItemPedido.Remove (iLinhaAtual)
        gobjPC.colItens.Remove (iLinhaAtual)
        Call gobjTelaProjetoInfo.Remove_Linha(iLinhaAtual)

        'exclui distribuicoes e
        For iIndice = 1 To objGridDistribuicao.iLinhasExistentes

            If (sProduto = GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col)) Then

                lErro = Grid_Exclui_Linha(objGridDistribuicao, iIndice)
                If lErro <> SUCESSO Then gError 56349
                iIndice = iIndice - 1
                
            End If

            For iItem = 1 To objGridItens.iLinhasExistentes

                'recalcula ValorIPI
                If Len(Trim(GridItens.TextMatrix(iItem, iGrid_ValorIPIItem_Col))) > 0 Then
                    If StrParaDbl(GridItens.TextMatrix(iItem, iGrid_ValorIPIItem_Col)) > 0 Then dValorIPI = dValorIPI + StrParaDbl(GridItens.TextMatrix(iItem, iGrid_ValorIPIItem_Col))
                End If
            Next
        Next
        
        '#############################################################
        'Inserido por Wagner 23/03/2006
        lErro = CF2(Me, "PC_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177160
        '#############################################################
        
        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 56155

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 56155, 56349, 177160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164606)

    End Select

    Exit Sub


End Sub

Private Sub objEventoAlmoxarifados_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto

    Set objEstoqueProduto = obj1

    If GridDistribuicao.Row <> 0 Then

        GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    End If

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoBotaoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sCclMascarado As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_objEventoBotaoCcl_evSelecao

    Set objCcl = obj1

    'Mascara o Ccl
    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 56054
    
    CentroCusto.promptinclude = False
    CentroCusto.Text = sCclMascarado
    CentroCusto.promptinclude = True
    
    'Coloca ccl mascarado no GridDistribuicao
    GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_CentroCusto_Col) = CentroCusto.Text

    Me.Show

    Exit Sub

Erro_objEventoBotaoCcl_evSelecao:

    Select Case Err

        Case 56054

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164607)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoPedidosAvulsos_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoBotaoPedidosAvulsos_evSelecao

    Set objPedidoCompra = obj1

    lErro = CF("PedidoCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56118 Then gError 178001
    
    If lErro = 56118 Then gError 178002

    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then gError 49494

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoBotaoPedidosAvulsos_evSelecao:

    Select Case gErr

        Case 49494, 178001


        Case 178002
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164608)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPedidoCompra = obj1

    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            'Le o Pedido de Compra
            lErro = CF("PedidoCompras_Le", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56118 Then Error 53242
            If lErro = 56118 Then Error 53243

            lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
            If lErro <> SUCESSO Then Error 53244

        End If

    End If
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 53242, 53244

        Case 53243
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164609)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim objCondicaoPagto As New ClassCondicaoPagto

    Set objCondicaoPagto = obj1

    'Coloca Condicao Pagamento na tela
    CondPagto.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondPagto_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sContaMascarado As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    'Mascara a conta contabil
    lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarado)
    If lErro <> SUCESSO Then Error 56056

    ContaContabil.promptinclude = False
    ContaContabil.Text = sContaMascarado
    ContaContabil.promptinclude = True

    'Coloca no Grid a Conta mascarada
    GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_ContaContabil_Col) = ContaContabil.Text

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 56056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164610)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornec_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    Fornec.Text = objFornecedor.sNomeReduzido

    'Dispara Validate de Fornec
    Fornec_Validate (bSGECancelDummy)
    
    Me.Show
    
    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Dispara Validate de Fornecedor
    Fornecedor_Validate (bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoObservacao_evSelecao(obj1 As Object)

Dim objObservacao As ClassObservacao

    Set objObservacao = obj1

    'Coloca observacao na tela
    ObservacaoPC.Text = objObservacao.sObservacao

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoProdutos_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    If GridItens.Row <> 0 Then

        'Verifica se o Produto nao está preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

            Set objProduto = obj1
            
            'Torna o produto enxuto
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 56050

            'Coloca produto enxuto na tela
            Produto.promptinclude = False
            Produto.Text = sProdutoEnxuto
            Produto.promptinclude = True

            '##########################################################
            'Inserido por Wagner 14/07/2006
            lErro = CF2(Me, "PC_Produto_Saida_Celula", objProduto)
            If lErro <> SUCESSO Then gError 181215

'            'Verifica se já existe o Produto em outra linha do Grid
'            For iIndice = 1 To objGridItens.iLinhasExistentes
'                If iIndice <> GridItens.Row Then
'                    If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 74940
'                End If
'
'            Next
            '##########################################################

            If Not (Me.ActiveControl Is Produto) Then

                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
                        
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 56051


            End If
            
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 56050, 56051, 181215
                
'        Case 74940
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE_PEDCOMPRA", gErr, Produto.Text, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164611)

    End Select

    Exit Sub

End Sub


Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As New ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportadora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Function Saida_Celula_PercentMaisReceb(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentMaisReceb As Double

On Error GoTo Erro_Saida_Celula_PercentMaisReceb

    Set objGridInt.objControle = PercentMaisReceb

    'Verifica se PercentMaisReceb foi preenchido
    If Len(Trim(PercentMaisReceb.Text)) > 0 Then

        'Critica porcentagem informada
        lErro = Porcentagem_Critica(PercentMaisReceb.Text)
        If lErro <> SUCESSO Then Error 49478

        dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text)

        'Coloca no formato fixed da tela
        PercentMaisReceb.Text = Format(dPercentMaisReceb, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49479

    Saida_Celula_PercentMaisReceb = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentMaisReceb:

    Saida_Celula_PercentMaisReceb = Err

    Select Case Err

        Case 49478

        Case 49479
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164612)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AliquotaICMS(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dAliquotaICMS As Double

On Error GoTo Erro_Saida_Celula_AliquotaICMS

    Set objGridInt.objControle = AliquotaICMS

    'Verifica se AliquotaICMS foi preenchida
    If Len(Trim(AliquotaICMS.Text)) > 0 Then

        'Critica a aliquota informada
        lErro = Porcentagem_Critica(AliquotaICMS.Text)
        If lErro <> SUCESSO Then Error 49475

        dAliquotaICMS = StrParaDbl(AliquotaICMS.Text)

        'Coloca aliquota no formato fixed da tela
        AliquotaICMS.Text = Format(dAliquotaICMS, "Fixed")

        If AliquotaICMS.Text = 100# Then Error 49476

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49477

    Saida_Celula_AliquotaICMS = SUCESSO

    Exit Function

Erro_Saida_Celula_AliquotaICMS:

    Saida_Celula_AliquotaICMS = Err

    Select Case Err

        Case 49475
            'Erro tratado na rotina chamada
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49476
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_IGUAL_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49477
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164613)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sUnidadeMed As String
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_UnidadeMed

    sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

    Set objGridInt.objControle = UnidadeMed

    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 56006
    
    If sUnidadeMed <> UnidadeMed.Text And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) Then
    
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 86124
        
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 86125
        If lErro <> SUCESSO Then gError 86126
                
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
            lErro = Distribuicao_Atualiza(GridItens.Row, objProduto)
            If lErro <> SUCESSO Then gError 86127
        End If

        '#############################################################
        'Inserido por Wagner 23/03/2006
        lErro = CF2(Me, "PC_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177161
        '#############################################################

    End If

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 56006, 86124, 86125, 86127, 177161
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86126
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164614)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dValorIPI As Double
Dim bPrecoUnitarioIgual As Boolean
Dim dPrecoUnitario As Double
Dim dPrecoTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitario

    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then Error 56097

        dPrecoUnitario = StrParaDbl(PrecoUnitario.Text)

        'Coloca PrecoUnitario no formato Standard da tela
        PrecoUnitario.Text = Format(dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) 'Alterado por Wagner

    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56098

    If bPrecoUnitarioIgual = False Then

        'Verifica se preco unitario e quantidade foram preenchidos
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 Then

'            'Verifica se percentual de desconto foi preenchido
'            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col))) > 0 Then
'
'                dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'                Call CF2(Me, "PC_PrecoTotal_Calcula", GridItens.Row, dQuantidade) 'Inserido por Wagner 17/07/2006
'
'                dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
'
'                'calcula o desconto
'                dDesconto = (dPrecoUnitario * dQuantidade) * (PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercentDesc_Col)))
'
'                If dDesconto > 0 Then
'                    GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
'                Else
'                    GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
'                End If

            Call PrecoTotal_Calcula(GridItens.Row) 'Inserido por Wagner 17/07/2006

'            End If

'            'Calcula o preco total
'            dPrecoTotal = StrParaDbl(PrecoUnitario.Text) * dQuantidade
'            dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
'            dPrecoTotal = dPrecoTotal - dDesconto
'
'            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
'
            If Len(Trim(Taxa.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Text))

'            'Verifica se AliquotaIPI foi preenchida
'            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col))) <> 0 Then
'
'                'Calcula valorIPI
'                dValorIPI = (dPrecoTotal) * (PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col)))
'                If dValorIPI > 0 Then
'                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "Standard")
'                Else
'                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""
'                End If
'
'            End If

        'Se preco unitario nao esta preenchido
        Else

            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = ""
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""


        End If

'        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner

'        'Calcula o valorIPI da tela
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'
'            dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
'            dValorTotal = dValorTotal + dValorIPI
'            ValorIPI.Text = Format(dValorTotal, "standard")
'        Next

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56121

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = Err

    Select Case Err

        Case 56097, 56098, 56121
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164615)

    End Select

    Exit Function

End Function

Function ValorTotal_Calcula() As Long

Dim dPrecoTotal As Double
Dim dValorTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_ValorTotal_Calcula

   For iIndice = 1 To objGridItens.iLinhasExistentes

        'Calcula a soma dos valores de produtos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If

        'Calcula Preco Total das linhas do GridItens
        dPrecoTotal = dPrecoTotal + (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)))

    Next

    'Coloca na tela o valor dos produtos
    ValorProdutos.Caption = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
    dValorTotal = (dPrecoTotal + StrParaDbl(ValorFrete.Text) + StrParaDbl(ValorSeguro.Text) + StrParaDbl(OutrasDespesas.Text) + StrParaDbl(ValorIPI.Text)) - StrParaDbl(ValorDesconto.Text)

    'Coloca na tela o valor total
    ValorTotal.Caption = Format(dValorTotal, PrecoTotal.Format) 'Alterado por Wagner


    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164616)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AliquotaIPI(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dAliquotaIPI As Double
Dim bAliquotaIPIIgual As Boolean
Dim dPrecoTotal As Double
Dim dValorTotal As Double
Dim dValorIPI As Double
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_AliquotaIPI

    Set objGridInt.objControle = AliquotaIPI

    If Len(Trim(AliquotaIPI.Text)) > 0 Then

        lErro = Porcentagem_Critica(AliquotaIPI.Text)
        If lErro <> SUCESSO Then Error 56099

        AliquotaIPI.Text = Format(AliquotaIPI.Text, "Fixed")

        'Verifica se AliquotaIPI foi alterada
        If PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col)) = StrParaDbl(AliquotaIPI.Text) Then bAliquotaIPIIgual = True

        If bAliquotaIPIIgual = False And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))) <> 0 Then

            dAliquotaIPI = StrParaDbl(AliquotaIPI.Text)
            If dAliquotaIPI = 100 Then Error 56100

            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

            dValorIPI = dPrecoTotal * dAliquotaIPI / 100
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "Standard")
        End If

    Else  'se aliquota nao estiver preenchida

        GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56101

    'Calcula valorIPI para o pedido
    For iIndice = 1 To objGridItens.iLinhasExistentes

        dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
        dValorTotal = dValorTotal + dValorIPI

    Next

    ValorIPI = Format(dValorTotal, "standard")

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56122

    Saida_Celula_AliquotaIPI = SUCESSO

    Exit Function


Erro_Saida_Celula_AliquotaIPI:

    Saida_Celula_AliquotaIPI = Err

    Select Case Err

        Case 56099, 56101, 56122
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56100
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_IGUAL_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164617)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorIPI(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorIPI As Double
Dim dPrecoTotal As Double
Dim dAliquotaIPI As Double
Dim bValorIPIIgual As Boolean
Dim iIndice As Integer
Dim dValorTotal As Double

On Error GoTo Erro_Saida_Celula_ValorIPI

    Set objGridInt.objControle = ValorIPIItem

    bValorIPIIgual = False

    'Verifica se ValorIPI foi preenchido
    If Len(Trim(ValorIPIItem.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_Positivo_Critica(ValorIPIItem.Text)
        If lErro <> SUCESSO Then Error 56102

        'Coloca no formato standard da tela
        ValorIPIItem.Text = Format(ValorIPIItem.Text, "Standard")

        dValorIPI = StrParaDbl(ValorIPIItem)

    End If

    'Comparação com Valor IPI anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorIPIItem_Col)) = StrParaDbl(ValorIPIItem.Text) Then bValorIPIIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56104

    If bValorIPIIgual = False Then

        dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        'Verifica se Preco Total foi preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))) <> 0 And dPrecoTotal > 0 Then

            If dValorIPI > dPrecoTotal Then Error 56103

            dAliquotaIPI = dValorIPI / dPrecoTotal

            'coloca no formato percent
            GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col) = Format(dAliquotaIPI, "Percent")

        End If

        'Calcula o valor IPI para o pedido
         For iIndice = 1 To objGridItens.iLinhasExistentes

            dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
            dValorTotal = dValorTotal + dValorIPI

        Next

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then Error 56123

    End If

    Saida_Celula_ValorIPI = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorIPI:

    Saida_Celula_ValorIPI = Err

    Select Case Err

        Case 56102, 56104, 56123
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56103
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORIPI_MAIOR_OU_IGUAL_PRECO_TOTAL", Err, dValorIPI, dPrecoTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164618)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Prod(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim bProdutoPresente As Boolean
Dim iIndice As Integer
Dim iLinha As Integer
Dim sProduto As String
Dim iIndex As Integer
Dim iIndexAux As Integer
Dim iCont As Integer

On Error GoTo Erro_Saida_Celula_Prod

    Set objGridInt.objControle = Prod
    
    iIndex = Prod.ListIndex
    
    iIndexAux = 0
    For iIndice = 0 To iIndex - 1
        If Prod.Text = Prod.List(iIndice) Then
            iIndexAux = iIndexAux + 1
        End If
    Next

    bProdutoPresente = False

    If Len(Trim(Prod.Text)) > 0 Then

        lErro = CF("Produto_Critica_Compra", Prod.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 63233 Then gError 56105

        'Se nao encontrou
       If lErro = 63233 Then gError 56208

        'Verifica se já está em outra linha do Grid
        iCont = -1
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Prod.Text Then
                iCont = iCont + 1
                bProdutoPresente = True
                If iCont = iIndexAux Then
                    Exit For
                End If
            End If
        Next

        'Guarda a linha do produto no grid itens
        iLinha = iIndice

        If bProdutoPresente = False Then gError 56210

        'Coloca descricao do produto no Grid Distribuicao
        GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_DescProd_Col) = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) 'objProduto.sDescricao
        GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_UnidMed_Col) = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

        bProdutoPresente = False

        'Verifica se está presente no Grid de Distribuicao
        For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
            If iIndice <> GridDistribuicao.Row Then
                If Trim(GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col)) = Prod.Text Then
                    bProdutoPresente = True
                End If
            End If
        Next

        If bProdutoPresente = False And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) <> 0 Then

            GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Quant_Col) = GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)
            GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_UnidMed_Col) = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
        End If


        'Acrescenta uma linha no Grid se for o caso
        If GridDistribuicao.Row - GridDistribuicao.FixedRows = objGridDistribuicao.iLinhasExistentes Then
            objGridDistribuicao.iLinhasExistentes = objGridDistribuicao.iLinhasExistentes + 1
        End If

    End If

    'GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col) = Prod.Text
    
    sProduto = Prod.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49470

    GridDistribuicao.TextMatrix(GridDistribuicao.Row, iGrid_Prod_Col) = sProduto

    Saida_Celula_Prod = SUCESSO

    Exit Function

Erro_Saida_Celula_Prod:

    Saida_Celula_Prod = gErr

    Select Case gErr

        Case 49470, 56105
            Call Grid_Trata_Erro_Saida_Celula(objGridDistribuicao)

        Case 56208
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Prod.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridDistribuicao)

        Case 56210
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DESVINCULADO_ITEM", gErr, Prod.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridDistribuicao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164619)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quant(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quant

    Set objGridInt.objControle = Quant

    'Se quantidade estiver preenchida
    If Len(Trim(Quant.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Quant.Text)
        If lErro <> SUCESSO Then Error 49472

        dQuantidade = CDbl(Quant.Text)

        'Coloca o valor Formatado na tela
        Quant.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49473

    Saida_Celula_Quant = SUCESSO

    Exit Function

Erro_Saida_Celula_Quant:

    Saida_Celula_Quant = Err

    Select Case Err

        Case 49472, 49473
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164620)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescProduto

    Set objGridInt.objControle = DescProduto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56374

    Saida_Celula_DescProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescProduto:

    Saida_Celula_DescProduto = Err

    Select Case Err

        Case 56374
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164621)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_Responsavel(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Responsavel

    Set objGridInt.objControle = ResponsavelBL

    'Verifica se ResponsavelBL foi informado
    If Len(Trim(ResponsavelBL.Text)) > 0 Then

        'Verifica se o TipodeBloqueio está preenchido
        If Len(Trim(GridBloqueios.TextMatrix(GridBloqueios.Row, iGrid_TipoBloqueio_Col))) = 0 Then gError 68338

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49485

    Saida_Celula_Responsavel = SUCESSO

    Exit Function

Erro_Saida_Celula_Responsavel:

    Saida_Celula_Responsavel = gErr

    Select Case gErr

        Case 49485
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 68338
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO", gErr, GridBloqueios.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164622)

    End Select

    Exit Function

End Function

'''Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long
'''
'''Dim lErro As Long
'''Dim sContaSimples As String
'''Dim objPlanoConta As New ClassPlanoConta
'''Dim sContaEnxuta As String
'''Dim sContaFormatada As String
'''Dim vbMsgRes As VbMsgBoxResult
'''Dim sContaMascarada As String
'''
'''On Error GoTo Erro_Saida_Celula_ContaContabil
'''
'''    Set objGridInt.objControle = ContaContabil
'''
'''    'Verifica se a conta contabil esta preenchida
'''    If Len(Trim(ContaContabil.ClipText)) > 0 Then
'''
'''        'Critica a conta informada
'''        lErro = CF("ContaSimples_Critica", sContaSimples, ContaContabil.ClipText, objPlanoConta)
'''        If lErro <> SUCESSO And lErro <> 44037 Then Error 56186
'''
'''        'Se for uma conta simples
'''        If lErro = SUCESSO Then
'''
'''            lErro = Mascara_RetornaContaEnxuta(sContaSimples, sContaEnxuta)
'''            If lErro <> SUCESSO Then Error 56187
'''
'''            'Coloca conta enxuta na tela
'''            ContaContabil.Text = sContaEnxuta
'''
'''        End If
'''
'''        'Se a conta nao e uma conta simples
'''        If lErro = 44037 Then
'''
'''            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
'''            If lErro <> SUCESSO And lErro <> 5700 Then Error 56188
'''
'''            'Nao encontrou a conta
'''            If lErro = 5700 Then Error 56189
'''
'''        End If
'''
'''        lErro = Mascara_MascararConta(sContaFormatada, sContaMascarada)
'''        If lErro <> SUCESSO Then Error 63622
'''
'''        'Coloca a descricao da conta na tela
'''        ContaContabil.PromptInclude = False
'''        ContaContabil.Text = sContaMascarada
'''        ContaContabil.PromptInclude = True
'''
'''    End If
'''
'''    lErro = Grid_Abandona_Celula(objGridInt)
'''    If lErro <> SUCESSO Then Error 56190
'''
'''    Saida_Celula_ContaContabil = SUCESSO
'''
'''    Exit Function
'''
'''Erro_Saida_Celula_ContaContabil:
'''
'''    Saida_Celula_ContaContabil = Err
'''
'''    Select Case Err
'''
'''        Case 56186, 56187, 56188, 56190, 63622
'''            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'''
'''        Case 56189
'''            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)
'''
'''            If vbMsgRes = vbYes Then
'''                objPlanoConta.sConta = sContaFormatada
'''
'''                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
'''                Call Chama_Tela("PlanoConta", objPlanoConta)
'''            Else
'''                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'''            End If
'''
'''        Case Else
'''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164623)
'''
'''    End Select
'''
'''    Exit Function
'''
'''End Function

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid)
'Faz a crítica da célula Tipo Bloqueio que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoBloqueio

    Set objGridInt.objControle = TipoBloqueio

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoBloqueio.Text)) > 0 Then

        'Tenta selecioná-lo na combo
        lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
        If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then Error 56162

        'Não foi encontrado
        If lErro = 25085 Then Error 56163
        If lErro = 25086 Then Error 56164

        'Verifica se já existe um bloqueio, não liberado, do mesmo tipo no grid.
        For iIndice = 1 To objGridInt.iLinhasExistentes

            If iIndice <> GridBloqueios.Row Then
                iCodigo = Codigo_Extrai(GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))

                If iCodigo = TipoBloqueio.ItemData(TipoBloqueio.ListIndex) Then
                    If Len(Trim(GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) = 0 Then Error 63613
                End If
            End If

        Next

        'Acrescenta uma linha no Grid se for o caso
        If GridBloqueios.Row - GridBloqueios.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

            'Coloca data e usuário
            GridBloqueios.TextMatrix(GridBloqueios.Row, iGrid_DataBloqueio_Col) = Format(gdtDataAtual, "dd/mm/yyyy")
            GridBloqueios.TextMatrix(GridBloqueios.Row, iGrid_CodUsuario_Col) = Format(gsUsuario)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56165

    Saida_Celula_TipoBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoBloqueio:

    Saida_Celula_TipoBloqueio = Err

    Select Case Err

        Case 56162, 56165
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56163
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56164
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", Err, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63613
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPV_REPETIDO", Err, Codigo_Extrai(TipoBloqueio.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164624)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_RecebForaFaixa(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RecebForaFaixa

    Set objGridInt.objControle = RecebForaFaixa

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49482

    Saida_Celula_RecebForaFaixa = SUCESSO

    Exit Function

Erro_Saida_Celula_RecebForaFaixa:

    Saida_Celula_RecebForaFaixa = Err

    Select Case Err

        Case 49482
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164625)

    End Select

    Exit Function


End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49483

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = Err

    Select Case Err

        Case 49483
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164626)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iLinha As Integer
Dim sProdutoFormatado As String, dPrecoUnitario As Double
Dim iIndice As Integer, objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se o Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 63256

        objProduto.sCodigo = sProdutoFormatado

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Verifica se o Produto existe e pode ser Comprado
            lErro = CF("Produto_Critica_Compra", Produto.Text, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25605 Then gError 49466

            'Se o produto não existir ==> erro
            If lErro = 25605 Then gError 49467

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
            If lErro <> SUCESSO Then gError 63849

            'Coloca o produto no grid
            Produto.promptinclude = False
            Produto.Text = sProduto
            Produto.promptinclude = True

            '###########################################################
            'Inserido por Wagner 14/07/2006
            lErro = CF2(Me, "PC_Produto_Saida_Celula", objProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'            'Verifica se já existe o Produto em outra linha do Grid
'            For iIndice = 1 To objGridItens.iLinhasExistentes
'                If iIndice <> GridItens.Row Then
'                    If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then Error 63612
'                End If
'
'            Next
            '###########################################################
        
        End If

        'Coloca as demais características do produto na tela
        lErro = ProdutoLinha_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 49469

        'Preço Unitário
        If Len(Trim(TabelaPreco.Text)) > 0 Then
    
            'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
        
            objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
            objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
            objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
            If lErro = 28014 Then
                objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
    
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
            End If
                
            If lErro <> 28014 Then
    
                lErro = PrecoUnitario_Calcula(GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                If dPrecoUnitario > 0 Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, PrecoUnitario.Format)
                Else
                    GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = ""
                End If
    
            End If
            
            'Remove o produto porque senão a Rotina_Grid_Enabled vai desabilitar o campo antes
            'do Grid_Abandona_Celula
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
    
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49470

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 49465, 49466, 49469, 49470, 63849, ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49467
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 49468
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 63612
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE_PEDCOMPRA", gErr, Produto.Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164627)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Tela_Preenche

    'Carrega objPedidoCompra com os dados passados em colCampoValor
    objPedidoCompra.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
    objPedidoCompra.dtData = colCampoValor.Item("Data").vValor
    objPedidoCompra.dtDataAlteracao = colCampoValor.Item("DataAlteracao").vValor
    objPedidoCompra.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
    objPedidoCompra.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
    objPedidoCompra.dValorFrete = colCampoValor.Item("ValorFrete").vValor
    objPedidoCompra.dValorIPI = colCampoValor.Item("ValorIPI").vValor
    objPedidoCompra.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
    objPedidoCompra.dValorTotal = colCampoValor.Item("ValorTotal").vValor
    objPedidoCompra.iComprador = colCampoValor.Item("Comprador").vValor
    objPedidoCompra.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
    objPedidoCompra.iFilial = colCampoValor.Item("Filial").vValor
    objPedidoCompra.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
    objPedidoCompra.iProxSeqBloqueio = colCampoValor.Item("ProxSeqBloqueio").vValor
    objPedidoCompra.iTipoBaixa = colCampoValor.Item("TipoBaixa").vValor
    objPedidoCompra.iTipoDestino = colCampoValor.Item("TipoDestino").vValor
    objPedidoCompra.lCodigo = colCampoValor.Item("Codigo").vValor
    objPedidoCompra.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
    objPedidoCompra.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    objPedidoCompra.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objPedidoCompra.iTransportadora = colCampoValor.Item("Transportadora").vValor
    objPedidoCompra.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objPedidoCompra.sAlcada = colCampoValor.Item("Alcada").vValor
    objPedidoCompra.sContato = colCampoValor.Item("Contato").vValor
    objPedidoCompra.sMotivoBaixa = colCampoValor.Item("MotivoBaixa").vValor
    objPedidoCompra.lObservacao = colCampoValor.Item("Observacao").vValor
    objPedidoCompra.sTipoFrete = colCampoValor.Item("TipoFrete").vValor
    objPedidoCompra.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
'leo
    objPedidoCompra.iEmbalagem = colCampoValor.Item("Embalagem").vValor
    objPedidoCompra.dTaxa = colCampoValor.Item("Taxa").vValor
    objPedidoCompra.iMoeda = colCampoValor.Item("Moeda").vValor
    objPedidoCompra.sObsEmbalagem = colCampoValor.Item("ObsEmbalagem").vValor


    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then Error 49487

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 49487

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164628)

    End Select

    Exit Sub

End Sub
Private Function Valida_Grid_Distribuicao() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dQuantidade As Double
Dim bTeste As Boolean

On Error GoTo Erro_Valida_Grid_Distribuicao

    For iIndice = 1 To objGridDistribuicao.iLinhasExistentes

        'Verifica se Produto foi preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col))) = 0 Then gError 56191

        'Verifica se Almoxarifado foi preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 56192

    Next

    For iIndice2 = 1 To objGridItens.iLinhasExistentes
        
        dQuantidade = 0
               
        For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
            'Acumula a quantidade distribuida
            
            '##################################
            'Inserido por Wagner 27/07/2006
            bTeste = True
            Call CF2(Me, "PC_Produto_Igual_Customizado", iIndice2, iIndice, bTeste)
            '##################################
                                    
            If GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridItens.TextMatrix(iIndice2, iGrid_Produto_Col) And bTeste Then dQuantidade = dQuantidade + StrParaDbl(GridDistribuicao.TextMatrix(iIndice, iGrid_Quant_Col))
        Next
        
        If dQuantidade > StrParaDbl(GridItens.TextMatrix(iIndice2, iGrid_Quantidade_Col)) Then gError 86123
    Next

    'verificar se nao há repeticao de produto-almoxarifado-ccl
    For iIndice = 1 To objGridDistribuicao.iLinhasExistentes
    
        For iIndice2 = 1 To objGridDistribuicao.iLinhasExistentes
        
            '##################################
            'Inserido por Wagner 27/07/2006
            bTeste = True
            Call CF2(Me, "PC_Produto_Igual_Customizado2", iIndice, iIndice2, bTeste)
            '##################################
        
            If bTeste And iIndice2 <> iIndice And _
                GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Produto_Col) And _
                GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Almoxarifado_Col) And _
                GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_CentroCusto_Col) Then gError 81625
        
        Next
    
    Next
    
    Valida_Grid_Distribuicao = SUCESSO

    Exit Function

Erro_Valida_Grid_Distribuicao:

    Valida_Grid_Distribuicao = gErr

    Select Case gErr

        Case 56191
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_ITEM_DISTRIBUICAO_VAZIO", gErr)

        Case 56192
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_DISTRIBUICAO_VAZIO", gErr)

        Case 86123
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTDISTRIBUIDA_SUPERIOR_ITEM", gErr, iIndice2)

        Case 81625
            Call Rotina_Erro(vbOKOnly, "ERRO_LOC_ITEM_PC_DUPLICADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164629)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Bloqueios() As Long

Dim iIndice As Integer

On Error GoTo Erro_Valida_Grid_Bloqueios

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        'Testa se preencheu tipo de bloqueio
        If Len(Trim(GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))) = 0 Then Error 56170

    Next

    Valida_Grid_Bloqueios = SUCESSO

    Exit Function

Erro_Valida_Grid_Bloqueios:

    Valida_Grid_Bloqueios = Err

    Select Case Err

        Case 56170
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO", Err, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164630)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dtDataLimite As Date
Dim dtData As Date
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Valida_Grid_Itens

    For iIndice = 1 To objGridItens.iLinhasExistentes

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 68537

        objFornecedorProdutoFF.sProduto = sProdutoFormatado
        objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
        objFornecedor.sNomeReduzido = Fornecedor.Text

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO Then gError 68538

        objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo
        objFornecedorProdutoFF.iFilialForn = Codigo_Extrai(Filial.Text)

        'Verifica se o produto está vinculado a Fornecedor/ Filial para a FilialEmpresa
        lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
        If lErro <> SUCESSO And lErro <> 54217 Then gError 68539

        'Se não encontrou ==>erro
        If lErro = 54217 Then gError 68540

        'verifica se a quantidade esta preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then

            'Verifica se a quantidade e positiva
            lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            If lErro <> SUCESSO Then Error 56167

        End If

        'Se a quantidade não está preenchida ==> erro
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 76075
        
        'Se o preco unitário não está preenchido ==> erro
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))) = 0 Then gError 76076
        
        'Verifica se o PrecoUnitario esta preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))) > 0 Then

            'Verifica se o PrecoUnitario e positivo
            lErro = Valor_NaoNegativo_Critica(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
            If lErro <> SUCESSO Then Error 56168

        End If

        'Verifica se DataLimite e Data estao preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col))) > 0 And Len(Trim(Data.Text)) > 0 Then

            dtDataLimite = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col))
            dtData = StrParaDate(Data.Text)

            'Verifica se DataLimite e inferior a Data do Pedido
            If dtDataLimite < dtData Then Error 56169

        End If
        
        '################################################
        'Inserido por Wagner 14/07/2006
        lErro = CF2(Me, "PC_Valida_Grid_Itens", iIndice)
        If lErro <> SUCESSO Then gError 181218
        '################################################

    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 56167, 56168, 68537, 68538, 68539, 181218

        Case 56169
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_ITEM_INFERIOR_DATAPEDIDO", gErr, iIndice)

        Case 68540
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORFILIALPRODUTO_NAO_CADASTRADA", gErr, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.lFornecedor)

        Case 76075
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 76076
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164631)

    End Select

    Exit Function

End Function

Private Function Traz_PedidoCompra_Tela(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objObservacao As New ClassObservacao
Dim iIndice As Integer
Dim objTransportadora As New ClassTransportadora
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Traz_PedidoCompra_Tela

    gbCarregandoTela = True
    
    Call Limpa_Tela_PedidoCompras

    gbCarregandoTela = True
    
    lErro = CF("PedidoCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56118 Then gError 178848

    'Le os itens do Pedido de Compra
    lErro = CF("ItensPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 49488

    'Le os Bloqueios do Pedido de Compra
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 49489
    
    lErro = CF("NotasPedCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103347
    
    lErro = CF("ItensPCEntrega_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103347
    
    'Coloca os dados na tela
    Codigo.promptinclude = False
    Codigo.Text = objPedidoCompra.lCodigo
    Codigo.promptinclude = True

    Contato.Text = objPedidoCompra.sContato

    Data.promptinclude = False
    Data.Text = Format(objPedidoCompra.dtData, "dd/mm/yy")
    Data.promptinclude = True

    Fornecedor.Text = objPedidoCompra.lFornecedor
    Fornecedor_Validate (bSGECancelDummy)
    
    Filial.Text = objPedidoCompra.iFilial
    Call Filial_Validate(bSGECancelDummy)

    If objPedidoCompra.dtDataAlteracao <> DATA_NULA Then
        DataAlteracao.Caption = Format(objPedidoCompra.dtDataAlteracao, "dd/mm/yyyy")
    End If
    
    If objPedidoCompra.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objPedidoCompra.dtDataEmissao, "dd/mm/yyyy")
    End If

    If objPedidoCompra.dtDataEnvio <> DATA_NULA Then
        DataEnvio.Text = Format(objPedidoCompra.dtDataEnvio, "dd/mm/yy")
    End If

    If objPedidoCompra.dtDataRefFluxo <> DATA_NULA Then
        DataRefFluxo.Text = Format(objPedidoCompra.dtDataRefFluxo, "dd/mm/yy")
    End If

    'Preenche o TipoDestino
    TipoDestino(objPedidoCompra.iTipoDestino).Value = True

    If objPedidoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA Then
        For iIndice = 0 To FilialEmpresa.ListCount - 1
            If objPedidoCompra.iFilialDestino = Codigo_Extrai(FilialEmpresa.List(iIndice)) Then
                FilialEmpresa.ListIndex = iIndice

                Exit For
    
            End If

        Next

    ElseIf objPedidoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then
        Fornec.Text = objPedidoCompra.lFornCliDestino
        Fornec_Validate (bSGECancelDummy)
        FilialFornec.Text = objPedidoCompra.iFilialDestino
        FilialFornec_Validate (bSGECancelDummy)

    End If

    'Preenche o TipoFrete
    For iIndice = 0 To TipoFrete.ListCount - 1
        If objPedidoCompra.sTipoFrete = TipoFrete.ItemData(iIndice) Then
            'coloca na tela
            TipoFrete.Text = TipoFrete.List(iIndice)
        End If
    Next

    If objPedidoCompra.iCondicaoPagto > 0 Then
        CondPagto.Text = objPedidoCompra.iCondicaoPagto
        CondPagto_Validate (bSGECancelDummy)
    End If

    'Se a tabela de preços estiver preenchida coloca na tela
    If objPedidoCompra.iTabelaPreco > 0 Then
        TabelaPreco.Text = objPedidoCompra.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        TabelaPreco.Text = ""
    End If

    'Verifica se Observacao esta preenchido
    If objPedidoCompra.lObservacao <> 0 Then

        objObservacao.lNumInt = objPedidoCompra.lObservacao
        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 56454
        If lErro = 53827 Then gError 57173
        ObservacaoPC.Text = objObservacao.sObservacao

    End If
    ObsEmbalagem.Text = objPedidoCompra.sObsEmbalagem
    'If objPedidoCompra.dValorProdutos > 0 Then ValorProdutos.Caption = Format(objPedidoCompra.dValorProdutos, PrecoTotal.Format) 'Alterado por Wagner
    ValorFrete.Text = Format(objPedidoCompra.dValorFrete, "standard")
    ValorSeguro.Text = Format(objPedidoCompra.dValorSeguro, "standard")
    OutrasDespesas.Text = Format(objPedidoCompra.dOutrasDespesas, "standard")
    ValorDesconto.Text = Format(objPedidoCompra.dValorDesconto, "standard")
    ValorIPI.Text = Format(objPedidoCompra.dValorIPI, "standard")
    ValorTotal.Caption = Format(objPedidoCompra.dValorTotal, PrecoTotal.Format) 'Alterado por Wagner

    'Verifica se a Transportadora está preenchida
    If objPedidoCompra.iTransportadora > 0 Then

        objTransportadora.iCodigo = objPedidoCompra.iTransportadora

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 63656
        If lErro = 19250 Then gError 63657  ''Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

'Leo
    If objPedidoCompra.dTaxa > 0 Then
    
        Taxa.promptinclude = False
        Taxa.Text = Format(objPedidoCompra.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
        Taxa.promptinclude = True
        
    End If
    
    If objPedidoCompra.iEmbalagem > 0 Then
        
        objEmbalagem.iCodigo = objPedidoCompra.iEmbalagem
        
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 103388
        
        If lErro = SUCESSO Then
             
            Embalagem.Text = objEmbalagem.sSigla
            
        End If
            
    End If
        
    For iIndice = 0 To Moeda.ListCount - 1
        
        If objPedidoCompra.iMoeda = Codigo_Extrai(Moeda.List(iIndice)) Then
            
            Moeda.ListIndex = iIndice
            
            Exit For
                     
        End If
                     
    Next
    
    'preenche o Grid com os Ítens do Pedido Compra
    lErro = Preenche_Grid_Itens(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56021

    'Preenche o GridBloqueio
    lErro = Preenche_Grid_Bloqueio(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56022

    'Preenche o GridDistribuicao
    lErro = Preenche_Grid_Distribuicao(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56023

    'preenche o Grid com as Notas do Pedido Compra
    lErro = Preenche_Grid_Notas(objPedidoCompra)
    If lErro <> SUCESSO Then gError 103346
    
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51336

    If objPedidoCompra.lCodigoPV <> 0 Then
        CodigoPV.promptinclude = False
        CodigoPV.Text = objPedidoCompra.lCodigoPV
        CodigoPV.promptinclude = True
    End If
    
    Set gobjInfoAdic = objPedidoCompra.objInfoAdic
    
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objPedidoCompra.objProjetoInfo, objPedidoCompra.lNumIntDoc, PRJ_CR_TIPO_PC)
    If lErro <> SUCESSO Then gError 51336

    gbCarregandoTela = False
    
    iAlterado = 0

    Traz_PedidoCompra_Tela = SUCESSO

    Exit Function

Erro_Traz_PedidoCompra_Tela:

    gbCarregandoTela = False
    
    Traz_PedidoCompra_Tela = gErr

    Select Case gErr

        Case 49488, 49489, 49490, 56021, 56022, 56023, 56405, 56454, 63656, 103346, 103347, 103388, 178848

        Case 56413
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
            FilialEmpresa.SetFocus

        Case 57173
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_INEXISTENTE", gErr, objObservacao.lNumInt)

        Case 63657
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164632)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Bloqueio(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPC As New ClassBloqueioPC
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_Preenche_Grid_Bloqueio

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC

        iIndice = iIndice + 1

        GridBloqueios.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objBloqueioPC.sCodUsuario
        GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelBL_Col) = objBloqueioPC.sResponsavel
        GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioPC.sCodUsuarioLib

        objTipoDeBloqueioPC.iCodigo = objBloqueioPC.iTipoBloqueio

        lErro = CF("TipoDeBloqueioPC_Le", objTipoDeBloqueioPC)
        If lErro <> SUCESSO And lErro <> 49143 Then Error 57250
        If lErro = 49143 Then Error 57251

        GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objBloqueioPC.iTipoBloqueio & SEPARADOR & objTipoDeBloqueioPC.sNomeReduzido

        If objBloqueioPC.dtDataLib <> DATA_NULA Then GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPC.dtDataLib, "dd/mm/yyyy")
        If (objBloqueioPC.dtData <> DATA_NULA) Then GridBloqueios.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPC.dtData, "dd/mm/yyyy")
    Next

    objGridBloqueio.iLinhasExistentes = iIndice

    iGridBloqueiosLinhasBloqueadas = iIndice

    Preenche_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Preenche_Grid_Bloqueio:

    Preenche_Grid_Bloqueio = Err

    Select Case Err
        
        Case 57250

        Case 57251
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164633)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Distribuicao(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim sCclMascarado As String
Dim sContaMascarada As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Preenche_Grid_Distribuicao

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridDistribuicao)

    iIndice = 0
    iItem = 0

    For Each objItemPC In objPedidoCompra.colItens

        iItem = iItem + 1

        For Each objLocalizacao In objItemPC.colLocalizacao

            iIndice = iIndice + 1

            'Coloca os dados de distribuicao na tela
            GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridItens.TextMatrix(iItem, iGrid_Produto_Col)
            GridDistribuicao.TextMatrix(iIndice, iGrid_Quant_Col) = Formata_Estoque(objLocalizacao.dQuantidade)
            GridDistribuicao.TextMatrix(iIndice, iGrid_DescProd_Col) = GridItens.TextMatrix(iItem, iGrid_DescProduto_Col)
            GridDistribuicao.TextMatrix(iIndice, iGrid_UnidMed_Col) = GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col)

            'Limpa a variável que será utilizada para obter o Ccl
            sCclMascarado = ""
            
            'Se o campo Ccl foi preenchido
            If Len(Trim(objLocalizacao.sCcl)) > 0 Then
                
                'Inicializa a string que receberá o Ccl
                sCclMascarado = String(STRING_CCL, 0)
                
                'Chama a função que irá formatar o Ccl
                lErro = Mascara_RetornaCclEnxuta(objLocalizacao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 56401
                
                'Preenche o campo Ccl com o Ccl encontrado
                CentroCusto.promptinclude = False
                CentroCusto.Text = sCclMascarado
                CentroCusto.promptinclude = True

                'Joga o Ccl no grid
                GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col) = CentroCusto.Text

            End If

            If Len(Trim(objLocalizacao.sContaContabil)) <> 0 Then

                lErro = Mascara_RetornaContaEnxuta(objLocalizacao.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then Error 56402

                ContaContabil.promptinclude = False
                ContaContabil.Text = sContaMascarada
                ContaContabil.promptinclude = True
                
                GridDistribuicao.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text

            End If

            If objLocalizacao.iAlmoxarifado > 0 Then
                objAlmoxarifado.iCodigo = objLocalizacao.iAlmoxarifado

                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then Error 51336
                If lErro <> SUCESSO Then Error 56403

                GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            End If
        Next

    Next

    objGridDistribuicao.iLinhasExistentes = iIndice

    Preenche_Grid_Distribuicao = SUCESSO

    Exit Function

Erro_Preenche_Grid_Distribuicao:

    Preenche_Grid_Distribuicao = Err

    Select Case Err

        Case 56401, 56402, 51336

        Case 56403
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164634)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Itens(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim dPrecoTotal As Double
Dim objProduto As New ClassProduto
Dim objObservacao As New ClassObservacao

On Error GoTo Erro_Preenche_Grid_Itens

    Set gcolItemPedido = New Collection
    Set gobjPC = objPedidoCompra

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    iIndice = 0

    For Each objItemPC In objPedidoCompra.colItens

        iIndice = iIndice + 1

        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 56106

        'Mascara o produto enxuto
        Produto.promptinclude = False
        Produto.Text = sProdutoMascarado
        Produto.promptinclude = True

        'Calcula o percentual de desconto
        If (objItemPC.dPrecoUnitario * objItemPC.dQuantidade) <> 0 Then
            dPercDesc = objItemPC.dValorDesconto / (objItemPC.dPrecoUnitario * objItemPC.dQuantidade)
        Else
            dPercDesc = 0
        End If
        
        dPrecoTotal = (objItemPC.dPrecoUnitario * objItemPC.dQuantidade) - objItemPC.dValorDesconto

        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text

        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemPC.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemPC.sUM
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemPC.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItemPC.dPrecoUnitario, PrecoUnitario.Format) 'Alterado por Wagner
        If dPercDesc > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentDesc_Col) = Format(dPercDesc, "Percent")
        If objItemPC.dValorDesconto > 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemPC.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner

        '#################################################
        'Inserido por Wagner 14/07/2006
        lErro = CF2(Me, "PC_Carrega_Grid_Itens_Linha", objItemPC, iIndice)
        If lErro <> SUCESSO Then gError 181212
        '#################################################

        If objItemPC.dtDataLimite <> DATA_NULA Then
            GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col) = Format(objItemPC.dtDataLimite, "dd/mm/yyyy")
            
            If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
                
                If objItemPC.dtDeliveryDate = DATA_NULA Then objItemPC.dtDeliveryDate = objItemPC.dtDataLimite
            
                GridItens.TextMatrix(iIndice, iGrid_DeliveryDate_Col) = Format(objItemPC.dtDeliveryDate, "dd/mm/yyyy")
                
                TempoTransito.promptinclude = False
                TempoTransito.Text = CStr(objItemPC.iTempoTransito)
                TempoTransito.promptinclude = True
                
                GridItens.TextMatrix(iIndice, iGrid_TempoTransito_Col) = TempoTransito.Text
            
            End If
            
        End If

        GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col) = Format(objItemPC.dPercentMaisReceb, "Percent")

        If objItemPC.dAliquotaIPI <> -1 Then
            GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objItemPC.dAliquotaIPI, "Percent")
        End If

        If objItemPC.dAliquotaICMS <> -1 Then
            GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objItemPC.dAliquotaICMS, "Percent")
        End If

        If objItemPC.lObservacao > 0 Then

            objObservacao.lNumInt = objItemPC.lObservacao

            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then gError 51337
            If lErro <> SUCESSO Then gError 51338

            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao

        End If

        For iItem = 0 To RecebForaFaixa.ListCount - 1
            If objItemPC.iRebebForaFaixa = RecebForaFaixa.ItemData(iItem) Then
                'coloca no Grid Itens RecebForaFaixa
                GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem)
            End If
        Next

        GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col) = objItemPC.dValorIPI
                
        'Le o produto
        objProduto.sCodigo = objItemPC.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 56381
        'Se nao encontrou => erro
        If lErro = 28030 Then gError 56437
        
        'Preenche a descrição completa do produto com a ObsFisica do produto na tabela de produtos
        GridItens.TextMatrix(iIndice, iGrid_DescCompleta_Col) = objProduto.sObsFisica
        
        'Armazena os números internos dos itens
        gcolItemPedido.Add objItemPC.lNumIntDoc

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    If objPedidoCompra.dTaxa > 0 Then
        Call ComparativoMoedaReal_Calcula(objPedidoCompra.dTaxa)
    End If

    Exit Function

Erro_Preenche_Grid_Itens:

    Preenche_Grid_Itens = gErr

    Select Case gErr

        Case 56106, 56404, 51337, 56381, 181212

        Case 56437
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 51338
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objObservacao.lNumInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164635)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Tela_Extrai

    sTabela = "PedidoCompra_Fornecedor"
    
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 49486
    
    
    
    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objPedidoCompra.lCodigo, 0, "Codigo"
    colCampoValor.Add "OutrasDespesas", objPedidoCompra.dOutrasDespesas, 0, "Outras Despesas"
    colCampoValor.Add "Data", objPedidoCompra.dtData, 0, "Data"
    colCampoValor.Add "DataAlteracao", objPedidoCompra.dtDataAlteracao, 0, "Data de Alteracao"
    colCampoValor.Add "DataEnvio", objPedidoCompra.dtDataEnvio, 0, "Data Envio"
    colCampoValor.Add "ValorDesconto", objPedidoCompra.dValorDesconto, 0, "Valor Desconto"
    colCampoValor.Add "ValorFrete", objPedidoCompra.dValorFrete, 0, "Valor Frete"
    colCampoValor.Add "ValorIPI", objPedidoCompra.dValorIPI, 0, "Valor IPI"
    colCampoValor.Add "ValorSeguro", objPedidoCompra.dValorSeguro, 0, "Valor Seguro"
    colCampoValor.Add "ValorTotal", objPedidoCompra.dValorTotal, 0, "Valor Total"
    colCampoValor.Add "Comprador", objPedidoCompra.iComprador, 0, "Comprador"
    colCampoValor.Add "CondicaoPagto", objPedidoCompra.iCondicaoPagto, 0, "Condicao Pagto"
    colCampoValor.Add "Filial", objPedidoCompra.iFilial, 0, "Filial"
    colCampoValor.Add "FilialDestino", objPedidoCompra.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "FilialEmpresa", objPedidoCompra.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "ProxSeqBloqueio", objPedidoCompra.iProxSeqBloqueio, 0, "ProxSeqBloqueio"
    colCampoValor.Add "TipoBaixa", objPedidoCompra.iTipoBaixa, 0, "TipoBaixa"
    colCampoValor.Add "TipoDestino", objPedidoCompra.iTipoDestino, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objPedidoCompra.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "Fornecedor", objPedidoCompra.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "NumIntDoc", objPedidoCompra.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Transportadora", objPedidoCompra.iTransportadora, 0, "Transportadora"
    colCampoValor.Add "Alcada", objPedidoCompra.sAlcada, STRING_BUFFER_MAX_TEXTO, "Alcada"
    colCampoValor.Add "Contato", objPedidoCompra.sContato, STRING_BUFFER_MAX_TEXTO, "Contato"
    colCampoValor.Add "MotivoBaixa", objPedidoCompra.sMotivoBaixa, STRING_BUFFER_MAX_TEXTO, "MotivoBaixa"
    colCampoValor.Add "Observacao", objPedidoCompra.lObservacao, 0, "Observacao"
    colCampoValor.Add "TipoFrete", objPedidoCompra.sTipoFrete, STRING_BUFFER_MAX_TEXTO, "TipoFrete"
    colCampoValor.Add "DataEmissao", objPedidoCompra.dtDataEmissao, 0, "Data de Emissao"
    colCampoValor.Add "TabelaPreco", objPedidoCompra.iTabelaPreco, 0, "TabelaPreco"
'''    colCampoValor.Add "NomeReduzido", Fornecedor.Text, STRING_BUFFER_MAX_TEXTO, "NomeReduzido"

'leo
    colCampoValor.Add "Moeda", objPedidoCompra.iMoeda, 0, "Moeda"
    colCampoValor.Add "Embalagem", objPedidoCompra.iEmbalagem, 0, "Embalagem"
    colCampoValor.Add "Taxa", objPedidoCompra.dTaxa, 0, "Taxa"
'leo
    colCampoValor.Add "ObsEmbalagem", objPedidoCompra.sObsEmbalagem, STRING_BUFFER_MAX_TEXTO, "ObsEmbalagem"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "DataEnvio", OP_IGUAL, DATA_NULA

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 49486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164636)

    End Select

    Exit Sub

End Sub

Public Sub ObsLabel_Click()

Dim objObservacao As New ClassObservacao
Dim colSelecao As New Collection

    'Coloca em objObservacao a Observacao da tela
    objObservacao.sObservacao = ObservacaoPC.Text

    Call Chama_Tela("ObservacaoLista", colSelecao, objObservacao, objEventoObservacao)

    Exit Sub

End Sub

Public Sub OutrasDespesas_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dOutrasDespesas As Double

On Error GoTo Erro_OutrasDespesas_Validate

    'Verifica se o OutrasDespesas esta preenchido
    If Len(Trim(OutrasDespesas.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(OutrasDespesas.Text)
        If lErro <> SUCESSO Then Error 56171

        dOutrasDespesas = StrParaDbl(OutrasDespesas.Text)

        'Coloca o valor no formato standard da tela
        OutrasDespesas.Text = Format(dOutrasDespesas, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56172

    Exit Sub

Erro_OutrasDespesas_Validate:

    Cancel = True

    Select Case Err

        Case 56171, 56172

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164637)

    End Select

    Exit Sub


End Sub

Public Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Public Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long

On Error GoTo Erro_TipoDestino_Click

   If Index = iFrameTipoDestinoAtual Then Exit Sub

   'Torna Frame correspondente a Index visivel
   FrameTipo(Index).Visible = True
   'Torna Frame atual invisivel
   FrameTipo(iFrameTipoDestinoAtual).Visible = False
   'Armazena novo valor de iFrameTipoDestinoAtual
   iFrameTipoDestinoAtual = Index

   'Limpa o frame de Endereco
   Call Limpa_Frame_Endereco

    If Index <> TIPO_DESTINO_EMPRESA Then

        'Limpa o Grid de Distribuicao
        Call Grid_Limpa(objGridDistribuicao)

        Frame1(4).Enabled = False

        'Verifica se o Fornecedor e sua Filial estão preenchidos
        If Len(Trim(Fornec.Text)) > 0 And Len(Trim(FilialFornec.Text)) > 0 Then

            FilialFornec_Click
        End If
    Else

        Frame1(4).Enabled = True
       
        If gobjCOM.iFilialCompra > 0 Then
            Call CF("Filial_Seleciona", FilialEmpresa, gobjCOM.iFilialCompra)
        Else
            Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)
        End If
        FilialEmpresa_Click

    End If

    Exit Sub

Erro_TipoDestino_Click:

    Select Case Err

        Case 56073

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164638)

    End Select

    Exit Sub

End Sub

Public Sub TipoFrete_Click()

Dim lErro As Long

    'Se nao Tiver nenhum TipoFrete selecionado
    If TipoFrete.ListIndex = -1 Then

        'Limpa e desabilita combo de Transportadora
        Transportadora.ListIndex = -1
        Transportadora.Enabled = False
        Exit Sub

    Else    'Se tiver um TipoFrete selecionado

        'Se for TIPO_FOB
        If TipoFrete.ItemData(TipoFrete.ListIndex) = TIPO_FOB Then

            'Habilita a combo de Transportadora
            Transportadora.Enabled = True
            TransportadoraLabel.Enabled = True

        Else    'Se nao for TIPO_FOB

            'Limpa  e desabilita a combo Transportadora
            Transportadora.ListIndex = -1
            Transportadora.Enabled = False
            TransportadoraLabel.Enabled = False

        End If

    End If

    Exit Sub

End Sub

Public Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    If Len(Trim(Transportadora.Text)) > 0 Then

        If Transportadora.ListIndex <> -1 Then Exit Sub

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Transportadora, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56074

        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then

            objTransportadora.iCodigo = iCodigo

            'Tenta ler Transportadora com esse código no BD
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then Error 56075
            If lErro = 19250 Then Error 56076  ''Não encontrou Transportadora no BD

            'Encontrou Transportadora no BD, coloca no Text da Combo
            Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

        End If

    End If

    'Não existe o ítem com a String na List da Combo
    If lErro = 6731 Then Error 56077

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case Err

        Case 56074, 56075

        Case 56076
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case 56077
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164639)

    End Select

    Exit Sub

End Sub

Public Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection

    'Verifica se transportadora esta preenchida
    If Len(Trim(Transportadora.Text)) > 0 Then
        objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)
    End If

    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

    Exit Sub

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui um dia em Data
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 49443

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case Err

        Case 49443

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164640)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta um dia em Data
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 49444

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case Err

        Case 49444

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164641)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataEnvio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_DownClick

    'Diminui um dia em DataEnvio
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 49446

    Exit Sub

Erro_UpDownDataEnvio_DownClick:

    Select Case Err

        Case 49446

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164642)

    End Select

    Exit Sub

End Sub
Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dAliquotaICMS As Double
Dim objICMSAliqExt As New ClassICMSAliqExterna
Dim objFornecProduto As New ClassFornecedorProduto
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objFornecedor As New ClassFornecedor
Dim objItemPC As New ClassItemPedCompra

On Error GoTo Erro_ProdutoLinha_Preenche
    
    'Preenche unidade de medida e descricao do produto
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
    
    'Se alíquotta IPI do Produto estiver preenchida
    If objProduto.dIPIAliquota <> -1 Then
        'Coloca no grid a alíquota IP do produto
        GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaIPI_Col) = Format(objProduto.dIPIAliquota, "Percent")
    End If
    
    'Preenche a descrição completa do produto com a ObsFisica do produto na tabela de produtos
    GridItens.TextMatrix(GridItens.Row, iGrid_DescCompleta_Col) = objProduto.sObsFisica 'leo
    
    'De acordo com o Estado da giFilialEmpresa e o Estado
    'da Filial de destino (FilialEmpresa ou FilialFornecedor)
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaICMS_Col))) = 0 Then
        
        lErro = ICMS_ObtemAliquotaPadrao(objProduto.sCodigo, dAliquotaICMS)
        If lErro <> SUCESSO Then gError 103361
        
        If dAliquotaICMS <> -1 Then GridItens.TextMatrix(GridItens.Row, iGrid_AliquotaICMS_Col) = Format(dAliquotaICMS, "Percent")
        
    End If
    
    If objProduto.iTemFaixaReceb <> 1 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_PercentMaisReceb_Col) = Format(objProduto.dPercentMaisReceb, "PERCENT")

        If objProduto.iRecebForaFaixa = REJEITA_RECEBIMENTO Then


            'coloca no Grid Itens RecebForaFaixa
            GridItens.TextMatrix(GridItens.Row, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(ERRO_E_REJEITA_RECEBIMENTO)

        Else
        
            'coloca no Grid Itens RecebForaFaixa
            GridItens.TextMatrix(GridItens.Row, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(AVISA_E_ACEITA_RECEBIMENTO)
        
        End If
    Else
        
        'coloca no Grid Itens RecebForaFaixa
        GridItens.TextMatrix(GridItens.Row, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(NAO_AVISA_E_ACEITA_RECEBIMENTO)
        
    End If

    'Por Leo
    'se o fornecedor estiver preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'lê o código do fornecedor através do nomereduzido
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 103360
        
        'se o fornecedor existe
        If lErro = SUCESSO Then
        
            objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo
            objFornecedorProdutoFF.sProduto = objProduto.sCodigo
            objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
            objFornecedorProdutoFF.iFilialForn = Codigo_Extrai(Filial.Text)
        
            'verifica se o produto em questão está associado ao fornecedor
            lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 54217 Then gError 103361
                
            'se não hover associação
            If lErro <> SUCESSO Then
            
                'pergunta se deseja criar a associação
                lErro = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_FORNECEDOR_DESEJA_CADASTRAR", objProduto.sCodigo, Fornecedor.Text)
                
                'Se o usuário respondeu que deseja
                If lErro = vbYes Then
                    
                    'Chama a tela de produto vs fornecedor
                    Call Chama_Tela("FornFilialProduto", objFornecedorProdutoFF)
                    
                End If
                    
            End If
                
        End If
    
    End If
    'Leo até aqui
    
    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        gcolItemPedido.Add (0)
        gobjPC.colItens.Add objItemPC
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

    End If

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr
        
        Case 103360, 103361 'leo
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164643)

    End Select

    Exit Function

End Function
Public Sub UpDownDataEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_UpClick

    'Aumenta um dia em DataEnvio
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 49447

    Exit Sub

Erro_UpDownDataEnvio_UpClick:

    Select Case Err

        Case 49447

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164644)

    End Select

    Exit Sub

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o ValorDesconto foi preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then Error 49450

        dValorDesconto = StrParaDbl(ValorDesconto.Text)

        'Coloca o valor no formato standard da tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

    End If

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then Error 56124

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case Err

        Case 49450, 56124

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164645)

    End Select

    Exit Sub

End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorFrete As Double

On Error GoTo Erro_ValorFrete_Validate

    'Verifica se ValorFrete foi preenchido
    If Len(Trim(ValorFrete.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then Error 49448

        dValorFrete = StrParaDbl(ValorFrete.Text)

        'Coloca o valor no formato standard da tela
        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56125

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case Err

        Case 49448, 56125

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164646)

    End Select

    Exit Sub

End Sub

Public Sub ValorIPI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorIPI As Double

On Error GoTo Erro_ValorIPI_Validate

    'Verifica se o ValorIPI foi preenchido
    If Len(Trim(ValorIPI.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorIPI.Text)
        If lErro <> SUCESSO Then Error 49448

        dValorIPI = StrParaDbl(ValorIPI.Text)

        'Coloca o valor no formato standard da tela
        ValorIPI.Text = Format(dValorIPI, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56127

    Exit Sub

Erro_ValorIPI_Validate:

    Cancel = True

    Select Case Err

        Case 49448, 56127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164647)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorSeguro As Double

On Error GoTo Erro_ValorSeguro_Validate

    'Verifica se o ValorSeguro esta preenchido
    If Len(Trim(ValorSeguro.Text)) > 0 Then

        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then Error 49449

        dValorSeguro = StrParaDbl(ValorSeguro.Text)

        'Coloca o valor no formato standard da tela
        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 56119

    Exit Sub

Erro_ValorSeguro_Validate:

    Cancel = True

    Select Case Err

        Case 49449, 56119

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164648)

    End Select

    Exit Sub

End Sub

Private Function Move_Pedido_Memoria(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Pedido_Memoria

    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56212

    lErro = Move_GridItens_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56213

    lErro = Move_GridDistribuicao_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56214

    lErro = Move_GridBloqueio_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56215
    
    lErro = Move_GridNotas_Memoria(objPedidoCompra) '??? por leo
    If lErro <> SUCESSO Then gError 103331
    
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_PC)
    If lErro <> SUCESSO Then gError 181761
    
    Set objPedidoCompra.objProjetoInfo = objProjetoInfo
    
    Set objPedidoCompra.objInfoAdic = gobjInfoAdic
    
    Move_Pedido_Memoria = SUCESSO

    Exit Function

Erro_Move_Pedido_Memoria:

    Move_Pedido_Memoria = gErr

    Select Case gErr

        Case 56212, 56213, 56214, 56215, 103331
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164649)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objComprador As New ClassComprador
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Move_Tela_Memoria

    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor.ClipText)) > 0 Then

        Set objFornecedor = New ClassFornecedor

        objFornecedor.sNomeReduzido = Fornecedor.Text
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 56016
        If lErro = 6681 Then gError 56017 'Se nao encontrar --> erro

        objPedidoCompra.lFornecedor = objFornecedor.lCodigo

    End If

    'Recolhe os demais campos da Tela
    objPedidoCompra.iFilial = Codigo_Extrai(Filial.Text)

    'Frame Cabecalho
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCompra.sContato = Contato.Text

    If Len(Trim(Comprador.Caption)) > 0 Then

        objComprador.sCodUsuario = gsUsuario
        lErro = CF("Comprador_Le_Usuario", objComprador)
        If lErro <> SUCESSO And lErro <> 50059 Then gError 56378
        If lErro = 50059 Then gError 56448

        objPedidoCompra.iComprador = objComprador.iCodigo

    End If

    objPedidoCompra.iCondicaoPagto = CondPagto_Extrai(CondPagto)
    objPedidoCompra.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)

    objPedidoCompra.dtData = StrParaDate(Data.Text)
    objPedidoCompra.dtDataAlteracao = StrParaDate(DataAlteracao.Caption)
    objPedidoCompra.dtDataEmissao = StrParaDate(DataEmissao.Caption)
    objPedidoCompra.dtDataEnvio = StrParaDate(DataEnvio.Text)
    objPedidoCompra.dtDataRefFluxo = StrParaDate(DataRefFluxo.Text)
    objPedidoCompra.sObservacao = ObservacaoPC.Text
    objPedidoCompra.sObsEmbalagem = ObsEmbalagem.Text
    'Frame Valores
    objPedidoCompra.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objPedidoCompra.dValorFrete = StrParaDbl(ValorFrete.Text)
    objPedidoCompra.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objPedidoCompra.dOutrasDespesas = StrParaDbl(OutrasDespesas.Text)
    objPedidoCompra.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objPedidoCompra.dValorIPI = StrParaDbl(ValorIPI.Text)
    objPedidoCompra.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa
    objPedidoCompra.dtDataBaixa = DATA_NULA
    objPedidoCompra.lCodigoPV = StrParaDbl(CodigoPV.Text)

    'Frame Local de Entrega

    If TipoDestino(0) Then
        objPedidoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA
        objPedidoCompra.iFilialDestino = Codigo_Extrai(FilialEmpresa.Text)
        
        If objPedidoCompra.iFilialDestino = 0 Then gError 185056
        
    ElseIf TipoDestino(1) Then
        objPedidoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR
        objPedidoCompra.iFilialDestino = Codigo_Extrai(FilialFornec.Text)

        If Len(Trim(Fornec.Text)) > 0 Then

            objFornecedor.sNomeReduzido = Fornec.Text

            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 56493
            If lErro = 6681 Then gError 56494

            'Preenche o Codigo do Fornecedor
            objPedidoCompra.lFornCliDestino = objFornecedor.lCodigo   'OK Como você vai extrair o código de um nome reduzido?

        'Se o Fornecedor não está preeenchido
        Else
            gError 68382
        End If
    End If

    'Frame Frete
    objPedidoCompra.iTransportadora = Codigo_Extrai(Transportadora.Text)
    'OK O Tipo Frete é 0 ou 1. É para pegar o valor do itemdata
    If TipoFrete.ListIndex <> -1 Then
        objPedidoCompra.sTipoFrete = TipoFrete.ItemData(TipoFrete.ListIndex)
    End If

'Leo

    objPedidoCompra.dTaxa = StrParaDbl(Taxa.Text)
    objPedidoCompra.iMoeda = Codigo_Extrai(Moeda.Text)
    
    If Len(Trim(Embalagem.Text)) > 0 Then

        objEmbalagem.sSigla = Embalagem.Text

        lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 95088 Then gError 103380
        
        If lErro <> SUCESSO Then gError 103379
        
        objPedidoCompra.iEmbalagem = objEmbalagem.iCodigo
        
    End If
    
'Leo até aqui
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 56016, 56378, 56493, 103380

        Case 56017
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case 56448
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO1", gErr, objComprador.sCodUsuario)

        Case 56494
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 68382
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_ENTREGA_NAO_PREENCHIDO", gErr)

        Case 103379
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_INEXISTENTE", gErr, objEmbalagem.sSigla)
            
        Case 185056
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_ENTREGA_NAO_PREENCHIDA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164650)

    End Select

    Exit Function

End Function
Private Sub Limpa_Frame_Endereco()

    Endereco.Caption = ""
    Bairro.Caption = ""
    Cidade.Caption = ""
    CEP.Caption = ""
    Estado.Caption = ""
    Pais.Caption = ""

    Exit Sub

End Sub

Private Function Move_GridItens_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemPedCompra As New ClassItemPedCompra
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim iItem As Integer
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemPedCompra = New ClassItemPedCompra

        'Verifica se o Produto está preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 56193

            objItemPedCompra.sProduto = sProduto
        End If

        'Armazena os dados do item
        objItemPedCompra.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemPedCompra.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemPedCompra.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
        objItemPedCompra.dtDataLimite = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col))
        objItemPedCompra.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        objItemPedCompra.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)
        objItemPedCompra.dPercentMaisReceb = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col))
        
        If gobjCOM.iPCExibeDeliveryDate = MARCADO Then
            objItemPedCompra.dtDeliveryDate = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DeliveryDate_Col))
            objItemPedCompra.iTempoTransito = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_TempoTransito_Col))
        End If

        'Verifica se AliquotaICMS foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col))) = 0 Then
            objItemPedCompra.dAliquotaICMS = -1
        Else
            objItemPedCompra.dAliquotaICMS = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col))
        End If

        'Verifica se AliquotaIPI foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col))) = 0 Then
            objItemPedCompra.dAliquotaIPI = -1
        Else
            objItemPedCompra.dAliquotaIPI = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col))
        End If

        objItemPedCompra.dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))

         For iItem = 0 To RecebForaFaixa.ListCount - 1
            If GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem) Then

                'coloca no Grid Itens RecebForaFaixa
                objItemPedCompra.iRebebForaFaixa = RecebForaFaixa.ItemData(iItem)

            End If
        Next


        objItemPedCompra.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItemPedCompra.lNumIntDoc = gcolItemPedido.Item(iIndice)

        objItemPedCompra.iStatus = ITEM_PED_COMPRAS_ABERTO
        
        '#############################################################
        'Inserido por Wagner 14/07/2006
        lErro = CF2(Me, "PC_Move_GridItem_Memoria", objItemPedCompra, iIndice)
        If lErro <> SUCESSO Then gError 181219
        '#############################################################

        Set objItemPedCompra.colDataEntrega = gobjPC.colItens(iIndice).colDataEntrega
        
        If objItemPedCompra.colDataEntrega.Count > 0 Then
        
            dQuantidade = 0
            For Each objDataEntrega In objItemPedCompra.colDataEntrega
                dQuantidade = dQuantidade + objDataEntrega.dQuantidade
            Next
            
            If dQuantidade > 0 Then
                If Abs(dQuantidade - objItemPedCompra.dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 183245
            End If
        
        End If

        'Adiciona o item na colecao de itens do pedido de compra
        objPedidoCompra.colItens.Add objItemPedCompra


    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 56193, 181219

        Case 183245
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, iIndice, objItemPedCompra.sProduto, objItemPedCompra.dQuantidade, dQuantidade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164651)

    End Select

    Exit Function

End Function

Private Function Move_GridBloqueio_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados do bloqueio no parametro

Dim iIndice As Integer
Dim objBloqueioPC As ClassBloqueioPC

    For iIndice = iGridBloqueiosLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes

        Set objBloqueioPC = New ClassBloqueioPC

        'Armazena os dados do bloqueio
        objBloqueioPC.dtData = StrParaDate(GridBloqueios.TextMatrix(iIndice, iGrid_DataBloqueio_Col))
        objBloqueioPC.dtDataLib = DATA_NULA
        objBloqueioPC.iTipoBloqueio = Codigo_Extrai(GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
        objBloqueioPC.sCodUsuario = GridBloqueios.TextMatrix(iIndice, iGrid_CodUsuario_Col)
        objBloqueioPC.sResponsavel = GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelBL_Col)
        objBloqueioPC.iFilialEmpresa = giFilialEmpresa

        'Adiciona o item na colecao de itens do pedido de compra
        objPedidoCompra.colBloqueiosPC.Add objBloqueioPC

    Next

    Move_GridBloqueio_Memoria = SUCESSO

End Function

Private Function Move_GridDistribuicao_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String, bAchou As Boolean
Dim objItemPC As ClassItemPedCompra
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim sCcl As String
Dim iCclPreenchida As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaFormata As String
Dim iContaPreenchida As Integer
Dim bTeste As Boolean

On Error GoTo Erro_Move_GridDistribuicao_Memoria

    For iIndice = 1 To objGridDistribuicao.iLinhasExistentes

        Set objLocalizacao = New ClassLocalizacaoItemPC

        'Verifica se o Produto está preenchido
        If Len(Trim(GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col))) > 0 Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then Error 56194

            iIndice2 = 1

            For Each objItemPC In objPedidoCompra.colItens
            
                '##################################
                'Inserido por Wagner 27/07/2006
                bTeste = True
                Call CF2(Me, "PC_Produto_Igual_Customizado", iIndice2, iIndice, bTeste)
                '##################################

                'Verifica se objItemPC é igual ao produto do grid
                If objItemPC.sProduto = sProduto And bTeste Then Exit For

                iIndice2 = iIndice2 + 1
            Next

            'Se o produto nao for igual
            If iIndice2 > objPedidoCompra.colItens.Count Then Error 56384

            'preenche objLocalizacao com os dados do grid
            objLocalizacao.dQuantidade = StrParaDbl(GridDistribuicao.TextMatrix(iIndice, iGrid_Quant_Col))

            objAlmoxarifado.sNomeReduzido = GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO Then Error 56385
            objLocalizacao.iAlmoxarifado = objAlmoxarifado.iCodigo

            lErro = CF("Ccl_Formata", GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col), sCcl, iCclPreenchida)
            If lErro <> SUCESSO Then Error 56372
            objLocalizacao.sCcl = sCcl

            lErro = CF("Conta_Formata", GridDistribuicao.TextMatrix(iIndice, iGrid_ContaContabil_Col), sContaFormata, iContaPreenchida)
            If lErro <> SUCESSO Then Error 63658

            objLocalizacao.sContaContabil = sContaFormata

            Set objItemPC = objPedidoCompra.colItens.Item(iIndice2)

            objItemPC.colLocalizacao.Add objLocalizacao

        End If

    Next

    Move_GridDistribuicao_Memoria = SUCESSO

    Exit Function

Erro_Move_GridDistribuicao_Memoria:

    Move_GridDistribuicao_Memoria = Err

    Select Case Err

        Case 56194, 63658
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164652)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValorTotal As Double
Dim objPedidoCompra As New ClassPedidoCompras
Dim dtDataEnvio As Date
Dim dtData As Date

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Codigo foi preenchido
    If (Len(Trim(Codigo.Text))) = 0 Then gError 56040

    'Verifica se Fornecedor foi preenchido
    If Len(Trim(Fornecedor)) = 0 Then gError 56041

    'Verifica se Filial foi preenchido
    If Len(Trim(Filial.Text)) = 0 Then gError 56042

    'Verifica se CondPagto foi preenchido
    If Len(Trim(CondPagto.Text)) = 0 Then gError 56043

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.Text)) = 0 Then gError 56388
    
    'Verifica se a Moeda foi preenchida
    If Len(Trim(Moeda.Text)) = 0 Then gError 103389 'leo
        
    If Len(Trim(DataEnvio.Text)) > 0 Then

        dtDataEnvio = StrParaDate(DataEnvio.Text)
        dtData = StrParaDate(Data.Text)

        'Verifica se DataEnvio < Data e  se DataEnvio é diferente da Data Nula
        If (dtDataEnvio < dtData) And (dtDataEnvio <> DATA_NULA) Then gError 56299

    End If

    'Verifica se nao existem itens no GridItens
    If objGridItens.iLinhasExistentes = 0 Then gError 56216

    'Verifica se ValorTotal foi preenchido
    dValorTotal = StrParaDbl(ValorTotal.Caption)
    If dValorTotal < 0 Then gError 56044

    'Valida os dados do Grid Itens
    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 56045

    'Valida os dados do Grid Distribuicao
    lErro = Valida_Grid_Distribuicao()
    If lErro <> SUCESSO Then gError 56046

    'Valida os dados do Grid Bloqueios
    lErro = Valida_Grid_Bloqueios()
    If lErro <> SUCESSO Then gError 56047

    'Recolhe os dados da tela
    lErro = Move_Pedido_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56048
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 56049

    'Grava o Pedido de Compra
    lErro = CF("PedidoCompra_Grava", objPedidoCompra)
    If lErro <> SUCESSO Then gError 56049
    
    If ImprimePedido.Value = vbChecked Then
    
        lErro = PedidoCompra_Imprimir(objPedidoCompra)
        If lErro <> SUCESSO Then gError 56049
    
    End If

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 56040
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 56041
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

         Case 56042
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 56043
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PREENCHIDA", gErr)

        Case 56044
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_PC_NEGATIVO", gErr)

        Case 56045, 56046, 56047, 56048, 56049
            'Erros tratados nas rotinas chamadas

        Case 56216
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_PC", gErr)

        Case 56299
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INFERIOR_DATAPEDIDO", gErr)

        Case 56388
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 103389
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164653)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Public Function Trata_Parametros(Optional objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            'Le o Pedido de Compra
            lErro = CF("PedidoCompras_Le", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56118 Then Error 56057
            If lErro = 56118 Then Error 56058

            lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
            If lErro <> SUCESSO Then Error 56059

        ElseIf (objPedidoCompra.lNumIntDoc = 0) And (objPedidoCompra.lCodigo <> 0) Then

            'Busca o Pedido de Compra no BD
            lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
            If lErro <> SUCESSO And lErro <> 56142 Then Error 63619

            'Se o Pedido não existir
            If lErro = 56142 Then

                'Coloca o código na tela
                Codigo.Text = objPedidoCompra.lCodigo
            Else
                lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
                If lErro <> SUCESSO Then Error 56059

            End If

        End If

    Else

        Codigo.promptinclude = False
        Codigo.Text = ""
        Codigo.promptinclude = True

    End If

    gbCarregandoTela = False
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    gbCarregandoTela = False
    
    Trata_Parametros = Err

    Select Case Err

        Case 56057, 56059, 56060, 63619

        Case 56058
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164654)

    End Select

    Exit Function

End Function

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub AliquotaIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AliquotaIPI_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub AliquotaIPI_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub AliquotaIPI_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AliquotaIPI
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridBloqueios_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Public Sub GridBloqueios_EnterCell()

    Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)

End Sub

Public Sub GridBloqueios_GotFocus()

    Call Grid_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub GridBloqueios_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Public Sub GridBloqueios_LeaveCell()

    Call Saida_Celula(objGridBloqueio)

End Sub

Public Sub GridBloqueios_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridBloqueio)

End Sub

Public Sub GridBloqueios_RowColChange()

    Call Grid_RowColChange(objGridBloqueio)

End Sub

Public Sub GridBloqueios_Scroll()

    Call Grid_Scroll(objGridBloqueio)

End Sub

Public Sub GridDistribuicao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDistribuicao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDistribuicao, iAlterado)
    End If

End Sub

Public Sub GridDistribuicao_EnterCell()

    Call Grid_Entrada_Celula(objGridDistribuicao, iAlterado)

End Sub

Public Sub GridDistribuicao_GotFocus()

    Call Grid_Recebe_Foco(objGridDistribuicao)

End Sub

Public Sub GridDistribuicao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDistribuicao, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDistribuicao, iAlterado)
    End If

End Sub

Public Sub GridDistribuicao_LeaveCell()

    Call Saida_Celula(objGridDistribuicao)

End Sub

Public Sub GridDistribuicao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridDistribuicao)

End Sub

Public Sub GridDistribuicao_RowColChange()

    Call Grid_RowColChange(objGridDistribuicao)

End Sub

Public Sub GridDistribuicao_Scroll()

    Call Grid_Scroll(objGridDistribuicao)

End Sub

Public Sub ValorIPIItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorIPIItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ValorIPIItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ValorIPIItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorIPIItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AliquotaICMS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AliquotaICMS_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub AliquotaICMS_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub AliquotaICMS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = AliquotaICMS
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub PercentMaisReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMaisReceb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentMaisReceb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentMaisReceb
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub RecebForaFaixa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RecebForaFaixa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RecebForaFaixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub RecebForaFaixa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub RecebForaFaixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = RecebForaFaixa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub Prod_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Prod_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Prod_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Public Sub Prod_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Public Sub Prod_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Prod
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescProd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescProd_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescProd_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DescProd_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DescProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescProd
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub



Public Sub CentroCusto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CentroCusto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CentroCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Public Sub CentroCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Public Sub CentroCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CentroCusto
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub UnidMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Public Sub UnidMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Public Sub UnidMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidMed
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub Quant_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Public Sub Quant_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Public Sub Quant_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quant
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabil_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabil_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDistribuicao)

End Sub

Public Sub ContaContabil_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDistribuicao)

End Sub

Public Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ContaContabil
    lErro = Grid_Campo_Libera_Foco(objGridDistribuicao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoBloqueio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub TipoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub TipoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TipoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ResponsavelBL_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ResponsavelBL_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub ResponsavelBL_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub ResponsavelBL_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = ResponsavelBL
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is CentroCusto Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoAlmoxarifados_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call BotaoContasContabeis_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is Fornec Then
            Call FornLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is ObservacaoPC Then
            Call ObsLabel_Click
        ElseIf Me.ActiveControl Is CondPagto Then
            Call CondPagtoLabel_Click
        ElseIf Me.ActiveControl Is Embalagem Then 'leo
            Call LabelEmbalagem_Click
        End If
    End If
    

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pedido de Compra Avulso"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PedidoCompras"

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

Function Distribuicao_Atualiza(iItem As Integer, objProduto As ClassProduto) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iFilialEmpresa As Integer
Dim sProduto As String
Dim iAlmoxarifadoPadrao As Integer
Dim sConta As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iIndice As Integer, iIndice2 As Integer
Dim sContaMascarada As String
Dim objLocItemPC As ClassLocalizacaoItemPC
Dim objItemPC As ClassItemPedCompra
Dim colItem As New Collection
Dim bTeste As Boolean

On Error GoTo Erro_Distribuicao_Atualiza

    'Guarda a quantidade do grid itens
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iItem, iGrid_Quantidade_Col))
   
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        Set objItemPC = New ClassItemPedCompra
        colItem.Add objItemPC
        
        For iIndice2 = 1 To objGridDistribuicao.iLinhasExistentes
            
            '##################################
            'Inserido por Wagner 27/07/2006
            bTeste = True
            Call CF2(Me, "PC_Produto_Igual_Customizado", iIndice, iIndice2, bTeste)
            '##################################
        
            If (GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridDistribuicao.TextMatrix(iIndice2, iGrid_Prod_Col)) And (iIndice <> iItem) And bTeste Then
                Set objLocItemPC = New ClassLocalizacaoItemPC
                objItemPC.sProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
                objItemPC.sDescProduto = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
                objItemPC.sObservacao = GridDistribuicao.TextMatrix(iIndice2, iGrid_Almoxarifado_Col)
                objItemPC.sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
                objLocItemPC.dQuantidade = StrParaDbl(GridDistribuicao.TextMatrix(iIndice2, iGrid_Quant_Col))
                objLocItemPC.sCcl = GridDistribuicao.TextMatrix(iIndice2, iGrid_CentroCusto_Col)
                objLocItemPC.sContaContabil = GridDistribuicao.TextMatrix(iIndice2, iGrid_ContaContabil_Col)
                objItemPC.colLocalizacao.Add objLocItemPC
            End If
        Next

    Next
    
    Set objItemPC = colItem(iItem)
        
    Call Grid_Limpa(objGridDistribuicao)

    iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)
        
    sProduto = objProduto.sCodigo
    'Busca em FilialProduto o Almoxarifado padrao do produto
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 56185

    If iAlmoxarifadoPadrao > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao

        'Busca a conta contabil na tabela de Almoxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 56342
        If lErro = 25056 Then gError 56436

        sConta = objAlmoxarifado.sContaContabil

    End If

    'Se nao encontrou a conta contabil
    If Len(Trim(objAlmoxarifado.sContaContabil)) = 0 Then

        'Busca a conta contabil do Produto
        sConta = objProduto.sContaContabil

    End If

    'Verifica se a conta contabil esta preenchida
    If Len(Trim(sConta)) > 0 Then

        lErro = Mascara_RetornaContaEnxuta(sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 79766
    
        ContaContabil.promptinclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.promptinclude = True

    End If

    'cria uma linha no grid Distribuicao

    objItemPC.sProduto = GridItens.TextMatrix(iItem, iGrid_Produto_Col)
    objItemPC.sDescProduto = GridItens.TextMatrix(iItem, iGrid_DescProduto_Col)
    objItemPC.sUM = GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col)
    objItemPC.sObservacao = objAlmoxarifado.sNomeReduzido
    Set objLocItemPC = New ClassLocalizacaoItemPC
    objLocItemPC.dQuantidade = dQuantidade
    objLocItemPC.sContaContabil = ContaContabil.Text
    objItemPC.colLocalizacao.Add objLocItemPC

    For Each objItemPC In colItem
        For Each objLocItemPC In objItemPC.colLocalizacao
            objGridDistribuicao.iLinhasExistentes = objGridDistribuicao.iLinhasExistentes + 1
            'Preenche linha do grid distribuicao
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_Almoxarifado_Col) = objItemPC.sObservacao
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_Prod_Col) = objItemPC.sProduto
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_DescProd_Col) = objItemPC.sDescProduto
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_UnidMed_Col) = objItemPC.sUM
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_Quant_Col) = Formata_Estoque(objLocItemPC.dQuantidade)
            GridDistribuicao.TextMatrix(objGridDistribuicao.iLinhasExistentes, iGrid_ContaContabil_Col) = objLocItemPC.sContaContabil
        Next
    Next
    
    Distribuicao_Atualiza = SUCESSO

    Exit Function

Erro_Distribuicao_Atualiza:

    Distribuicao_Atualiza = gErr

    Select Case gErr

        Case 56185, 79766, 56319, 56336 To 56345

        Case 56436
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164655)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'Faz a critica da celula de Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = DescProduto

    If Len(Trim(DescProduto.Text)) = 0 Then gError 86175

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 61965

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 61965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 86175
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164656)

    End Select

    Exit Function

End Function

'??? Leo daqui p/ baixo
Public Sub GridNotas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNotas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNotas, iAlterado)
    End If

End Sub

Public Sub GridNotas_EnterCell()

    Call Grid_Entrada_Celula(objGridNotas, iAlterado)

End Sub

Public Sub GridNotas_GotFocus()

    Call Grid_Recebe_Foco(objGridNotas)

End Sub

Public Sub GridNotas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNotas, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNotas, iAlterado)
    End If

End Sub

Public Sub GridNotas_LeaveCell()

    Call Saida_Celula(objGridNotas)

End Sub

Public Sub GridNotas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNotas)

End Sub

Public Sub GridNotas_RowColChange()

    Call Grid_RowColChange(objGridNotas)

End Sub

Public Sub GridNotas_Scroll()

    Call Grid_Scroll(objGridNotas)

End Sub

Public Sub GridNotas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridNotas)

End Sub

Private Function Inicializa_GridNotas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do gridNotas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nota")
    
    ' campos de edição do grid
    objGridInt.colCampo.Add (NotaPC.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_NotaPC_Col = 1

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridNotas

    'Linhas do grid
    objGridInt.objGrid.Rows = 20

    GridBloqueios.ColWidth(0) = 300

    'linhas visiveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 15
    Else
        objGridInt.iLinhasVisiveis = 4
    End If

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridNotas = SUCESSO

    Exit Function

End Function

Public Sub NotaPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NotaPC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNotas)

End Sub

Public Sub NotaPC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNotas)

End Sub

Public Sub NotaPC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNotas.objControle = NotaPC
    lErro = Grid_Campo_Libera_Foco(objGridNotas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_GridNotas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridNotas
    
    lErro = Saida_Celula_NotasPC(objGridInt)
    If lErro <> SUCESSO Then gError 103328

    Saida_Celula_GridNotas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridNotas:

    Saida_Celula_GridNotas = gErr

    Select Case gErr

        Case 103328

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164657)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NotasPC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NotasPC

    Set objGridInt.objControle = NotaPC

    'Verifica se a nota foi informada
    If Len(Trim(NotaPC.Text)) > 0 Then
        
        GridNotas.TextMatrix(GridNotas.Row, iGrid_NotaPC_Col) = NotaPC.Text
        
        lErro = Grid_Abandona_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 103330
    
    End If

    Saida_Celula_NotasPC = SUCESSO

    Exit Function

Erro_Saida_Celula_NotasPC:

    Saida_Celula_NotasPC = gErr

    Select Case gErr

        Case 103330
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164658)

    End Select

    Exit Function

End Function

Private Function Move_GridNotas_Memoria(objPedidoCompra As ClassPedidoCompras) As Long
'Recolhe do Grid os dados das notas

Dim iIndice As Integer
Dim sNota As String

    For iIndice = 1 To objGridNotas.iLinhasExistentes

        'Armazena os dados das Notas
        sNota = GridNotas.TextMatrix(iIndice, iGrid_NotaPC_Col)
       
        'Adiciona o item na colecao de NotasPedCompras
        objPedidoCompra.colNotasPedCompras.Add sNota

    Next

    Move_GridNotas_Memoria = SUCESSO

End Function

Private Function Preenche_Grid_Notas(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Grid_Notas

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNotas)

    iIndice = 0

    For iIndice = 1 To objPedidoCompra.colNotasPedCompras.Count

        GridNotas.TextMatrix(iIndice, iGrid_NotaPC_Col) = objPedidoCompra.colNotasPedCompras.Item(iIndice)
        
    Next

    objGridNotas.iLinhasExistentes = iIndice - 1

    Preenche_Grid_Notas = SUCESSO

    Exit Function

Erro_Preenche_Grid_Notas:

    Preenche_Grid_Notas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164659)

    End Select

    Exit Function

End Function

Private Sub objEventoEmbalagens_evSelecao(obj1 As Object)
'leo
Dim objEmbalagem As ClassEmbalagem

On Error GoTo Erro_objEventoEmbalagens_evSelecao
                       
    'Define o tipo de obj recebido (Tipo Embalagem)
    Set objEmbalagem = obj1
        
    Embalagem.Text = objEmbalagem.sSigla
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoEmbalagens_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164660)
              
    End Select
    
    Exit Sub
        
End Sub

'leo
Function Carrega_Moeda()

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas) 'leo colocar CF
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164661)
    
    End Select

End Function

Private Sub ComparativoMoedaReal_Calcula(ByVal dTaxa As Double)
'Preenche as colunas INFORMATIVAS de proporção da moeda R$.

Dim iIndice As Integer

On Error GoTo Erro_ComparativoMoedaReal_Calcula

    'Para cada linha do grid de Itens será claculado o correspondente em R$
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) * dTaxa) > 0 Then
            'Preço Unitário em R$ = Preço Unitário na Moeda selecionada dividido pela taxa de conversão
            GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col) = Format((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col)) - (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))) * dTaxa, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        Else
            GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col) = ""
        End If
        
        'Preço Total em R$ = Preço Unitário em R$ x Quantidade do produto
        GridItens.TextMatrix(iIndice, iGrid_TotalMoedaReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col)) * StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), TotalMoedaReal.Format) 'Alterado por Wagner
        
    Next

    Exit Sub
    
Erro_ComparativoMoedaReal_Calcula:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164662)

    End Select

End Sub

Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGridInt.objControle = ContaContabil

    'Se a Conta Contábil foi preenchida
    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabil.ClipText, objPlanoConta, MODULO_COMPRAS)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 61942

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaEnxuta = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 61945

            ContaContabil.promptinclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.promptinclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_COMPRAS)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 61943

            'conta não cadastrada
            If lErro = 5700 Then Error 61946

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 61944

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = Err

    Select Case Err

        Case 61942, 61943, 61944
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61945
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 61946
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164663)

    End Select

End Function

Public Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134066

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164664)

    End Select
    
    Exit Sub

End Sub

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioMoedaReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

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

'########################################################
'Inserido por Wagner 14/07/2006
Public Function PC_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    PC_Calcula_PesoKg = SUCESSO
End Function

Public Function PC_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    PC_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function PC_Carrega_Grid_Itens_Linha(ByVal objCT As Object, ByVal objItemPC As ClassItemPedCompra, ByVal iLinha As Integer) As Long
    PC_Carrega_Grid_Itens_Linha = SUCESSO
End Function

Public Function PC_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objItemPC As ClassItemPedCompra, ByVal iLinha As Integer) As Long
    PC_Move_GridItem_Memoria = SUCESSO
End Function

Public Function PC_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    PC_Saida_Celula_GridItens = SUCESSO
End Function

Public Function PC_Produto_Saida_Celula(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PC_Produto_Saida_Celula

    'Verifica se já existe o Produto em outra linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 181222
        End If

    Next

    PC_Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_PC_Produto_Saida_Celula:

    PC_Produto_Saida_Celula = gErr

    Select Case gErr

        Case 181222
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE_PEDCOMPRA", gErr, Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181220)

    End Select

    Exit Function
    
End Function

Public Function PC_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    PC_Rotina_Grid_Enable = SUCESSO
End Function

Public Function PC_Valida_Grid_Itens(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    
    PC_Valida_Grid_Itens = SUCESSO
End Function

Public Function PrecoTotal_Calcula(ByVal iLinha As Integer) As Long

Dim iIndice As Integer
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPrecoTotal As Double
Dim dValorIPI As Double
Dim dValorTotal As Double
Dim dPrecoUnitario As Double

On Error GoTo Erro_PrecoTotal_Calcula:

    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
    
    Call CF2(Me, "PC_PrecoTotal_Calcula", iLinha, dQuantidade)
        
    'Verifica se percentual de desconto foi preenchido
    If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercentDesc_Col))) > 0 Then

        'calcula o desconto
        dDesconto = (dPrecoUnitario * dQuantidade) * (PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercentDesc_Col)))
        
        If dDesconto > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If
        
    End If
            
    dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
    dPrecoTotal = (dPrecoUnitario * dQuantidade) - dDesconto
    
    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format)

    'Calcula valorIPI
    dValorIPI = dPrecoTotal * (PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_AliquotaIPI_Col)))
    If dValorIPI > 0 Then
        GridItens.TextMatrix(iLinha, iGrid_ValorIPIItem_Col) = Format(dValorIPI, "Standard")
    Else
        GridItens.TextMatrix(iLinha, iGrid_ValorIPIItem_Col) = ""
    End If

    'Calcula o valorIPI da tela
    For iIndice = 1 To objGridItens.iLinhasExistentes
        dValorIPI = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col))
        dValorTotal = dValorTotal + dValorIPI
        ValorIPI.Text = Format(dValorTotal, "standard")
    Next
    
    PrecoTotal_Calcula = SUCESSO

    Exit Function

Erro_PrecoTotal_Calcula:

    PrecoTotal_Calcula = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181221)

    End Select

    Exit Function
    
End Function

Public Function PC_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, dQuantidade As Double) As Long
    PC_PrecoTotal_Calcula = SUCESSO
End Function

Function PC_Produto_Igual_Customizado(ByVal objCT As Object, ByVal iIndice1 As Integer, ByVal iIndice2 As Integer, bTeste As Boolean) As Long
    PC_Produto_Igual_Customizado = SUCESSO
End Function

Function PC_Produto_Igual_Customizado2(ByVal objCT As Object, ByVal iIndice1 As Integer, ByVal iIndice2 As Integer, bTeste As Boolean) As Long
    PC_Produto_Igual_Customizado2 = SUCESSO
End Function

'########################################################

Public Sub CodigoPV_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CodigoPV_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(CodigoPV.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(CodigoPV.Text)
    If lErro <> SUCESSO Then gError 178824

    objPV.iFilialEmpresa = giFilialEmpresa
    objPV.lCodigo = StrParaLong(CodigoPV.Text)
    
    'Lê o Pedido de Venda a partir do Código passado.
    lErro = CF("PedidoDeVenda_Le", objPV)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 178821

    If lErro <> SUCESSO Then
    
        lErro = CF("PedidoVendaBaixado_Le", objPV)
        If lErro <> SUCESSO And lErro <> 46135 Then gError 178822
    
        If lErro <> SUCESSO Then gError 178823
    
    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 178821, 178822, 178824
        
        Case 178823
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178825)

    End Select

    Exit Sub

End Sub

Public Sub CodigoPV_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPV, iAlterado)
    
End Sub

Public Sub DataRefFluxo_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataRefFluxo, iAlterado)
    
End Sub

Public Sub DataRefFluxo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRefFluxo_Validate

    'Verifica se a data foi preenchida
    If Len(DataRefFluxo.ClipText) = 0 Then Exit Sub

    'Critica a data informada
    lErro = Data_Critica(DataRefFluxo.Text)
    If lErro <> SUCESSO Then gError 188593

    Exit Sub

Erro_DataRefFluxo_Validate:

    Cancel = True

    Select Case gErr

        Case 188593

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188594)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataRefFluxo_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataRefFluxo_DownClick

    'Diminui um dia em DataRefFluxo
    lErro = Data_Up_Down_Click(DataRefFluxo, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 188594

    Exit Sub

Erro_UpDownDataRefFluxo_DownClick:

    Select Case gErr

        Case 188594

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188595)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataRefFluxo_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataRefFluxo_UpClick

    'Aumenta um dia em DataRefFluxo
    lErro = Data_Up_Down_Click(DataRefFluxo, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 188596

    Exit Sub

Erro_UpDownDataRefFluxo_UpClick:

    Select Case gErr

        Case 188596

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188597)

    End Select

    Exit Sub

End Sub

Private Function ICMS_ObtemAliquotaPadrao(ByVal sProduto As String, dAliquotaICMS As Double) As Long

Dim lErro As Long, iFilialForn As Integer
Dim objICMSAliqExt As New ClassICMSAliqExterna
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_ICMS_ObtemAliquotaPadrao

    dAliquotaICMS = -1

    objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Text)
    If Len(Trim(Fornecedor.Text)) <> 0 And objFilialFornecedor.iCodFilial <> 0 Then
    
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 184415
        
        If lErro = SUCESSO Then
        
            objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError 184416
            
            'se for importacao
            If objEndereco.iCodigoPais <> 1 Then
            
                'Busca a alíquota e coloca na tela
                For Each objICMSAliqExt In colAliqExternas
            
                    If objICMSAliqExt.sSiglaOrig = objICMSAliqExt.sSiglaDest Then
                        dAliquotaICMS = objICMSAliqExt.dAliquota
                        Exit For
                    End If
                Next

            Else
            
                If objFilialFornecedor.sInscricaoEstadual <> "" And objFilialFornecedor.iGeraCredICMS <> 0 And objFilialFornecedor.iRegimeTributario <> REGIME_TRIBUTARIO_SIMPLES Then
                
                    'Busca a alíquota e coloca na tela
                    For Each objICMSAliqExt In colAliqExternas
                
                        If objICMSAliqExt.sSiglaOrig = objEndereco.sSiglaEstado Then
                            dAliquotaICMS = objICMSAliqExt.dAliquota
                            Exit For
                        End If
                    Next
            
                Else
                
                    dAliquotaICMS = 0
                    
                End If
                
            End If
            
        End If
        
    End If
    
    ICMS_ObtemAliquotaPadrao = SUCESSO
    
    Exit Function
    
Erro_ICMS_ObtemAliquotaPadrao:

    ICMS_ObtemAliquotaPadrao = gErr

    Select Case gErr

        Case 184415, 184416
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Private Function PedidoCompra_Imprimir(objPedidoCompra) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objBloqueioPC As ClassBloqueioPC

On Error GoTo Erro_PedidoCompra_Imprimir

    If objPedidoCompra.lCodigo = 0 Then gError 76054

    'Verifica se o Pedido de Compra informado existe
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76027
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76028
    
'    If objPedidoCompra.dtDataRegAprov = DATA_NULA Then
'        If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
'    End If
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76056
    
    'If objPedidoCompra.colBloqueiosPC.Count > 0 Then gError 76050
    
    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76050
    
    Next
    
    'Alterado por Wagner
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Compra Gerado", "REL_PCOM.PC_NumIntDoc = @NPEDCOM", 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76029
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 53306

    'se nao encontrar ---> erro
    If lErro = 56348 Then gError 53307
    
    PedidoCompra_Imprimir = SUCESSO

    Exit Function

Erro_PedidoCompra_Imprimir:

    PedidoCompra_Imprimir = gErr

    Select Case gErr
    
        Case 53306, 76056

        Case 53307
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76021
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            
        Case 76027, 76029
        
        Case 76028
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
        
        Case 76050
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76054
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Public Function BotaoProjetos_Click() As Long
        Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Function

Public Function LabelProjeto_Click() As Long
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Function

Public Function Projeto_GotFocus() As Long
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Function

Public Function Projeto_Change() As Long
    iAlterado = REGISTRO_ALTERADO
End Function

Public Function Projeto_Validate(Cancel As Boolean) As Long
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR
        
        objItensPRJCR.sItem = CStr(iIndice)
        objItensPRJCR.sDescricao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NF
        
        colItensPRJ.Add objItensPRJCR

    Next
    
    'Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ)
    
    Call Preenche_Itens_Projeto_Aux(colItensPRJ, "FR", "Frete", StrParaDbl(ValorFrete.Text), 1, 0, "")
    Call Preenche_Itens_Projeto_Aux(colItensPRJ, "SG", "Seguro", StrParaDbl(ValorSeguro.Text), 1, 0, "")
    Call Preenche_Itens_Projeto_Aux(colItensPRJ, "OD", "Outras Despesas", StrParaDbl(OutrasDespesas.Text), 1, 0, "")
    Call Preenche_Itens_Projeto_Aux(colItensPRJ, "IPI", "IPI", StrParaDbl(ValorIPI.Text), 1, 0, "")
    Call Preenche_Itens_Projeto_Aux(colItensPRJ, "DESC", "Desconto", StrParaDbl(ValorDesconto.Text), 1, 1, "")

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

Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sObs As String = "") As Long

Dim objItensPRJCR As New ClassItensPRJCR

On Error GoTo Erro_Preenche_Itens_Projeto_Aux

    objItensPRJCR.sItem = sItem
    objItensPRJCR.sDescricao = sDescricao
    objItensPRJCR.dValorOriginal = dValor
    objItensPRJCR.iNegativo = iNegativo
    objItensPRJCR.iTipoOrigem = iTipo
    objItensPRJCR.sObservacao = sObs

    colItensPRJ.Add objItensPRJCR

    Exit Function

Erro_Preenche_Itens_Projeto_Aux:

    Preenche_Itens_Projeto_Aux = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
    
End Function

Public Sub BotaoInfoAdic_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoInfoAdic_Click

    If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic

    Call Chama_Tela_Modal("RetiradaEntrega", gobjInfoAdic, Me)

    Exit Sub

Erro_BotaoInfoAdic_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207733)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoCompra", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 26482

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CStr(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next
    
    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 26482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157915)

    End Select

    Exit Function

End Function

Public Sub TabelaPreco_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    If TabelaPreco.ListIndex = -1 Then Exit Sub

    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 30525

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 30525

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157964)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26538

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 26539

        If lErro <> SUCESSO Then gError 26540 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 30527

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 26541

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 26538, 26539, 30527

    Case 26540  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        Else
            'Segura o foco

        End If

    Case 26541

        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else

        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157979)

    End Select

    Exit Sub

End Sub

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco

    If Not gbCarregandoTela Then
    
        For iLinha = 1 To objGridItens.iLinhasExistentes
    
            lErro = Trata_TabelaPreco_Item(iLinha)
            If lErro <> SUCESSO Then gError 25437
    
        Next
    
        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51046

    End If
    
    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 25437 'tratado na rotina chamada

        Case 51046

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158052)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer, iTabela As Integer
Dim iItem As Integer

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 39147

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        
        iTabela = Codigo_Extrai(TabelaPreco.Text)
        
        objTabelaPrecoItem.iCodTabela = iTabela
        
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 39148

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 39149

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(iLinha, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 39150
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, PrecoUnitario.Format)
            Else
                GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = ""
            End If
            
            Call PrecoTotal_Calcula(iLinha)
            
        End If

    End If
    
    If Len(Trim(Taxa.Text)) > 0 Then Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Text))

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 39147, 39148, 39149, 39150, 132828 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158053)

    End Select

    Exit Function

End Function

Private Function PrecoUnitario_Calcula(ByVal iLinha As Integer, ByVal sUM As String, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long, objTabelaPreco As New ClassTabelaPreco
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26638

    If lErro = 28030 Then gError 26639
    'Converte a quantidade para a UM de Compra
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMCompra, dFator)
    If lErro <> SUCESSO Then gError 26640

    dPrecoUnitario = objTabelaPrecoItem.dPreco * dFator

    If objTabelaPrecoItem.iCodTabela <> 0 Then
    
        objTabelaPreco.iCodigo = objTabelaPrecoItem.iCodTabela
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 26638
        If lErro = SUCESSO And objTabelaPreco.iMoeda <> MOEDA_REAL And Codigo_Extrai(Moeda.Text) = MOEDA_REAL Then
            
            objCotacao.dtData = StrParaDate(Data.Text)
            objCotacao.iMoeda = objTabelaPreco.iMoeda
            objCotacaoAnterior.iMoeda = objTabelaPreco.iMoeda
        
            'Chama função de leitura
            lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
            If lErro <> SUCESSO Then gError 26638
            
            'Se nao existe cotacao para a data informada => Mostra a última.
            dPrecoUnitario = dPrecoUnitario * StrParaDbl(IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000")))
        
        End If
    
    End If
    
    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 26638, 26640

        Case 26639
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157981)

    End Select

    Exit Function

End Function

Public Function BotaoEntrega_Click() As Long

Dim lErro As Long
Dim sProdutoTela As String
Dim dQuantidade As Double
Dim objItemPC As ClassItemPedCompra
Dim sProduto As String, iPreenchido As Integer

On Error GoTo Erro_BotaoEntrega_Click

    If GridItens.Row = 0 Then gError 183202

    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))

    If Len(sProdutoTela) = 0 Then gError 183203
    
    Set objItemPC = gobjPC.colItens(GridItens.Row)
    
    gobjPC.lCodigo = StrParaLong(Codigo.Text)
    
    'Formata o produto
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objItemPC.sProduto = sProduto
    objItemPC.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    objItemPC.dQuantidade = dQuantidade

    Call Chama_Tela_Modal("DataEntregaCOM", gobjPC, objItemPC)
    
    If giRetornoTela = vbOK Then
        If objItemPC.colDataEntrega.Count > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_DataLimite_Col) = Format(objItemPC.colDataEntrega.Item(1).dtDataEntrega, "dd/mm/yyyy")
        End If
    End If
    
    BotaoEntrega_Click = SUCESSO

    Exit Function

Erro_BotaoEntrega_Click:

    BotaoEntrega_Click = gErr

    Select Case gErr

        Case 183202
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 183203
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183196)

    End Select

    Exit Function
    
End Function

Public Sub DeliveryDate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DeliveryDate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DeliveryDate_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DeliveryDate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DeliveryDate
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TempoTransito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TempoTransito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub TempoTransito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub TempoTransito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TempoTransito
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_DeliveryDate(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtData As Date
Dim dtDeliveryDate As Date
Dim dtDataLimite As Date
Dim iTempo As Integer

On Error GoTo Erro_Saida_Celula_DeliveryDate

    Set objGridInt.objControle = DeliveryDate

    'Verifica se a data limite foi preenchida
    If Len(Trim(DeliveryDate.ClipText)) > 0 Then

        'Critica a data limite
        lErro = Data_Critica(DeliveryDate.Text)
        If lErro <> SUCESSO Then gError 56078

        dtDeliveryDate = CDate(DeliveryDate.Text)

        'Verifica se a data do pedido foi preenchida
        If Len(Trim(Data.ClipText)) > 0 Then

            dtData = CDate(Data.Text)

            'comparacao entre data limite e data do pedido
            If dtDeliveryDate < dtData Then gError 56079

        End If
        
        iTempo = StrParaInt(GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_TempoTransito_Col))
        
        dtDataLimite = DateAdd("d", iTempo, dtDeliveryDate)
        
        GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_DataLimite_Col) = Format(dtDataLimite, "dd/mm/yyyy")
        
    Else
    
        GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_DataLimite_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56080

    Saida_Celula_DeliveryDate = SUCESSO

    Exit Function

Erro_Saida_Celula_DeliveryDate:

    Saida_Celula_DeliveryDate = gErr

    Select Case gErr

        Case 56078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56079
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INFERIOR_DATAPEDIDO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56080
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164571)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TempoTransito(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtData As Date
Dim dtDeliveryDate As Date
Dim dtDataLimite As Date
Dim iTempo As Integer

On Error GoTo Erro_Saida_Celula_TempoTransito

    Set objGridInt.objControle = TempoTransito

    dtDeliveryDate = StrParaDate(GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_DeliveryDate_Col))

    'Verifica se a data limite foi preenchida
    If Len(Trim(TempoTransito.ClipText)) > 0 Then

        'Critica a data limite
        'lErro = Inteiro_Critica(TempoTransito.Text)
        'If lErro <> SUCESSO Then gError 56078

        iTempo = StrParaInt(TempoTransito.Text)
        
        
        dtDataLimite = DateAdd("d", iTempo, dtDeliveryDate)
        
        If dtDeliveryDate <> DATA_NULA Then GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_DataLimite_Col) = Format(dtDataLimite, "dd/mm/yyyy")
        
    Else
    
        If dtDeliveryDate <> DATA_NULA Then GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_DataLimite_Col) = GridItens.TextMatrix(objGridInt.objGrid.Row, iGrid_DeliveryDate_Col)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56080

    Saida_Celula_TempoTransito = SUCESSO

    Exit Function

Erro_Saida_Celula_TempoTransito:

    Saida_Celula_TempoTransito = gErr

    Select Case gErr

        Case 56078
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56079
            Call Rotina_Erro(vbOKOnly, "ERRO_DATALIMITE_INFERIOR_DATAPEDIDO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56080
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164571)

    End Select

    Exit Function

End Function

