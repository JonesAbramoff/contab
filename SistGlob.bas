Attribute VB_Name = "SistGlob"
Public SistGlob_iDebug As Integer
Public SistGlob_lUltimoErro As Long
Public SistGlob_lSistema As Long
Public SistGlob_lAcesso As Long
Public SistGlob_lConexao As Long
Public SistGlob_lConexaoDic As Long
Public SistGlob_lConexaoRel As Long
Public SistGlob_lConexaoBrowse As Long
Public SistGlob_lConexaoDicBrowse As Long
Public SistGlob_lTransacao As Long
Public SistGlob_lTransacaoDic As Long
Public SistGlob_lMascara As Long
Public SistGlob_objMascaraGenerica As Object
Public SistGlob_lMascProd As Long
Public SistGlob_lDicDados As Long
Public SistGlob_lEmpresa As Long
Public SistGlob_iGrupoEmpresarial As Integer
Public SistGlob_sNomeEmpresa As String
Public SistGlob_sNomePrinc As String
Public SistGlob_sUsuario As String
Public SistGlob_iFilialEmpresa As Integer
Public SistGlob_iDesconsideraFechamentoPeriodo As Integer
Public SistGlob_sModulo As String
Public SistGlob_sNomeFilialEmpresa As String
Public SistGlob_objContabInt As Object
Public SistGlob_objKeepAlive As AdmKeepAlive
Public SistGlob_dtDataAtual As Date
Public SistGlob_dtDataHoje As Date
Public SistGlob_objCheckboxChecked As Picture
Public SistGlob_objCheckboxUnchecked As Picture
Public SistGlob_objCheckboxGrayed As Picture
Public SistGlob_objOptionButtonChecked As Picture
Public SistGlob_objOptionButtonUnchecked As Picture
Public SistGlob_objButton As Picture
Public SistGlob_objAdmColModulo As AdmColModulo
Public SistGlob_objMDIForm As Object
Public SistGlob_colCampos As Collection
Public SistGlob_colTiposMovEst As Collection
Public SistGlob_colErrosBatch As Collection
Public SistGlob_lErro As Long
Public SistGlob_iTipoVersao As Long
Public SistGlob_objEstInicial As Object
Public SistGlob_lpPrevWndProc As Long
Public SistGlob_lpPrevWndProc0 As Long
Public SistGlob_lpPrevWndProc00 As Long
Public SistGlob_lpPrevWndProc1 As Long
Public SistGlob_lpPrevWndProc2 As Long
Public SistGlob_lpPrevWndProc3 As Long
Public SistGlob_lpPrevWndProc4 As Long
Public SistGlob_lpPrevWndProc5 As Long
Public SistGlob_lpPrevWndProc6 As Long
Public SistGlob_lpPrevWndProc7 As Long
Public SistGlob_lpPrevWndProc8 As Long
Public SistGlob_lpPrevWndProc9 As Long
Public SistGlob_lpPrevWndProc10 As Long
Public SistGlob_lpPrevWndProc11 As Long
Public SistGlob_lpPrevWndProc12 As Long
Public SistGlob_lpPrevWndProc13 As Long
Public SistGlob_lpPrevWndProc14 As Long
Public SistGlob_lpPrevWndProc15 As Long
Public SistGlob_objTelaAtiva As Object
Public SistGlob_objControleDrag As Object
Public SistGlob_objControleAlvo As Object
Public SistGlob_objmenuEdicao As Menu
Public SistGlob_colWndProc As New Collection
Public SistGlob_sngEdicaoX As Single
Public SistGlob_sngEdicaoY As Single
Public SistGlob_sngDragX As Single
Public SistGlob_sngDragY As Single
Public SistGlob_iLeft As Integer
Public SistGlob_iTop As Integer
Public SistGlob_objPropriedades As Object
Public SistGlob_objCamposInvisiveis As Object
Public SistGlob_iProxMouseMove As Integer
Public SistGlob_iProxButtonUp As Integer
'Edicao Tela - Raphael
Public SistGlob_colEdicaoTela As Collection

Public SistGlob_iSQLTipoOrdParam As Integer
Public SistGlob_iSQLTipoOrdParamDic As Integer
Public SistGlob_iLocalOperacao As Integer
Public SistGlob_iContabGerencial As Integer
Public SistGlob_iFilialAuxiliar As Integer
Public SistGlob_iCliAtrasoDestacar As Integer

'Incluído em 02/07/2001 por Luiz Gustavo de Freitas Nogueira
Public SistGlob_objExcel As Object

Public SistGlob_STRING_ENDERECO As Integer
Public SistGlob_STRING_BAIRRO As Integer
Public SistGlob_STRING_CIDADE As Integer

Public SistGlob_STRING_TELEFONE As Integer
Public SistGlob_STRING_FAX As Integer
Public SistGlob_STRING_EMAIL As Integer
Public SistGlob_STRING_CONTATO As Integer

Public SistGlob_STRING_CLIENTE_RAZAO_SOCIAL As Integer
Public SistGlob_STRING_CLIENTE_NOME_REDUZIDO As Integer
Public SistGlob_STRING_CLIENTE_OBSERVACAO As Integer

Public SistGlob_STRING_FORNECEDOR_RAZAO_SOC As Integer
Public SistGlob_STRING_FORNECEDOR_NOME_REDUZIDO As Integer

Public SistGlob_STRING_TRANSPORTADORA_NOME As Integer
Public SistGlob_STRING_TRANSPORTADORA_NOME_REDUZIDO As Integer

Public SistGlob_iBrowsePosicaoAntigo As Integer

Public SistGlob_iForcaSistemaIntegrado As Integer
Public SistGlob_iSistemaIntegradoForcado As Integer

Public SistGlob_NUM_MAX_ITENS_REQUISICAO As Integer
Public SistGlob_NUM_MAX_ITENS_PEDIDO_COTACAO As Integer
Public SistGlob_NUM_MAX_ITENS_PEDIDO_COMPRAS As Integer
Public SistGlob_NUM_MAX_ITENS_DISTRIBUICAO As Integer
Public SistGlob_NUM_MAX_ITENS_GERACAO As Integer
Public SistGlob_NUM_MAX_PRODUTOS_COTACAO As Integer
Public SistGlob_NUM_MAX_FORNECEDORES_COTACAO As Integer
Public SistGlob_NUM_MAX_NFS_ITEMREQ As Integer
Public SistGlob_NUM_MAX_PEDIDOS_ITEMREQ As Integer
Public SistGlob_NUM_MAX_COTACOES As Integer
Public SistGlob_NUM_MAX_NFS_ITEMPED As Integer
Public SistGlob_NUM_MAX_ITENS_MOV_ESTOQUE As Integer

Public SistGlob_STRING_PRODUTO_NOME_REDUZIDO As Integer
Public SistGlob_STRING_PRODUTO_REFERENCIA As Integer
Public SistGlob_STRING_PRODUTO_DESCRICAO_TELA As Integer
Public SistGlob_STRING_PRODUTO_MODELO As Integer

Public SistGlob_STRING_ORDEM_DE_PRODUCAO As Integer
Public SistGlob_STRING_LOTE_RASTREAMENTO As Integer

Public SistGlob_iTelaTamanhoVariavel As Integer
Public SistGlob_sExtensaoGerRelExp As String

Public SistGlob_lADMCount As Long
Public SistGlob_bVPN As Boolean
Public SistGlob_bPreLoadGravar As Boolean
Public SistGlob_colRotinasSGE As Object
Public SistGlob_bTelaReordenando As Boolean

'p/depuracao de comandos SQL
Public glCom(1 To 2000) As Long
Public giNumCom As Integer
Public gsCom(1 To 2000) As String

Public giNumCallStack As Integer
Public gsCallStack(1 To 1000) As String
Public gsLogonId As String
