VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function BloqueioEstoquePV_Libera(ByVal objBloqueioPV As ClassBloqueioPV) As Long
'Libera o Bloqueio de Pedido de Venda passado em objBloqueioPV

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_BloqueioEstoquePV_Libera

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46037

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 46041

    'Lê o bloqueio de Pedido de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND Sequencial = ? AND (TipoDeBloqueio = ? OR TipoDeBloqueio = ?) AND DataLib = ?", 0, iFilialEmpresa, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial, BLOQUEIO_PARCIAL, BLOQUEIO_NAO_RESERVA, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then Error 46038

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46039

    If lErro = AD_SQL_SUCESSO Then

        'Data e responsável pela liberação do Bloqueio
        objBloqueioPV.dtDataLib = gdtDataHoje
        objBloqueioPV.sResponsavelLib = BLOQUEIO_AUTO_RESP
        objBloqueioPV.sCodUsuarioLib = gsUsuario

        'Libera o bloqueio
        lErro = Comando_ExecutarPos(lComando1, "UPDATE BloqueiosPV SET DataLib = ?, ResponsavelLib = ?, CodUsuarioLib = ? ", lComando, objBloqueioPV.dtDataLib, objBloqueioPV.sResponsavelLib, objBloqueioPV.sCodUsuarioLib)
        If lErro <> AD_SQL_SUCESSO Then Error 46040
    End If

    'Finaliza os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    BloqueioEstoquePV_Libera = SUCESSO

    Exit Function

Erro_BloqueioEstoquePV_Libera:

    BloqueioEstoquePV_Libera = Err

    Select Case Err

        Case 46037, 46041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46038, 46039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOS_PV_LIBERACAO", Err)

        Case 46040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIOPV", Err, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152501)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Reserva_Automatico3(ByVal lComando As Long, ByVal lComando2 As Long, ByVal lCodigo As Long) As Long
'grava o proximo codigo de reserva disponivel

Dim lErro As Long

On Error GoTo Erro_Reserva_Automatico3

    'atualiza o numero da Reserva gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE MATConfig SET Conteudo = ?", lComando, CStr(lCodigo + 1))
    If lErro <> AD_SQL_SUCESSO Then Error 36678

    Reserva_Automatico3 = SUCESSO

    Exit Function

Erro_Reserva_Automatico3:

    Reserva_Automatico3 = Err

    Select Case Err

        Case 36678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CPRCONFIG", Err, CStr(lCodigo + 1))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152502)

    End Select

    Exit Function

End Function

Function Reserva_Insere(ByVal lComando As Long, ByVal lComando1 As Long, ByVal objItemPedido As ClassItemPedido, ByVal dQuantidadeReserva As Double, ByVal lCodigoReserva As Long, ByVal iAlmoxarifado As Integer) As Long
'Insere uma reserva a partir dos dados passados como parametro
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim objReserva As New ClassReserva
Dim lErro As Long, lNumIntDoc As Long

On Error GoTo Erro_Reserva_Insere

    '??? o codigo passado já foi obtido anteriormente, só vai dar erro se usuario entrar codigo manualmente
'    'descobre o proximo codigo de reserva disponivel
'    lErro = CF("Reserva_Automatico2",lComando1, lCodigoReserva)
'    If lErro <> SUCESSO Then Error 36685

    'descobre o proximo NumIntDoc de reserva disponivel
    lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_RESERVA", lNumIntDoc)
    If lErro <> SUCESSO Then Error 41603

    'insere a Reserva
    lErro = Comando_Executar(lComando, "INSERT INTO Reserva(FilialEmpresa,Codigo,Produto,Almoxarifado,TipoDoc,DocOrigem,NumIntOrigem,Quantidade,DataValidade,DataReserva,CodUsuario, Responsavel,NumIntDoc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemPedido.iFilialEmpresa, lCodigoReserva, objItemPedido.sProduto, iAlmoxarifado, TIPO_PEDIDO_COD, objItemPedido.lCodPedido, objItemPedido.lNumIntDoc, dQuantidadeReserva, DATA_NULA, gdtDataAtual, gsUsuario, "", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 36686
    
    Reserva_Insere = SUCESSO

    Exit Function

Erro_Reserva_Insere:

    Reserva_Insere = Err
    
    Select Case Err

        Case 36685, 36686
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152503)

    End Select

    Exit Function

End Function


'Function PedidoDeVenda_Baixar(objPedidoVenda As ClassPedidoVenda) As Long
'
'Dim lErro As Long, objItemPV As ClassItemPedido
'On Error GoTo Erro_PedidoDeVenda_Baixar
'
'    'insere o pedido na tabela de Pedidos de Venda Baixados
'
'    'percorre os itens do pedido
'    For Each objItemPV In objPedidoVenda.colItensPedido
'
'        'le registro de itempv
'        'exclui registro
'
'        'insere registro em itens de pv baixados
'
'    Next
'
'    'exclui registro na tabela de PV
'
'    PedidoDeVenda_Baixar = SUCESSO
'
'    Exit Function
'
'Erro_PedidoDeVenda_Baixar:
'
'    PedidoDeVenda_Baixar = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152504)
'
'    End Select
'
'    Exit Function
'
'End Function


Friend Function Verifica_EstoqueMaximo(ByVal objItemOP As ClassItemOP) As Long
'Verifica se quantidade ordenada + disponivel > quantidade máxima de estoque
Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim sComando_SQL As String
Dim dEstoqueMaximo As Double
Dim dTotal As Double

On Error GoTo Erro_Verifica_EstoqueMaximo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41331
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 41332

    '??? numa versao melhorada poderemos incluir no teste as qtdes já em OPs nao baixadas e qtdes em requisicao de compra aberta.
    
    sComando_SQL = "SELECT SUM (EstoqueProduto.QuantDispNossa + EstoqueProduto.QuantConsig3 + EstoqueProduto.QuantReservada + EstoqueProduto.QuantReservadaConsig) FROM EstoqueProduto, Almoxarifado WHERE (EstoqueProduto.Produto = ?) AND (Almoxarifado.FilialEmpresa = ?) AND (Almoxarifado.Codigo = EstoqueProduto.Almoxarifado)"
    lErro = Comando_Executar(lComando, sComando_SQL, dTotal, objItemOP.sProduto, objItemOP.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41645
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41646
    
    If lErro = AD_SQL_SEM_DADOS Then Error 41647
    
   'Lê de ProdutosFilial a Quantidade máxima de estoque desse produto
    sComando_SQL = "SELECT EstoqueMaximo FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ?"
    lErro = Comando_Executar(lComando1, sComando_SQL, dEstoqueMaximo, objItemOP.iFilialEmpresa, objItemOP.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 41333

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41334

    If lErro = AD_SQL_SUCESSO Then
            
        If dEstoqueMaximo > 0 Then
            If objItemOP.dQuantidade + dTotal > dEstoqueMaximo Then Error 41335
        End If
        
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Verifica_EstoqueMaximo = SUCESSO

    Exit Function

Erro_Verifica_EstoqueMaximo:

    Verifica_EstoqueMaximo = Err

    Select Case Err

        Case 41331, 41332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41333, 41334, 41645, 41646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err)
            
        Case 41647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DETERMINACAO_QUANTMAX", Err, objItemOP.iFilialEmpresa)
        
        Case 41335

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152505)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

''Function OrdemDeProducao_Grava(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'''cria ou atualiza uma ordem de producao
''
''Dim lErro As Long
''Dim iIndice As Integer, lNumIntDocItem As Long, lEmpenho As Long
''Dim sDescricao As String
''Dim dtDataEmissao As Date
''Dim iNumItens As Integer
''Dim iNumItensBaixados As Integer
''Dim objItemOp As New ClassItemOP
''Dim alComando(16) As Long
''Dim lTransacao As Long
''Dim vbMsgRes As VbMsgBoxResult
''
''On Error GoTo Erro_OrdemDeProducao_Grava
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''        'Abertura comandos
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 31521
''    Next
''
''    'Abertura transação
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then Error 31522
''
''    'Leitura da tabela de Ordens de Produção Baixadas
''    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumItens FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ?", dtDataEmissao, iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 31523
''
''    lErro = Comando_BuscarPrimeiro(alComando(0))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31524
''
''    'se já existe uma OP baixada com este codigo => erro
''    If lErro = AD_SQL_SUCESSO Then Error 31525
''
''    'Leitura da tabela de Ordens de Produção
''    lErro = Comando_ExecutarPos(alComando(1), "SELECT DataEmissao, NumItens, NumItensBaixados FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, dtDataEmissao, iNumItens, iNumItensBaixados, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 31526
''
''    lErro = Comando_BuscarPrimeiro(alComando(1))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31527
''
''    'Existe a OP no BD
''    If lErro = AD_SQL_SUCESSO Then
''
''        'Lock na OP
''        lErro = Comando_LockExclusive(alComando(1))
''        If lErro <> AD_SQL_SUCESSO Then Error 22968
''
''        'atualiza os itens pré-existentes da OP
''        lErro = OrdemDeProducao_Grava4(objOrdemDeProducao)
''        If lErro <> SUCESSO Then Error 22969
''
''        'se trocou a DataEmissao, NumItens ou NumItensBaixados
''        If objOrdemDeProducao.dtDataEmissao <> dtDataEmissao Or objOrdemDeProducao.iNumItens <> iNumItens Or objOrdemDeProducao.iNumItensBaixados <> iNumItensBaixados Then
''
''            lErro = Comando_ExecutarPos(alComando(2), "UPDATE OrdensDeProducao SET DataEmissao = ?, NumItens = ?, NumItensBaixados = ?", alComando(1), objOrdemDeProducao.dtDataEmissao, objOrdemDeProducao.iNumItens, objOrdemDeProducao.iNumItensBaixados)
''            If lErro <> AD_SQL_SUCESSO Then Error 22970
''
''        End If
''
''    Else 'Nao existe a OP no BD
''
''        'Insere nova OP
''        lErro = Comando_Executar(alComando(2), "INSERT INTO OrdensDeProducao (Codigo, FilialEmpresa, DataEmissao, NumItens, NumItensBaixados) VALUES (?,?,?,?,?) ", objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.dtDataEmissao, objOrdemDeProducao.colItens.Count, 0)
''        If lErro <> AD_SQL_SUCESSO Then Error 31528
''
''    End If
''
''    'obtem o proximo codigo para empenhos e para NumIntItemOP
''    lErro = OrdemDeProducao_Grava3(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
''    If lErro <> AD_SQL_SUCESSO Then Error 22962
''
''    'inclui os novos itens da OP
''    For Each objItemOp In objOrdemDeProducao.colItens
''
''         If objItemOp.lNumIntDoc = 0 Then
''
''            'Verifica o estoque máximo do produto utilizado
''            lErro = Verifica_EstoqueMaximo(objItemOp)
''            If lErro <> SUCESSO And lErro <> 41335 Then Error 40663
''
''            If lErro = 41335 Then
''
''                'Confirmar se vai gravar com a Quantidade maior que a quantidade maxima do estoque
''                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ESTOQUEMAXIMO", objItemOp.sProduto)
''
''                If vbMsgRes = vbNo Then Error 40622
''
''            End If
''
''            objItemOp.lNumIntDoc = lNumIntDocItem
''
''            lErro = OrdemDeProducao_Grava1(objItemOp, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
''            If lErro <> SUCESSO Then Error 31531
''
''            lNumIntDocItem = lNumIntDocItem + 1
''
''        End If
''
''    Next
''
''    'atualiza o proximo codigo para lNumIntDocItem e lEmpenho
''    lErro = OrdemDeProducao_Grava5(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
''    If lErro <> AD_SQL_SUCESSO Then Error 19471
''
''    If objOrdemDeProducao.iNumItens = objOrdemDeProducao.iNumItensBaixados Then
''
''        'Baixa a ordem de producao, seus itens e seus empenhos
''        lErro = CF("OrdemDeProducao_Baixa",objOrdemDeProducao.sCodigo)
''        If lErro <> SUCESSO Then Error 55273
''
''    End If
''
''    'Confirma transação
''    lErro = Transacao_Commit()
''    If lErro <> AD_SQL_SUCESSO Then Error 31532
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    OrdemDeProducao_Grava = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava:
''
''    OrdemDeProducao_Grava = Err
''
''    Select Case Err
''
''        Case 22968
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORDENSDEPRODUCAO", Err)
''
''        Case 22970
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ORDENSDEPRODUCAO", Err, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
''
''        Case 27509, 27510
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", Err)
''
''        Case 31521
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 31522
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
''
''        Case 31523, 31524
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS", Err)
''
''        Case 31525
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", Err, objOrdemDeProducao.sCodigo)
''
''        Case 31526, 31527
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", Err)
''
''        Case 31528
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDEPRODUCAO", Err)
''
''        Case 31530
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG", Err)
''
''        Case 19471, 22962, 22969, 31531, 40622, 40663, 55273
''
''        Case 31532
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152506)
''    End Select
''
''    'Fechamento transação
''    Call Transacao_Rollback
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function

''Private Function OrdemDeProducao_Grava1(objItemOp As ClassItemOP, alComando() As Long, dtDataEmissao As Date, lEmpenho As Long) As Long
'''Inclui um item numa OP
'''Tem que ser chamada dentro de transação
''
''Dim lErro As Long
''Dim iPCP As Integer
''Dim iCodigo As Integer
''Dim iItem As Integer
''Dim iControleEstoque As Integer
''Dim iNivel As Integer
''Dim objProdutoKit As New ClassProdutoKit
''Dim colProdutos As New Collection
''Dim dQuantidadeRaiz As Double
''Dim objItemPV As New ClassItemPedido
''Dim sCcl As String, objEstoqueProduto As New ClassEstoqueProduto
''Dim iClasseUM As Integer
''Dim sSiglaUMEstoque As String
''Dim iCompras As Integer
''
''On Error GoTo Erro_OrdemDeProducao_Grava1
''
''    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
''
''    'Leitura da tabela Produtos
''    lErro = Comando_ExecutarLockado(alComando(4), "SELECT PCP, SiglaUMEstoque, ClasseUM, Compras FROM Produtos WHERE Codigo = ? ", iPCP, sSiglaUMEstoque, iClasseUM, iCompras, objItemOp.sProduto)
''    If lErro <> AD_SQL_SUCESSO Then Error 31533
''
''    lErro = Comando_BuscarPrimeiro(alComando(4))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31534
''
''    If lErro = AD_SQL_SEM_DADOS Then Error 31535
''
''    lErro = Comando_LockShared(alComando(4))
''    If lErro <> AD_SQL_SUCESSO Then Error 31536
''
''    If iPCP = PRODUTO_PCP_NAOPODE Or iCompras <> PRODUTO_PRODUZIVEL Then Error 55277
''
''    objItemOp.sSiglaUMEstoque = sSiglaUMEstoque
''    objItemOp.iClasseUM = iClasseUM
''
''    'garantir existencia do almoxarifado
''    lErro = Comando_ExecutarLockado(alComando(5), "SELECT Codigo FROM Almoxarifado WHERE Codigo = ? ", iCodigo, objItemOp.iAlmoxarifado)
''    If lErro <> AD_SQL_SUCESSO Then Error 31537
''
''    lErro = Comando_BuscarPrimeiro(alComando(5))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31538
''
''    If lErro = AD_SQL_SEM_DADOS Then Error 31539
''
''    lErro = Comando_LockShared(alComando(5))
''    If lErro <> AD_SQL_SUCESSO Then Error 31540
''
''    'garantir a existencia da Ccl
''    If objItemOp.sCcl <> "" Then
''
''        sCcl = String(STRING_CCL, 0)
''
''        lErro = Comando_ExecutarLockado(alComando(6), "SELECT Ccl FROM Ccl WHERE Ccl = ? ", sCcl, objItemOp.sCcl)
''        If lErro <> AD_SQL_SUCESSO Then Error 31541
''
''        lErro = Comando_BuscarPrimeiro(alComando(6))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31542
''
''        If lErro = AD_SQL_SEM_DADOS Then Error 31543
''
''        lErro = Comando_LockShared(alComando(6))
''        If lErro <> AD_SQL_SUCESSO Then Error 31544
''
''    End If
''
''    If objItemOp.lCodPedido <> 0 Then
''
''        objItemPV.iFilialEmpresa = objItemOp.iFilialPedido
''        objItemPV.lCodPedido = objItemOp.lCodPedido
''        objItemPV.sProduto = objItemOp.sProduto
''
''        lErro = OrdemDeProducao_Grava6(alComando(14), alComando(15), objItemPV, objItemOp.dQuantidade)
''        If lErro <> SUCESSO And lErro <> 41346 Then Error 40664
''
''        If lErro = 41346 Then Error 40666
''
''        objItemOp.lNumIntOrigem = objItemPV.lNumIntDoc
''
''    End If
''
''    'Insere item da OP
''    lErro = Comando_Executar(alComando(7), "INSERT INTO ItensOrdemProducao (NumIntDoc, FilialEmpresa, Codigo, Item, Produto, FilialPedido, " & _
''    "CodPedido, NumIntOrigem, Almoxarifado, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, " & _
''    "Situacao, Destinacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
''    objItemOp.lNumIntDoc, objItemOp.iFilialEmpresa, objItemOp.sCodigo, objItemOp.iItem, objItemOp.sProduto, objItemOp.iFilialPedido, objItemOp.lCodPedido, objItemOp.lNumIntOrigem, _
''    objItemOp.iAlmoxarifado, objItemOp.sCcl, objItemOp.sSiglaUM, objItemOp.dQuantidade, objItemOp.dtDataInicioProd, objItemOp.dtDataFimProd, objItemOp.iPrioridade, _
''    objItemOp.dQuantidadeProd, objItemOp.iSituacao, objItemOp.iDestinacao)
''    If lErro <> SUCESSO Then Error 31551
''
''    'cria os empenhos
''    lErro = CF("OrdemDeProducao_Grava2",objItemOp, alComando, dtDataEmissao, lEmpenho)
''    If lErro <> SUCESSO Then Error 31552
''
''    'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
''    objEstoqueProduto.sProduto = objItemOp.sProduto
''    objEstoqueProduto.iAlmoxarifado = objItemOp.iAlmoxarifado
''
''    lErro = CF("EstoqueProduto_AlterarQtdOP",objEstoqueProduto, objItemOp.dQuantidade)
''    If lErro <> SUCESSO Then Error 41552
''
''    OrdemDeProducao_Grava1 = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava1:
''
''    OrdemDeProducao_Grava1 = Err
''
''    Select Case Err
''
''        Case 31533, 31534
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err)
''
''        Case 31535
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objItemOp.sProduto)
''
''        Case 31536
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", Err)
''
''        Case 31537, 31538
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", Err)
''
''        Case 31539
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objItemOp.iAlmoxarifado)
''
''        Case 31540
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO", Err)
''
''        Case 31541, 31542
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err)
''
''        Case 31543
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objItemOp.iAlmoxarifado)
''
''        Case 31544
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", Err)
''
''        Case 31549, 31550
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", Err)
''
''       Case 31551
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSORDENSDEPRODUCAO", Err)
''
''        Case 31552, 40664
''
''        Case 40666
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", Err, objItemOp.lCodPedido, objItemOp.sProduto)
''
''        Case 55277
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", Err, objItemOp.sProduto)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152507)
''
''    End Select
''
''    Exit Function
''
''End Function

''Private Function OrdemDeProducao_Grava2(objItemOp As ClassItemOP, alComando() As Long, dtDataEmissao As Date, lEmpenho As Long) As Long
'''cria empenhos correspondentes ao objItemOP
'''IMPORTANTE: Tem que ser chamada dentro de transação
''
''Dim lErro As Long
''Dim iNivel As Integer
''Dim iClasseUM As Integer, iClasseUMRaiz As Integer, iAlmoxPadraoProd As Integer
''Dim iControleEstoque As Integer
''Dim dFator As Double, dFatorRaiz As Double, dFatorItemOP As Double
''Dim sVersao As String
''Dim sProduto As String
''Dim sSiglaUMEstoque As String, sSiglaUMEstoqueRaiz As String, sSiglaUMKit As String
''Dim objProdutoKit As New ClassProdutoKit
''Dim tProdutoKit As typeProdutoKit, objEmpenho As New ClassEmpenho
''Dim tProdutoKitRaiz As typeProdutoKit
''
''On Error GoTo Erro_OrdemDeProducao_Grava2
''
''    'procura a composicao do item de acordo com a data de emissao da op
''    sVersao = String(STRING_KIT_VERSAO, 0)
''
''    lErro = Comando_ExecutarLockado(alComando(8), "SELECT Versao FROM Kit WHERE ProdutoRaiz = ? AND Data <= ? ORDER BY Data DESC", sVersao, objItemOp.sProduto, dtDataEmissao)
''    If lErro <> AD_SQL_SUCESSO Then Error 31553
''
''    lErro = Comando_BuscarPrimeiro(alComando(8))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31554
''
''    If lErro = AD_SQL_SUCESSO Then
''
''        lErro = Comando_LockShared(alComando(8))
''        If lErro <> AD_SQL_SUCESSO Then Error 31555
''
''        sSiglaUMEstoqueRaiz = String(STRING_UM_SIGLA, 0)
''        tProdutoKitRaiz.sUnidadeMed = String(STRING_UM_SIGLA, 0)
''
''        'pega dados do raiz (nivel = 0)
''        lErro = Comando_Executar(alComando(10), "SELECT Quantidade, ProdutoKit.UnidadeMed, SiglaUMEstoque, ClasseUM FROM ProdutoKit, Produtos WHERE ProdutoKit.Produto = ? AND ProdutoKit.Versao = ? AND ProdutoKit.Nivel = ? AND ProdutoKit.Produto = Produtos.Codigo", tProdutoKitRaiz.dQuantidade, tProdutoKitRaiz.sUnidadeMed, sSiglaUMEstoqueRaiz, iClasseUMRaiz, objItemOp.sProduto, sVersao, 0)
''        If lErro <> AD_SQL_SUCESSO Then Error 31556
''
''        lErro = Comando_BuscarPrimeiro(alComando(10))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31557
''
''        'descobre o fator de conversao da UM p/UM de estoque do produto raiz
''        lErro = CF("UM_Conversao_Trans",iClasseUMRaiz, tProdutoKitRaiz.sUnidadeMed, sSiglaUMEstoqueRaiz, dFatorRaiz)
''        If lErro <> SUCESSO Then Error 55248
''
''        'descobre o fator de conversao da UM p/UM de estoque do produto ao qual o item da OP se refere
''        lErro = CF("UM_Conversao_Trans",objItemOp.iClasseUM, objItemOp.sSiglaUM, objItemOp.sSiglaUMEstoque, dFatorItemOP)
''        If lErro <> SUCESSO Then Error 55249
''
''        sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
''        tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA, 0)
''        tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
''
''        'para obter apenas os nós do nivel 1
''        lErro = Comando_Executar(alComando(9), "SELECT ProdutoKit.Produto, Quantidade, ProdutoKit.UnidadeMed, Composicao, ControleEstoque, SiglaUMEstoque, ClasseUM FROM ProdutoKit, Produtos WHERE ProdutoRaiz = ? AND Versao = ? AND ProdutoKit.Produto = Produtos.Codigo AND ProdutoKit.Nivel = 1", tProdutoKit.sProduto, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, iControleEstoque, sSiglaUMEstoque, iClasseUM, objItemOp.sProduto, sVersao)
''        If lErro <> AD_SQL_SUCESSO Then Error 31558
''
''        lErro = Comando_BuscarPrimeiro(alComando(9))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31559
''
''        Do While lErro <> AD_SQL_SEM_DADOS
''
''            If iControleEstoque = PRODUTO_CONTROLE_RESERVA Or iControleEstoque = PRODUTO_CONTROLE_ESTOQUE Then
''
''                'le a tabela produtofilial para tentar descobrir o almoxarifado padrao associado ao produto
''                lErro = Comando_ExecutarLockado(alComando(16), "SELECT Almoxarifado FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ? ", iAlmoxPadraoProd, objItemOp.iFilialEmpresa, tProdutoKit.sProduto)
''                If lErro <> AD_SQL_SUCESSO Then Error 55209
''
''                lErro = Comando_BuscarPrimeiro(alComando(16))
''                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55210
''
''                lErro = Comando_LockExclusive(alComando(16))
''                If lErro <> AD_SQL_SUCESSO Then Error 55211
''
''                'se achou o almoxarifado padrão do produto ==> cria o empenho neste almoxarifado, senao nao cria o empenho
''                If lErro = SUCESSO And iAlmoxPadraoProd <> 0 Then
''
''                    If tProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
''
''                        tProdutoKit.dQuantidade = tProdutoKit.dQuantidade * ((objItemOp.dQuantidade * dFatorItemOP) / (tProdutoKitRaiz.dQuantidade * dFatorRaiz))
''
''                    End If
''
''                    lErro = CF("UM_Conversao_Trans",iClasseUM, tProdutoKit.sUnidadeMed, sSiglaUMEstoque, dFator)
''                    If lErro <> SUCESSO Then Error 31564
''
''                    objEmpenho.dQuantidade = tProdutoKit.dQuantidade * dFator
''                    objEmpenho.dtData = dtDataEmissao
''                    objEmpenho.iAlmoxarifado = iAlmoxPadraoProd
''                    objEmpenho.iFilialEmpresa = objItemOp.iFilialEmpresa
''                    objEmpenho.iItemOP = objItemOp.iItem
''                    objEmpenho.lCodigo = lEmpenho
''                    objEmpenho.lNumIntDocItemOP = objItemOp.lNumIntDoc
''                    objEmpenho.sCodigoOP = objItemOp.sCodigo
''                    objEmpenho.sProduto = tProdutoKit.sProduto
''                    objEmpenho.sUnidadeMed = sSiglaUMEstoque
''
''                    lEmpenho = lEmpenho + 1
''
''                    lErro = CF("Empenho_Inserir_Batch",objEmpenho, alComando(11))
''                    If lErro <> SUCESSO Then Error 31565
''
''                End If
''
''            End If
''
''            lErro = Comando_BuscarProximo(alComando(9))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31567
''
''        Loop
''
''        lErro = Comando_Unlock(alComando(8))
''        If lErro <> SUCESSO Then Error 31568
''
''    End If
''
''    OrdemDeProducao_Grava2 = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava2:
''
''    OrdemDeProducao_Grava2 = Err
''
''    Select Case Err
''
''        Case 31553, 31554
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", Err)
''
''        Case 31556, 31557, 31558, 31559
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", Err)
''
''        Case 31564, 31565, 55248, 55249
''
''        Case 55209, 55210
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, iAlmoxPadraoProd, objItemOp.sProduto)
''
''        Case 55211
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOSFILIAL", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152508)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function OrdemDeProducao_Exclui(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'''Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao
''
''Dim lErro As Long
''Dim iItem As Integer
''Dim iIndice As Integer
''Dim iSituacao As Integer, sProduto As String, iAlmoxarifado As Integer, dQtde As Double
''Dim iNumItens As Integer
''Dim lNumIntDoc As Long, lNumInt As Long
''Dim alComando(11) As Long
''Dim lTransacao As Long
''Dim objItemOp As New ClassItemOP, objEstoqueProduto As New ClassEstoqueProduto
''Dim lCodPedido As Long
''Dim objItemPV As New ClassItemPedido
''
''On Error GoTo Erro_OrdemDeProducao_Exclui
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''
''        'Abertura comando
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 21932
''
''    Next
''
''    'Abertura transação
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then Error 21933
''
''    'tenta ler se existe a OP
''    lErro = Comando_ExecutarPos(alComando(3), "SELECT NumItens FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 21934
''
''    lErro = Comando_BuscarPrimeiro(alComando(3))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21935
''
''    If lErro = AD_SQL_SEM_DADOS Then Error 21936
''
''    lErro = Comando_LockExclusive(alComando(3))
''    If lErro <> AD_SQL_SUCESSO Then Error 21937
''
''    'seleciona os itens da OP
''    sProduto = String(STRING_PRODUTO, 0)
''
''    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Item, Situacao, Produto, Quantidade, Almoxarifado, CodPedido FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, iItem, iSituacao, sProduto, dQtde, iAlmoxarifado, lCodPedido, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 21938
''
''    lErro = Comando_BuscarPrimeiro(alComando(1))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21939
''
''    'percorre os itens
''    Do While lErro <> AD_SQL_SEM_DADOS
''
''        If iSituacao <> ITEMOP_SITUACAO_NORMAL And iSituacao <> ITEMOP_SITUACAO_DESAB Then Error 21941
''
''        'pesquisa movimento de estoque associado ao item da OP
''        lErro = Comando_Executar(alComando(5), "SELECT NumIntDoc FROM MovimentoEstoque WHERE FilialEmpresa = ? AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", lNumInt, objOrdemDeProducao.iFilialEmpresa, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDoc)
''        If lErro <> AD_SQL_SUCESSO Then Error 21942
''
''        lErro = Comando_BuscarPrimeiro(alComando(5))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21943
''
''        If lErro = AD_SQL_SUCESSO Then Error 21944
''
''        objItemOp.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
''        objItemOp.lNumIntDoc = lNumIntDoc
''        objItemOp.sProduto = sProduto
''
''        'exclui empenhos associados ao itemOP
''        lErro = CF("Empenhos_Excluir_ItemOP",objItemOp)
''        If lErro <> SUCESSO Then Error 21945
''
''        'diminui qtde em op associado ao produto-almoxarifado
''        objEstoqueProduto.sProduto = sProduto
''        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
''
''        lErro = CF("EstoqueProduto_AlterarQtdOP",objEstoqueProduto, -dQtde)
''        If lErro <> SUCESSO Then Error 41553
''
''        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensOrdemProducao", alComando(1))
''        If lErro <> AD_SQL_SUCESSO Then Error 21946
''
''        'acerta a quantOP em Item de Pedido de Venda
''        If lCodPedido <> 0 Then
''
''            objItemPV.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
''            objItemPV.lCodPedido = lCodPedido
''            objItemPV.sProduto = objItemOp.sProduto
''
''            'diminui a quantidade em OP do Pedido de Venda
''            lErro = OrdemDeProducao_Grava6(alComando(10), alComando(11), objItemPV, -dQtde)
''            If lErro <> SUCESSO And lErro <> 41346 Then Error 55880
''
''            If lErro = 41346 Then Error 55881
''
''        End If
''
''        lErro = Comando_BuscarProximo(alComando(1))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21947
''
''    Loop
''
''    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM OrdensDeProducao", alComando(3))
''    If lErro <> AD_SQL_SUCESSO Then Error 21948
''
''    'Confirma transação
''    lErro = Transacao_Commit()
''    If lErro <> AD_SQL_SUCESSO Then Error 21949
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    OrdemDeProducao_Exclui = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Exclui:
''
''    OrdemDeProducao_Exclui = Err
''
''    Select Case Err
''
''        Case 21932
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 21933
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
''
''        Case 21934, 21935
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", Err)
''
''        Case 21938, 21939, 21947
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", Err)
''
''        Case 21936, 21945
''
''        Case 21941
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_SITUACAO_NAO_EXCLUIVEL", Err, iItem, objOrdemDeProducao.sCodigo)
''
''        Case 21942, 21943
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)
''
''        Case 21944
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_ASSOCIADO_MOVESTOQUE", Err, iItem, objOrdemDeProducao.sCodigo)
''
''        Case 21946
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSORDENSDEPRODUCAO", Err)
''
''        Case 21948
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORDENSDEPRODUCAO", Err)
''
''        Case 21949
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
''
''        Case 41553, 55880
''
''        Case 55881
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", Err, lCodPedido, objItemOp.sProduto)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152509)
''
''    End Select
''
''    'Fechamento transação
''    Call Transacao_Rollback
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
''
''
''Private Function OrdemDeProducao_Grava5(lComando As Long, lComando1 As Long, lNumIntDocItem As Long, lEmpenho As Long) As Long
'''atualiza numeração de lNumIntDocItem e lEmpenho
''
''Dim lErro As Long
''Dim alComando(0 To 1) As Long
''Dim iIndice As Integer
''
''On Error GoTo Erro_OrdemDeProducao_Grava5
''
''    'Abre novos comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 22943
''    Next
''
''    lErro = Comando_ExecutarPos(alComando(0), "UPDATE MATConfig SET Conteudo = ?", lComando, lNumIntDocItem)
''    If lErro <> AD_SQL_SUCESSO Then Error 21954
''
''    lErro = Comando_ExecutarPos(alComando(1), "UPDATE MATConfig SET Conteudo = ?", lComando1, lEmpenho)
''    If lErro <> AD_SQL_SUCESSO Then Error 22980
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    OrdemDeProducao_Grava5 = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava5:
''
''    OrdemDeProducao_Grava5 = Err
''
''    Select Case Err
''
''        Case 22943
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 22980
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", Err, "NUM_PROX_EMPENHO")
''
''        Case 21954
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", Err, "NUM_PROX_ITEM_OP")
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152510)
''
''    End Select
''
''    'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
''
''Private Function OrdemDeProducao_Grava3(lComando As Long, lComando1 As Long, lNumIntDocItem As Long, lEmpenho As Long) As Long
'''obtem o proximo codigo para lNumIntDocItem e lEmpenho
''
''Dim lErro As Long, sConteudo As String
''Dim iIndice As Integer
''
''On Error GoTo Erro_OrdemDeProducao_Grava3
''
''    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
''
''    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ?", 0, sConteudo, "NUM_PROX_ITEM_OP")
''    If lErro <> AD_SQL_SUCESSO Then Error 22955
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO Then Error 22956
''
''    lErro = Comando_LockExclusive(lComando)
''    If lErro <> AD_SQL_SUCESSO Then Error 22957
''
''    lNumIntDocItem = CLng(sConteudo)
''
''    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
''
''    lErro = Comando_ExecutarPos(lComando1, "SELECT Conteudo FROM MatConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, "NUM_PROX_EMPENHO", giFilialEmpresa)
''    If lErro <> AD_SQL_SUCESSO Then Error 22959
''
''    lErro = Comando_BuscarProximo(lComando1)
''    If lErro <> AD_SQL_SUCESSO Then Error 22960
''
''    lErro = Comando_LockExclusive(lComando1)
''    If lErro <> AD_SQL_SUCESSO Then Error 22961
''
''    lEmpenho = CLng(sConteudo)
''
''    OrdemDeProducao_Grava3 = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava3:
''
''    OrdemDeProducao_Grava3 = Err
''
''    Select Case Err
''
''        Case 22955, 22956
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", Err, "NUM_PROX_ITEM_OP")
''
''        Case 22957
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG", Err, "NUM_PROX_ITEM_OP")
''
''        Case 22959, 22960
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", Err, "NUM_PROX_EMPENHO")
''
''        Case 22961
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG", Err, "NUM_PROX_EMPENHO")
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152511)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function OrdemDeProducao_Grava4(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'''atualiza os itens pré-existentes da OP
''
''Dim alComando(1 To 11) As Long, iAlterar As Integer, iIndice As Integer
''Dim tItemOp As typeItemOP, objItemOp As ClassItemOP, objItemPV As New ClassItemPedido
''Dim iAchou As Integer, sCcl As String, lErro As Long
''Dim objEstoqueProduto As New ClassEstoqueProduto
''Dim sParam1 As String, sParam2 As String
''
''On Error GoTo Erro_OrdemDeProducao_Grava4
''
''    'Abre os comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 22971
''    Next
''
''    tItemOp.sCcl = String(STRING_CCL, 0)
''    tItemOp.sProduto = String(STRING_PRODUTO, 0)
''
''    'tenta selecionar item pre-existente da OP
''    lErro = Comando_ExecutarPos(alComando(1), "SELECT Produto, NumIntDoc, Item, FilialPedido, CodPedido, Almoxarifado, Ccl, Quantidade, DataInicioProd, DataFimProd, Prioridade, Situacao, Destinacao, QuantidadeProd FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Codigo = ? ORDER BY Item", 0, _
''            tItemOp.sProduto, tItemOp.lNumIntDoc, tItemOp.iItem, tItemOp.iFilialPedido, tItemOp.lCodPedido, tItemOp.iAlmoxarifado, tItemOp.sCcl, tItemOp.dQuantidade, tItemOp.dtDataInicioProd, tItemOp.dtDataFimProd, tItemOp.iPrioridade, tItemOp.iSituacao, tItemOp.iDestinacao, tItemOp.dQuantidadeProd, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 22972
''
''    lErro = Comando_BuscarPrimeiro(alComando(1))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22973
''
''    Do While lErro <> AD_SQL_SEM_DADOS
''
''        iAchou = 0
''
''        'localiza item na coleção
''        For Each objItemOp In objOrdemDeProducao.colItens
''            If objItemOp.iItem = tItemOp.iItem Then
''                iAchou = 1
''                objItemOp.lNumIntDoc = tItemOp.lNumIntDoc
''                Exit For
''            End If
''        Next
''
''        If iAchou = 0 Then Error 22975
''
''        'Quantidade (só pode mudar p/maior ou igual a produzida)
''        If objItemOp.dQuantidade < tItemOp.dQuantidadeProd Then Error 22976
''
''        iAlterar = 0
''
''        'se trocou FilialPedido ou CodPedido ou a quantidade
''        If objItemOp.iFilialPedido <> tItemOp.iFilialPedido Or objItemOp.lCodPedido <> tItemOp.lCodPedido Or objItemOp.dQuantidade <> tItemOp.dQuantidade Then
''
''            iAlterar = 1
''
''            'se trocou o pedido
''            If tItemOp.lCodPedido <> objItemOp.lCodPedido Then
''
''                If tItemOp.lCodPedido <> 0 Then
''
''                    'chamar funcao que faz update tirando a qtde anterior do pedido anterior
''                    objItemPV.iFilialEmpresa = tItemOp.iFilialPedido
''                    objItemPV.lCodPedido = tItemOp.lCodPedido
''                    objItemPV.sProduto = tItemOp.sProduto
''
''                    lErro = OrdemDeProducao_Grava6(alComando(4), alComando(5), objItemPV, -tItemOp.dQuantidade)
''                    If lErro <> SUCESSO And lErro <> 41346 Then Error 41601
''
''                    If lErro = 41346 Then Error 41599
''
''                End If
''
''                If objItemOp.lCodPedido <> 0 Then
''
''                    'chamar funcao que faz update incluindo qtde nova no pedido novo
''                    objItemPV.iFilialEmpresa = objItemOp.iFilialPedido
''                    objItemPV.lCodPedido = objItemOp.lCodPedido
''                    objItemPV.sProduto = objItemOp.sProduto
''
''                    lErro = OrdemDeProducao_Grava6(alComando(6), alComando(7), objItemPV, objItemOp.dQuantidade)
''                    If lErro <> SUCESSO And lErro <> 41346 Then Error 41602
''
''                    If lErro = 41346 Then Error 41436
''
''                End If
''
''            Else
''
''                If objItemOp.lCodPedido <> 0 Then
''
''                    'se trocou a qtde
''                    If tItemOp.dQuantidade <> objItemOp.dQuantidade Then
''
''                        objItemPV.iFilialEmpresa = objItemOp.iFilialPedido
''                        objItemPV.lCodPedido = objItemOp.lCodPedido
''                        objItemPV.sProduto = objItemOp.sProduto
''
''                       'chamar funcao que faz update pela diferenca
''                       lErro = OrdemDeProducao_Grava6(alComando(8), alComando(9), objItemPV, objItemOp.dQuantidade - tItemOp.dQuantidade)
''                       If lErro <> SUCESSO And lErro <> 41346 Then Error 41603
''
''                       If lErro = 41346 Then Error 41600
''
''                    End If
''
''                End If
''
''            End If
''
''        End If
''
''        If tItemOp.sCcl <> objItemOp.sCcl Then
''
''            iAlterar = 1
''
''            'bloquear ccl
''            sCcl = String(STRING_CCL, 0)
''            lErro = Comando_ExecutarLockado(alComando(3), "SELECT Ccl FROM Ccl WHERE Ccl = ?", sCcl, objItemOp.sCcl)
''            If lErro <> AD_SQL_SUCESSO Then Error 41338
''
''            lErro = Comando_BuscarPrimeiro(alComando(3))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41339
''
''            If lErro = AD_SQL_SEM_DADOS Then Error 41340
''
''            lErro = Comando_LockShared(alComando(3))
''            If lErro <> AD_SQL_SUCESSO Then Error 41341
''
''        End If
''
''        lErro = Verifica_Situacao(tItemOp.iSituacao, objItemOp.iSituacao)
''        If lErro <> SUCESSO Then Error 41342
''
''        'se houve alguma alteracao , atualiza o Item da OP
''        If iAlterar = 1 Or tItemOp.iAlmoxarifado <> objItemOp.iAlmoxarifado Or tItemOp.dQuantidade <> objItemOp.dQuantidade Or tItemOp.dtDataInicioProd <> objItemOp.dtDataInicioProd Or tItemOp.dtDataFimProd <> objItemOp.dtDataFimProd Or tItemOp.iPrioridade <> objItemOp.iPrioridade Or tItemOp.iSituacao <> objItemOp.iSituacao Or tItemOp.iDestinacao <> objItemOp.iDestinacao Then
''
''            With objItemOp
''                lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensOrdemProducao SET FilialPedido = ?, CodPedido = ?, Almoxarifado = ?, Ccl = ?, Quantidade = ?, DataInicioProd = ?, DataFimProd = ?, Prioridade = ?, Situacao = ?, Destinacao = ?", alComando(1), _
''                    .iFilialPedido, .lCodPedido, .iAlmoxarifado, .sCcl, .dQuantidade, .dtDataInicioProd, .dtDataFimProd, .iPrioridade, .iSituacao, .iDestinacao)
''            End With
''            If lErro <> AD_SQL_SUCESSO Then Error 22979
''
''            'se alterou a qtde a ser produzida
''            If objItemOp.dQuantidade <> tItemOp.dQuantidade Then
''
''                'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
''
''                objEstoqueProduto.sProduto = tItemOp.sProduto
''                objEstoqueProduto.iAlmoxarifado = tItemOp.iAlmoxarifado
''
''                lErro = CF("EstoqueProduto_AlterarQtdOP",objEstoqueProduto, objItemOp.dQuantidade - tItemOp.dQuantidade)
''                If lErro <> SUCESSO Then Error 41551
''
''            End If
''
''            'se alterou a situacao para baixado ==> diminuir a quantidade em ordem de producao da quantidade que ainda faltava ser produzido
''            If objItemOp.iSituacao = ITEMOP_SITUACAO_BAIXADA And objItemOp.iSituacao <> tItemOp.iSituacao Then
''
''                'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
''                objEstoqueProduto.sProduto = tItemOp.sProduto
''                objEstoqueProduto.iAlmoxarifado = tItemOp.iAlmoxarifado
''
''                If tItemOp.dQuantidade > tItemOp.dQuantidadeProd Then
''
''                    lErro = CF("EstoqueProduto_AlterarQtdOP",objEstoqueProduto, -(tItemOp.dQuantidade - tItemOp.dQuantidadeProd))
''                    If lErro <> SUCESSO Then Error 55254
''
''                End If
''
''                objItemOp.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
''                objItemOp.lNumIntDoc = tItemOp.lNumIntDoc
''
''                'exclui empenhos associados ao itemOP
''                lErro = CF("Empenhos_Excluir_ItemOP1",objItemOp, alComando(10), alComando(11))
''                If lErro <> SUCESSO Then Error 55259
''
''
''            End If
''
''        End If
''
''        lErro = Comando_BuscarProximo(alComando(1))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22974
''
''    Loop
''
''    'Fecha os comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        lErro = Comando_Fechar(alComando(iIndice))
''    Next
''
''    OrdemDeProducao_Grava4 = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava4:
''
''    OrdemDeProducao_Grava4 = Err
''
''    Select Case Err
''
''        Case 41338, 41339
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err)
''
''        Case 41340
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objItemOp.iAlmoxarifado)
''
''        Case 41341
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", Err)
''
''        Case 41342
''            If tItemOp.iSituacao = ITEMOP_SITUACAO_NORMAL Then sParam1 = STRING_NORMAL
''            If tItemOp.iSituacao = ITEMOP_SITUACAO_DESAB Then sParam1 = STRING_DESABILITADA
''            If tItemOp.iSituacao = ITEMOP_SITUACAO_SACR Then sParam1 = STRING_SACRAMENTADA
''            If tItemOp.iSituacao = ITEMOP_SITUACAO_BAIXADA Then sParam1 = STRING_BAIXADA
''            If objItemOp.iSituacao = ITEMOP_SITUACAO_NORMAL Then sParam2 = STRING_NORMAL
''            If objItemOp.iSituacao = ITEMOP_SITUACAO_DESAB Then sParam2 = STRING_DESABILITADA
''            If objItemOp.iSituacao = ITEMOP_SITUACAO_SACR Then sParam2 = STRING_SACRAMENTADA
''            If objItemOp.iSituacao = ITEMOP_SITUACAO_BAIXADA Then sParam2 = STRING_BAIXADA
''
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SITUACAO", Err, sParam1, sParam2)
''
''        Case 22977, 41429, 41601, 41602, 41603
''
''        Case 22971
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 22972 To 22974
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", Err)
''
''        Case 22975
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_NAO_EXCLUIVEL", Err, tItemOp.iItem, objOrdemDeProducao.sCodigo)
''
''        Case 22976
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_QTDE_MENOR_PROD", Err)
''
''        Case 22979
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSORDENSDEPRODUCAO", Err, objItemOp.iItem, objItemOp.sCodigo, objItemOp.iFilialEmpresa)
''
''        Case 40665
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", Err, objItemOp.lCodPedido, objItemOp.sProduto)
''
''        Case 41599, 41436, 41600
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", Err, objItemOp.lCodPedido, objItemOp.sProduto)
''
''        Case 41551, 55254
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152512)
''
''    End Select
''
''    'Fecha os comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        lErro = Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
''
''Private Function OrdemDeProducao_Grava6(lComando As Long, lComando1 As Long, _
''                 objItemPV As ClassItemPedido, dQuantidade As Double) As Long
'''Seleciona itemPV , Lock e a ajusta a quantidade
''
''Dim lErro As Long
''Dim sComando_Sql As String
''Dim lNumIntDoc As Long
''Dim dQuantOP As Double
''
''On Error GoTo Erro_OrdemDeProducao_Grava6
''
''   'Seleciona QuantOp e  NumIntDoc
''    sComando_Sql = "SELECT NumIntDoc, QuantOP FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ? AND Produto = ?"
''
''    lErro = Comando_ExecutarPos(lComando, sComando_Sql, 0, lNumIntDoc, dQuantOP, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto)
''    If lErro <> AD_SQL_SUCESSO Then Error 41344
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41345
''
''    If lErro = AD_SQL_SEM_DADOS Then Error 41346
''
''    objItemPV.lNumIntDoc = lNumIntDoc
''
''    'Lock no registro
''    lErro = Comando_LockExclusive(lComando)
''    If lErro <> AD_SQL_SUCESSO Then Error 41347
''
''    'Atualiza a quantOP
''    sComando_Sql = "UPDATE ItensPedidoDeVenda SET QuantOP = QuantOP + ?"
''
''    lErro = Comando_ExecutarPos(lComando1, sComando_Sql, lComando, dQuantidade)
''    If lErro <> AD_SQL_SUCESSO Then Error 41348
''
''    OrdemDeProducao_Grava6 = SUCESSO
''
''    Exit Function
''
''Erro_OrdemDeProducao_Grava6:
''
''    OrdemDeProducao_Grava6 = Err
''
''    Select Case Err
''
''        Case 41344, 41345
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)
''
''        Case 41346
''
''        Case 41347
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDODEVENDA", Err)
''
''        Case 41348
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152513)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Friend Function Verifica_Situacao(iSituacaoAnterior As Integer, iSituacaoNova As Integer) As Long
'''Verifica se a nova situacao pode ser efetivada, em função do que era a situação anterior
''
''Dim lErro As Long
''
''On Error GoTo Erro_Verifica_Situacao
''
''    Select Case iSituacaoNova
''
''        Case ITEMOP_SITUACAO_NORMAL
''
''            If (iSituacaoAnterior = ITEMOP_SITUACAO_SACR) Or (iSituacaoAnterior = ITEMOP_SITUACAO_BAIXADA) Then Error 41650
''
''        Case ITEMOP_SITUACAO_DESAB
''
''            If (iSituacaoAnterior = ITEMOP_SITUACAO_SACR) Or (iSituacaoAnterior = ITEMOP_SITUACAO_BAIXADA) Then Error 41651
''
''        Case ITEMOP_SITUACAO_SACR
''
''            If (iSituacaoAnterior = ITEMOP_SITUACAO_BAIXADA) Then Error 41652
''
''    End Select
''
''    Verifica_Situacao = SUCESSO
''
''    Exit Function
''
''Erro_Verifica_Situacao:
''
''    Verifica_Situacao = Err
''
''    Select Case Err
''
''        Case 41650, 41651, 41652
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152514)
''
''    End Select
''
''    Exit Function
''
''End Function

Function BloqueioPV_Automatico(ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, iProxSeqBloqueio As Integer) As Long
'retorna o número da próximo Sequencial de Bloqueio disponivel
'OBS: ESTA FUNÇÃO DEVE SER CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim iFim_de_Arquivo As Integer
Dim iSequencial_Automatico As Integer
Dim iConteudo As Integer
Dim iConteudo2 As Integer

On Error GoTo Erro_BloqueioPV_Automatico

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30051

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 30053

    'Pesquisa o número do Bloqueio gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT ProxSeqBloqueio FROM PedidosDeVenda WHERE Codigo = ? AND FilialEmpresa = ?", 0, iConteudo, lCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 30054

    'Lê o registro com o número do Sequencial automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 30055

    'Loca o registro que guarda o numero do Sequencial gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 30056

    If iConteudo = 0 Then iConteudo = 1 'p/garantir que comece no 1 em vez zero
    
    iSequencial_Automatico = iConteudo

    'Valor de retorno da funcao
    iProxSeqBloqueio = iSequencial_Automatico

    iSequencial_Automatico = iSequencial_Automatico + 1

    iConteudo2 = iSequencial_Automatico

    'Atualiza o numero do Sequencial gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE PedidosDeVenda SET ProxSeqBloqueio = ?", lComando, iConteudo2)
    If lErro <> AD_SQL_SUCESSO Then Error 30058

    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    BloqueioPV_Automatico = SUCESSO

    Exit Function

Erro_BloqueioPV_Automatico:

    BloqueioPV_Automatico = Err

    Select Case Err

    Case 30051, 30053
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 30054, 30055
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", Err)

    Case 30056
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, lCodigo)

    Case 30057
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", Err)

    Case 30058
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", Err, lCodigo)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152515)

    End Select

    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

'Function GeracaoOP_Grava(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
''cria ordens de produção a partir de pedidos de venda
'
'Dim lErro As Long
'Dim iIndice As Integer, lNumIntDocItem As Long, lEmpenho As Long
'Dim sDescricao As String
'Dim dtDataEmissao As Date
'Dim iNumItens As Integer
'Dim iNumItensBaixados As Integer
'Dim objItemOP As New ClassItemOP
'Dim alComando(16) As Long
'Dim lTransacao As Long
'Dim objKit As New ClassKit
'Dim objOP As New ClassOrdemDeProducao
'
'On Error GoTo Erro_GeracaoOP_Grava
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        'Abertura comandos
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 36627
'    Next
'
'    'Abertura transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 36628
'
'    'verifica se o codigo da OP não coincide com uma OP baixada - Leitura da tabela de Ordens de Produção Baixadas
'    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumItens FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 36629
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36630
'
'    'se já existe uma OP baixada com este codigo => erro
'    If lErro = AD_SQL_SUCESSO Then gError 36631
'
'    'verifica se o codigo da OP não coincide com uma OP  - Leitura da tabela de Ordens de Produção
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT DataEmissao, NumItens, NumItensBaixados FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, dtDataEmissao, iNumItens, iNumItensBaixados, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 36632
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36633
'
'    'se já existe uma OP com este codigo => erro
'    If lErro = AD_SQL_SUCESSO Then gError 36634
'
'    'insere a ordem de produção
'    lErro = Comando_Executar(alComando(2), "INSERT INTO OrdensDeProducao (Codigo, FilialEmpresa, DataEmissao, NumItens, NumItensBaixados) VALUES (?,?,?,?,?) ", objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.dtDataEmissao, objOrdemDeProducao.colItens.Count, 0)
'    If lErro <> AD_SQL_SUCESSO Then gError 36635
'
'    '#######################################################
'    'Inserido por Wagner
'    'Se o item tiver como originador o Orcamento de Venda, busca a versao
'    lErro = CF("OrdemProducao_Le_Versao_OV", objOrdemDeProducao)
'    If lErro <> SUCESSO Then gError 131315
'    '#######################################################
'
'    'Se a OP for gerar outra OP
'    '###############################################
'    'Alterado por Wagner
'    If objOrdemDeProducao.iGeraOPsArvore <> MARCADO And objOrdemDeProducao.iGeraOP = MARCADO Then
'    '###############################################
'
'        objOP.sCodigo = objOrdemDeProducao.sCodOPGerar
'        objOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
'
'        lErro = CF("OrdemProducao_Le", objOP)
'        If lErro <> SUCESSO And lErro <> 30368 Then gError 36636
'        If lErro = SUCESSO Then gError 36634
'
'    End If
'
'    'Se a OP for gerar RC ou OP de subproduto
'    '###################################################
'    'Alterado por Wagner
'    If objOrdemDeProducao.iGeraOPsArvore <> MARCADO And (objOrdemDeProducao.iGeraOP = MARCADO Or objOrdemDeProducao.iGeraReqCompra = MARCADO) Then
'        lErro = OrdemProducao_Grava_OPSub_ReqCom(objOrdemDeProducao)
'        If lErro <> SUCESSO Then gError 36636
'
'    End If
'    '###################################################
'
'    'obtem o proximo codigo para empenhos
'    lErro = OrdemDeProducao_Grava3(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
'    If lErro <> AD_SQL_SUCESSO Then gError 36636
'
'    'inclui os itens da OP
'    For Each objItemOP In objOrdemDeProducao.colItens
'
'        If objItemOP.sVersao = "" Then
'
'            'Verifica se é um kit
'            objKit.sProdutoRaiz = objItemOP.sProduto
'            lErro = CF("Kit_Le_Padrao", objKit)
'            If lErro <> SUCESSO And lErro <> 106304 Then gError 106653
'            If lErro <> SUCESSO Then gError 106654
'
'            objItemOP.sVersao = objKit.sVersao
'
'        End If
'
'        lErro = Verifica_EstoqueMaximo(objItemOP)
'        If lErro <> SUCESSO Then gError 36637
'
'        objItemOP.lNumIntDoc = lNumIntDocItem
'
'        lErro = CF("OrdemDeProducao_Grava1", objItemOP, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
'        If lErro <> SUCESSO Then gError 36638
'
'        lNumIntDocItem = lNumIntDocItem + 1
'
'    Next
'
'    'atualiza o proximo codigo para lNumIntDocItem e lEmpenho
'    lErro = OrdemDeProducao_Grava5(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
'    If lErro <> AD_SQL_SUCESSO Then gError 36639
'
'    '#####################
'    'INSERIDO POR WAGNER
'    If objOrdemDeProducao.iGeraOPsArvore = MARCADO Then
'        lErro = CF("OrdemProducao_Grava_Nodes", objOrdemDeProducao)
'        If lErro <> SUCESSO Then gError 131314
'    End If
'    '#####################
'
'    'Inserido por Jorge Specian - 14/06/2005
'    '---------------------------------------
'    lErro = CF("GeracaoOP_GravaOperacoes_EmTrans", objOrdemDeProducao)
'    If lErro <> SUCESSO Then gError 137182
'    '---------------------------------------
'
'    'Confirma transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 36640
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    GeracaoOP_Grava = SUCESSO
'
'    Exit Function
'
'Erro_GeracaoOP_Grava:
'
'    GeracaoOP_Grava = gErr
'
'    Select Case gErr
'
'        Case 36627
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 36628
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 36629, 36630
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS", gErr)
'
'        Case 36631
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemDeProducao.sCodigo)
'
'        Case 36632, 36633
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)
'
'        Case 36634
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOrdemDeProducao.sCodigo)
'
'        Case 36635
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDEPRODUCAO", gErr)
'
'        Case 36636, 36637, 36638, 36639, 106653, 126479, 126481, 126482
'
'        Case 36640
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 106654
'            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz, objKit.sVersao)
'
'        Case 126480
'            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOrdemDeProducao.sCodigo)
'
'        Case 131314, 131315, 137182  '(137182 - inserido por Jorge Specian - 14/06/2005)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152516)
'
'    End Select
'
'    'Fechamento transação
'    Call Transacao_Rollback
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function PedidoVenda_Grava_EmTrans(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava o Pedido de Venda

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 29) As Long
'Dim lTransacao As Long
Dim lNumIntDoc As Long
Dim objPedVenda_BD As New ClassPedidoDeVenda
Dim tPedVenda_BD As typePedidoVenda
Dim dQuantNaoCancelada As Double
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_PedidoVenda_Grava_EmTrans

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 42941
    Next

    lErro = CF("PedidoVenda_Grava_Customizacoes", objPedidoVenda)
    If lErro <> SUCESSO Then gError 126630

    'Faz o lock nos objetos vinculados ao Pedido
    lErro = CF("PVObjetosVinculados_Lock", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42924
    
    'Obtém o andamento automático
    lErro = CF("PV_Andamento_Auto", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42924

    'Verifica se o Pedido passado está baixado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PedidosDeVendaBaixados WHERE Codigo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objPedidoVenda.lCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 42939

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 42940

    'Se estiver --> gErro. Sai.
    If lErro = AD_SQL_SUCESSO Then gError 42972

    'Lê o Pedido
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, DataEmissao, Cliente, Filial, FilialEmpresaFaturamento FROM PedidosDeVenda WHERE Codigo = ? AND FilialEmpresa = ?", 0, tPedVenda_BD.lNumIntDoc, tPedVenda_BD.dtDataEmissao, tPedVenda_BD.lCliente, tPedVenda_BD.iFilial, tPedVenda_BD.iFilialEmpresaFaturamento, objPedidoVenda.lCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 42939

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 42973

    'Se não existir
    If lErro <> AD_SQL_SUCESSO Then
        
        'Insere o Pedido novo no Banco de Dados
        lErro = CF("PedidoVenda_Insere_BD", objPedidoVenda)
        If lErro <> SUCESSO Then gError 42942
        
        'Grava os dados estatísticos - INSERÇÃO
        lErro = CF("PedVenda_Grava_Estatisticas", objPedidoVenda)
        If lErro <> SUCESSO Then gError 52829
        
    
    'Se existir
    Else

        objPedVenda_BD.lNumIntDoc = tPedVenda_BD.lNumIntDoc
        objPedVenda_BD.dtDataEmissao = tPedVenda_BD.dtDataEmissao
        objPedidoVenda.lNumIntDoc = tPedVenda_BD.lNumIntDoc
        objPedVenda_BD.lCliente = tPedVenda_BD.lCliente
        objPedVenda_BD.iFilial = tPedVenda_BD.iFilial
        objPedVenda_BD.iFilialEmpresaFaturamento = tPedVenda_BD.iFilialEmpresaFaturamento

        'Altera o Pedido de Venda e armazena os ítens do Pedido no BD em objPedVenda_BD
        lErro = CF("PedidoVenda_Altera_BD", objPedidoVenda, objPedVenda_BD, alComando)
        If lErro <> SUCESSO Then gError 42943
                
        'Grava os dados estatísticos - ALTERAÇÃO
        lErro = CF("PedVenda_Grava_Estatisticas", objPedidoVenda, objPedVenda_BD)
        If lErro <> SUCESSO Then gError 52830
        
    End If

    lErro = CF("ItensPVEntrega_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 183207

    'Verifica se tem quantidade não cancelada
    lErro = PedVenda_QuantNaoCancelada(objPedidoVenda, dQuantNaoCancelada)
    If lErro <> SUCESSO Then gError 36713

    'Se nao tiver quantidade nao cancelada, baixa o Pedido
    If dQuantNaoCancelada = 0 Then

        lErro = CF("PedidoDeVenda_Baixar", alComando, objPedidoVenda, False)
        If lErro <> SUCESSO Then gError 46087
         
    End If
    
    'Vincula ou desvincula um Orçamento de Venda a um Pedido de Venda.
    lErro = PedidoVenda_AlteraVinculo_Orcamento(objPedidoVenda) 'Inserida por Leo em 26/04/02
    If lErro <> SUCESSO Then gError 94464
    
    lErro = CF("ProjetoInfo_Grava", objPedidoVenda.objProjetoInfo, PRJ_CR_TIPO_PV, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181706
    
    objPedidoVenda.objInfoAdic.lNumIntDoc = objPedidoVenda.lNumIntDoc
    objPedidoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PV
    lErro = CF("InfoAdicionais_Grava", objPedidoVenda.objInfoAdic)
    If lErro <> SUCESSO Then gError 181706
    
    lErro = CF("PedidoVenda_Grava_Customizado", objPedidoVenda)
    If lErro <> SUCESSO Then gError 140305
    
    lErro = CF("PV_Grava_Andamento_EmTrans", objPedidoVenda, gsUsuario)
    If lErro <> SUCESSO Then gError 140305
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoVenda_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PedidoVenda_Grava_EmTrans:

    PedidoVenda_Grava_EmTrans = gErr

    Select Case gErr

        Case 36713, 42924, 42942, 42943, 46087, 46093, 46094, 52829, 52830, 94464, 126630, 178179, 181706, 183207, 140305 'Tratados nas rotinas chamadoras

        Case 42923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 42939, 42940, 42973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDABAIXADOS", gErr, objPedidoVenda.lCodigo)

        Case 42941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 42972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDVENDA_BAIXADO_ALTERACAO", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152517)

    End Select

'    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoVenda_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava o Pedido de Venda

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PedidoVenda_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 42923

    lErro = CF("PedidoVenda_Grava_EmTrans", objPedidoVenda)
    If lErro <> SUCESSO Then gError 190054
    
    If objPedidoVenda.iParc > 1 Then
        lErro = CF("PedidoVenda_Grava_Parc", objPedidoVenda)
        If lErro <> SUCESSO Then gError 190054
    End If
    
    'Faz o commit na transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 42944

    lErro = CF("WorkFlow_Trata_Transacao", "FAT", objPedidoVenda.objTela, gsUsuario, objPedidoVenda)
    If lErro <> SUCESSO Then gError 178179

    PedidoVenda_Grava = SUCESSO

    Exit Function

Erro_PedidoVenda_Grava:

    PedidoVenda_Grava = gErr

    Select Case gErr

        Case 42923
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 42944
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 178179, 190054

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190055)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

'??? Inserida por Leo em 26/04/02 ... Confirmar se este realmente será o local desta função
Function PedidoVenda_AlteraVinculo_Orcamento(objPedidoVenda As ClassPedidoDeVenda)
'Vincula ou desvincula um Orçamento de Venda a um Pedido de Venda.

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim lComando As Long
Dim lComando2 As Long
Dim lNumIntPedVenda As Long
Dim lNumIntNFiscal As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PedidoVenda_AlteraVinculo_Orcamento

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94465
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 94466
        
    'Se o usuário selecionou um Orcamento
    If objPedidoVenda.lNumIntOrcVenda <> 0 Then
    
        objOrcamentoVenda.lNumIntDoc = objPedidoVenda.lNumIntOrcVenda
    
        'Seleciona o Orçamento na Tela de OrcamentoVenda
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntPedVenda, NumIntNFiscal, FilialEmpresa, Codigo FROM OrcamentoVenda WHERE NumIntDoc = ?", _
        0, lNumIntPedVenda, lNumIntNFiscal, iFilialEmpresa, lCodigo, objOrcamentoVenda.lNumIntDoc)
        If lErro <> SUCESSO Then gError 94467
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94468
        
        'Se o Orçamento não existe ---> Erro.
        If lErro <> AD_SQL_SUCESSO Then gError 94469
                
        objOrcamentoVenda.lNumIntPedVenda = lNumIntPedVenda
        objOrcamentoVenda.lNumIntNFiscal = lNumIntNFiscal
        
        'Se o Orcamento estiver vinculado a uma Nota Fiscal ou a um pedido diferente ---> Erro.
        If objOrcamentoVenda.lNumIntPedVenda <> 0 And objOrcamentoVenda.lNumIntPedVenda <> objPedidoVenda.lNumIntDoc Then gError 94470
        If objOrcamentoVenda.lNumIntNFiscal <> 0 Then gError 94471
    
        'Vincula o Orçamento ao Pedido
        lErro = Comando_ExecutarPos(lComando2, "UPDATE OrcamentoVenda SET NumIntPedVenda = ?", lComando, objPedidoVenda.lNumIntDoc)
        If lErro <> SUCESSO Then gError 94472
        
    'Se o não foi seleciona Orçamento
    Else
        
        objOrcamentoVenda.lNumIntPedVenda = objPedidoVenda.lNumIntDoc
        
        'Tenta selecionar algum Orçamento Vinculado ao Pedido.
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntNFiscal, Codigo, FilialEmpresa FROM OrcamentoVenda WHERE NumIntPedVenda = ?", _
        0, lNumIntNFiscal, lCodigo, iFilialEmpresa, objOrcamentoVenda.lNumIntPedVenda)
        If lErro <> SUCESSO Then gError 94473
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94474
        
        'Se encontrou um Orçamento Vinculado ao Pedido
        If lErro = AD_SQL_SUCESSO Then
            
            'Desvincula o Orçamento do Pedido, pois o usuário está gravando um Pedido sem Orçamento.
            lErro = Comando_ExecutarPos(lComando2, "UPDATE OrcamentoVenda SET NumIntPedVenda = ?", lComando, objPedidoVenda.lNumIntOrcVenda)
            If lErro <> SUCESSO Then gError 94475
            
        End If
        
    End If
            
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    PedidoVenda_AlteraVinculo_Orcamento = SUCESSO
    
    Exit Function
    
Erro_PedidoVenda_AlteraVinculo_Orcamento:
    
    PedidoVenda_AlteraVinculo_Orcamento = gErr
    
    Select Case gErr
                              
        Case 94465, 94466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                                          
        Case 94467, 94468, 94473, 94474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)
            
        Case 94469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, lCodigo, iFilialEmpresa)
            
        Case 94470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_PEDVENDA", gErr, lCodigo, iFilialEmpresa)
            
        Case 94471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_NFISCAL", gErr, lCodigo, iFilialEmpresa)
                       
        Case 94472, 94475
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOVENDA", gErr, lCodigo, iFilialEmpresa)
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152518)
     
    End Select
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
        
    Exit Function

End Function
    
Function PedidoVenda_Grava_BD(ByVal objPedidoVenda As ClassPedidoDeVenda, Optional ByVal lComando1 As Long = 0) As Long
'Insere o Pedido de Venda passado na Tabela de Pedidos de Venda
'ou atualiza caso o Pedido já exista

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer

On Error GoTo Erro_PedidoVenda_Grava_BD

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 42958
    Next

    'Verifica se é uma Inclusão
    If lComando1 = 0 Then
        'Insere o novo Pedido no BD
        lErro = CF("PedidoDeVenda_Insere", objPedidoVenda, alComando(0))
        If lErro <> SUCESSO Then gError 42959
    Else 'lComando1 foi passado
        'Altera o Pedido de Venda
        lErro = Comando_ExecutarPos(alComando(0), "UPDATE PedidosDeVenda SET DataEmissao = ?, FilialEntrega = ?, CodTransportadora = ?,CondicaoPagto =?, PercAcrescFinanceiro= ?, DataReferencia =?, MensagemNota =?,NaturezaOp = ?, ValorProdutos =?,ValorTotal =?, ValorFrete=?, ValorDesconto=?, ValorSeguro =?, ValorOutrasDespesas =?, PedidoCliente = ?, CanalVenda = ?,TabelaPreco =?, FaturaIntegral = ?, ComissaoAutomatica = ?, CobrancaAutomatica = ?,Placa = ? ,PlacaUF = ? ,VolumeQuant = ?, VolumeEspecie = ? ,VolumeMarca = ?,PesoLiq = ? ,PesoBruto = ?, VolumeNumero = ?, FreteRespons = ?, FilialEmpresaFaturamento = ?, CodTranspRedesp = ?, DetPagFrete = ?, DataEntrega = ?, PedidoRepresentante = ?, DataRefFluxo = ?, VolumeTotal = ?, ValorDescontoTit = ?,ValorDescontoItens = ?,ValorItens = ?, Parc = ?, UsuarioUltAlteracao = ?, DataAlteracao = ?, HoraAlteracao = ?", lComando1, _
            objPedidoVenda.dtDataEmissao, objPedidoVenda.iFilialEntrega, objPedidoVenda.iCodTransportadora, objPedidoVenda.iCondicaoPagto, objPedidoVenda.dPercAcrescFinanceiro, objPedidoVenda.dtDataReferencia, objPedidoVenda.sMensagemNota, objPedidoVenda.sNaturezaOp, objPedidoVenda.dValorProdutos, objPedidoVenda.dValorTotal, objPedidoVenda.dValorFrete, objPedidoVenda.dValorDesconto, objPedidoVenda.dValorSeguro, objPedidoVenda.dValorOutrasDespesas, objPedidoVenda.sPedidoCliente, objPedidoVenda.iCanalVenda, objPedidoVenda.iTabelaPreco, objPedidoVenda.iFaturaIntegral, objPedidoVenda.iComissaoAutomatica, objPedidoVenda.iCobrancaAutomatica, objPedidoVenda.sPlaca, objPedidoVenda.sPlacaUF, objPedidoVenda.lVolumeQuant, objPedidoVenda.lVolumeEspecie, objPedidoVenda.lVolumeMarca, objPedidoVenda.dPesoLiq, objPedidoVenda.dPesoBruto, objPedidoVenda.sVolumeNumero, objPedidoVenda.iFreteRespons, _
            objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.iCodTranspRedesp, objPedidoVenda.iDetPagFrete, objPedidoVenda.dtDataEntrega, objPedidoVenda.sPedidoRepresentante, objPedidoVenda.dtDataRefFluxo, objPedidoVenda.dVolumeTotal, objPedidoVenda.dValorDescontoTit, objPedidoVenda.dValorDescontoItens, objPedidoVenda.dValorItens, objPedidoVenda.iParc, gsUsuario, Date, CDbl(Time))
        If lErro <> AD_SQL_SUCESSO Then gError 46004
        
        '#################################################
        'Inserido por Wagner 05/07/2006
        lErro = CF("PV_Grava_Customizado", objPedidoVenda)
        If lErro <> SUCESSO Then gError 180411
        '#################################################

        lErro = CF("PedidoVenda_ExcluiTributacao", objPedidoVenda, alComando(1), alComando(2), alComando(5), alComando(6))
        If lErro <> SUCESSO Then gError 42377
    End If

    lErro = CF("PedidoVenda_GravaTributacao", objPedidoVenda, alComando(3), alComando(4))
    If lErro <> SUCESSO Then gError 42375

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoVenda_Grava_BD = SUCESSO

    Exit Function

Erro_PedidoVenda_Grava_BD:

    PedidoVenda_Grava_BD = gErr

    Select Case gErr

        Case 42375, 42377, 42959, 180411

        Case 42958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 46004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152519)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ComissoesPV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda, Optional ByVal bDesfazendoBaixa As Boolean = False) As Long
'Insere no Bd as Comissões do Pedido de Venda passado

Dim lErro As Long
Dim lComando As Long, lComando2 As Long
Dim objComissaoPV As ClassComissaoPedVendas
Dim objComissaoPorItem As ClassComissaoPorItem

On Error GoTo Erro_ComissoesPV_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 42960

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 42960

    'Para cada Comissão da coleção
    For Each objComissaoPV In objPedidoVenda.colComissoes
    
        'Insere um registro em ComissoesPV
        'Tulio - Incluido campo Direto/Indireto na gravacao...
        lErro = Comando_Executar(lComando, "INSERT INTO ComissoesPedVendas (FilialEmpresa,PedidoDeVendas,CodVendedor,ValorBase,Percentual,Valor,PercentualEmissao,ValorEmissao,Indireta,Seq) VALUES (?,?,?,?,?,?,?,?,?,?) ", objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objComissaoPV.iCodVendedor, objComissaoPV.dValorBase, objComissaoPV.dPercentual, objComissaoPV.dValor, objComissaoPV.dPercentualEmissao, objComissaoPV.dValorEmissao, objComissaoPV.iIndireta, objComissaoPV.iSeq)
        If lErro <> AD_SQL_SUCESSO Then Error 42961

        If bDesfazendoBaixa = False Then
        
            For Each objComissaoPorItem In objComissaoPV.colLinhas
            
                objComissaoPorItem.iTipoDoc = TIPO_PV
                objComissaoPorItem.iSeqComissao = objComissaoPV.iSeq
                
                If objComissaoPorItem.iLinha >= 1 And objPedidoVenda.colItens.Count >= objComissaoPorItem.iLinha Then objComissaoPorItem.lNumIntDocItem = objPedidoVenda.colItens.Item(objComissaoPorItem.iLinha).lNumIntDoc
                
                With objComissaoPorItem
                    lErro = Comando_Executar(lComando2, "INSERT INTO ComissoesPorItem (TipoDoc, NumIntDocItem, SeqComissao, ValorBase, Percentual, Valor, Linha, PercentualEmissao, ValorEmissao) VALUES (?,?,?,?,?,?,?,?,?)", _
                        .iTipoDoc, .lNumIntDocItem, .iSeqComissao, .dValorBase, .dPercentual, .dValor, .iLinha, .dPercentualEmissao, .dValorEmissao)
                End With
                If lErro <> AD_SQL_SUCESSO Then Error 42961
            
            Next

        End If

    Next

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    ComissoesPV_Grava = SUCESSO

    Exit Function

Erro_ComissoesPV_Grava:

    ComissoesPV_Grava = Err
    
    Select Case Err

        Case 42960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 42961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESPV", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152520)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function ParcelasPV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere as no BD Parcelas do Pedido passado

Dim lErro As Long
Dim lComando As Long
Dim objParcelaPV As ClassParcelaPedidoVenda

On Error GoTo Erro_ParcelasPV_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 42962

    'Para cada parcela da coleção
    For Each objParcelaPV In objPedidoVenda.colParcelas

        'Insere em ParcelasPedidoDeVenda
        lErro = Comando_Executar(lComando, "INSERT INTO ParcelasPedidoDeVenda (FilialEmpresa,PedidodeVenda,NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objParcelaPV.iNumParcela, objParcelaPV.dtDataVencimento, objParcelaPV.dValor, objParcelaPV.iDesconto1Codigo, objParcelaPV.dtDesconto1Ate, objParcelaPV.dDesconto1Valor, _
        objParcelaPV.iDesconto2Codigo, objParcelaPV.dtDesconto2Ate, objParcelaPV.dDesconto2Valor, objParcelaPV.iDesconto3Codigo, objParcelaPV.dtDesconto3Ate, objParcelaPV.dDesconto3Valor, objParcelaPV.iTipoPagto, objParcelaPV.iCodConta, objParcelaPV.dtDataCredito, objParcelaPV.dtDataEmissaoCheque, objParcelaPV.iBancoCheque, objParcelaPV.sAgenciaCheque, objParcelaPV.sContaCorrenteCheque, objParcelaPV.lNumeroCheque, objParcelaPV.dtDataDepositoCheque, objParcelaPV.iAdmMeioPagto, objParcelaPV.iParcelamento, objParcelaPV.sNumeroCartao, objParcelaPV.dtValidadeCartao, objParcelaPV.sAprovacaoCartao, objParcelaPV.dtDataTransacaoCartao)
        If lErro <> AD_SQL_SUCESSO Then Error 42963

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ParcelasPV_Grava = SUCESSO

    Exit Function

Erro_ParcelasPV_Grava:

    ParcelasPV_Grava = Err

    Select Case Err

        Case 42962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", Err)

        Case 42963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELASPV", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152521)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere os itens do Pedido de Venda passado no BD

Dim alComando(0 To 5) As Long
Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOP As New ClassItemOP
Dim sProduto As String
Dim iStatus As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ItensPV_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 42964
    Next

    'Para cada item da coleção
    For Each objItemPV In objPedidoVenda.colItensPedido
    
        For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
            objItemRomaneio.lNumIntDoc = 0
        Next

        sProduto = String(STRING_PRODUTO, 0)
        
        '############################################################
        'Alterado por Wagner 26/04/06
        lErro = CF("ItensPVAux_Grava", alComando(0), sProduto, iStatus, lNumIntDoc, objPedidoVenda, objItemPV)
        If lErro <> SUCESSO Then gError 46045
'        'Localiza o item no Banco de Dados
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, Status, NumIntDoc FROM ItensPedidoDeVenda WHERE CodPedido = ? AND FilialEmpresa = ? AND Produto = ? ", 0, sProduto, iStatus, lNumIntDoc, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, objItemPV.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 46045
        '############################################################

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46046
        If lErro <> AD_SQL_SUCESSO Then

            'Obtém o número interno automático do item
            lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ITEM_PEDVENDA", lNumIntDoc)
            If lErro <> SUCESSO Then gError 42965

            objItemPV.iFilialEmpresa = giFilialEmpresa
            objItemPV.lCodPedido = objPedidoVenda.lCodigo
            objItemPV.lNumIntDoc = lNumIntDoc
            
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            '############################################################
            'Alterado por Wagner 26/04/06
            lErro = CF("ItensPV_Insere2", alComando(0), objItemPV, alComando(5))
            If lErro <> SUCESSO Then gError 177399
'            'Insere o item no BD
'            lErro = Comando_Executar(alComando(0), "INSERT INTO ItensPedidoDeVenda (NumIntDoc,FilialEmpresa,CodPedido,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.sLote, objItemPV.iStatus, objItemPV.dValorAbatComissao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.dQuantOP, objItemPV.dQuantSC)
'            If lErro <> AD_SQL_SUCESSO Then gError 42966
            '############################################################
        
        'Se for uma alteração
        Else
            
            objItemPV.lNumIntDoc = lNumIntDoc
            
            'Se houve alteracao de Produto
            If sProduto <> objItemPV.sProduto Then
            
                'Verifica se ítem estava vinculado a Ordem Produção
                lErro = CF("ItemOP_Le_ItemPV", objItemOP, objItemPV)
                If lErro <> SUCESSO And lErro <> 46074 Then gError 24461
                If lErro = SUCESSO Then gError 24462
            
            End If
            
            objItemPV.iStatus = iStatus
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            'Faz a atualização no Banco de Dados
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedidoDeVenda SET Produto =?, Quantidade = ?, PrecoUnitario = ?, PrecoTotal = ?, ClasseUM = ?, UnidadeMed = ?, ValorDesconto =?, DataEntrega = ?, Descricao = ?, QuantCancelada = ?, QuantReservada = ?, QuantFaturada = ?, Status =?, PercentMaisReceb = ?, PercentMenosReceb = ?, RecebForaFaixa = ?, QuantFaturadaAMais = ?, Prioridade = ?, Comissao = ?, TabelaPreco = ?", alComando(0), objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.iStatus, objItemPV.dPercentMaisReceb, objItemPV.dPercentMenosReceb, objItemPV.iRecebForaFaixa, objItemPV.dQuantFaturadaAMais, objItemPV.iPrioridade, objItemPV.dComissao, objItemPV.iTabelaPreco)
            If lErro <> AD_SQL_SUCESSO Then gError 46048

            lErro = CF("ItemPV_ExcluiTributacao", objItemPV, alComando(2), alComando(3))
            If lErro <> SUCESSO Then gError 42985
            
            lErro = CF("ItensPVGrade_Exclui", objItemPV, True)
            If lErro <> SUCESSO Then gError 86338
            
            objItemPV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMPV
            objItemPV.objInfoAdicDocItem.lNumIntDocItem = objItemPV.lNumIntDoc
        
            lErro = CF("InfoAdicDocItem_Altera", objItemPV.objInfoAdicDocItem)
            If lErro <> SUCESSO Then gError 86338
        
        
        End If

        If objItemPV.colItensRomaneioGrade.Count > 0 Then
        
            lErro = CF("ItensPVGrade_Inclui", objItemPV)
            If lErro <> SUCESSO Then gError 86339
    
        End If
        
        objItemPV.lNumIntDocPV = objPedidoVenda.lNumIntDoc
        
        lErro = CF("ItemPV_GravaTributacao", objItemPV, alComando(4))
        If lErro <> SUCESSO Then gError 42365

    Next

    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("PVItens_Grava_Customizado", objPedidoVenda)
    If lErro <> SUCESSO Then gError 177047
    '############################################

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPV_Grava = SUCESSO

    Exit Function

Erro_ItensPV_Grava:

    ItensPV_Grava = gErr

    Select Case gErr

        Case 24461, 42365, 42376, 42985, 51386, 86338, 86339, 177047, 177399
        
        Case 24462
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_VINCULADO_ITEM_OP", gErr, objItemPV.iItem, objItemOP.lNumIntDoc)

        Case 42964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPV", gErr)

        Case 46045, 46046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 46047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_ENCONTRADO", gErr, objItemPV.lNumIntDoc)

        Case 46048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr, objItemPV.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152522)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Reserva_Automatico_Trans(lCodigo As Long) As Long
'Retorna o número da próxima Reserva disponível
'Chamada em Transação

Dim lErro As Long

On Error GoTo Erro_Reserva_Automatico_Trans

    lErro = CF("Config_ObterAutomatico_EmTrans", "MATConfig", NUM_PROX_RESERVA, "Reserva", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 52801

    Reserva_Automatico_Trans = SUCESSO

    Exit Function

Erro_Reserva_Automatico_Trans:

    Reserva_Automatico_Trans = Err

    Select Case Err

        Case 52801 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152523)

    End Select

    Exit Function

End Function

Function Reserva_Automatico(lCodigo As Long) As Long
'Retorna o número da próxima reserva disponível

Dim lErro As Long

On Error GoTo Erro_Reserva_Automatico

    lErro = CF("Config_ObterAutomatico", "MATConfig", NUM_PROX_RESERVA, "Reserva", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 23932

    Reserva_Automatico = SUCESSO

    Exit Function

Erro_Reserva_Automatico:

    Reserva_Automatico = Err

    Select Case Err

        Case 23932

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152524)

    End Select

    Exit Function

End Function

Function ReservasPV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava as Reservas

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim objReserva As New ClassReserva
Dim objItem As ClassItemPedido
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_ReservasPV_Grava

    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 42967
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 42967

    'Para cada item do pedido
    For Each objItem In objPedidoVenda.colItensPedido

        'Zera a variável de número interno
        lNumIntDoc = 0

        'Gera o numero internos para as reservas
        lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_RESERVA", lNumIntDoc, objItem.ColReserva.Count)
        If lErro <> SUCESSO Then gError 42968

        If objItem.colItensRomaneioGrade.Count = 0 Then
    
            'Para cada reserva de cada item
            For Each objReserva In objItem.ColReserva
                
                'Obtém o Código Automático
                lErro = CF("Reserva_Automatico_Trans", lCodigo)
                If lErro <> SUCESSO Then gError 42969
    
                 With objReserva
                    .lNumIntDoc = lNumIntDoc
                    .iTipoDoc = PEDIDO_DE_VENDA
                    .dtDataReserva = gdtDataHoje
                    .sCodUsuario = gsUsuario
                    .lDocOrigem = objPedidoVenda.lCodigo
                    .lNumIntOrigem = objItem.lNumIntDoc
                    .lCodigo = lCodigo
                    .sProduto = objItem.sProduto
                    .iFilialEmpresa = objItem.iFilialEmpresa
                    
                
                    'Insere a reserva no Banco de Dados
                    lErro = Comando_Executar(lComando, "INSERT INTO Reserva (FilialEmpresa,Codigo,Produto,Almoxarifado,TipoDoc,DocOrigem,NumIntOrigem,Quantidade,DataReserva,DataValidade,CodUsuario,Responsavel,NumIntDoc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .lCodigo, .sProduto, .iAlmoxarifado, .iTipoDoc, .lDocOrigem, .lNumIntOrigem, .dQuantidade, .dtDataReserva, .dtDataValidade, .sCodUsuario, .sResponsavel, .lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 42970
    
                End With
                lNumIntDoc = lNumIntDoc + 1
            Next

        End If

        For Each objItemRomaneio In objItem.colItensRomaneioGrade
            
            For Each objReservaItem In objItemRomaneio.colLocalizacao
                   
                lNumIntDoc = 0
    
                'Gera o numero internos para as reservas
                lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_RESERVA", lNumIntDoc)
                If lErro <> SUCESSO Then gError 42968
                   
                'Obtém o Código Automático
                lErro = CF("Reserva_Automatico_Trans", lCodigo)
                If lErro <> SUCESSO Then gError 42969
    
                 With objReserva
                    .lNumIntDoc = lNumIntDoc
                    .iTipoDoc = PEDIDO_DE_VENDA_GRADE
                    .dtDataReserva = gdtDataHoje
                    .sCodUsuario = gsUsuario
                    .lDocOrigem = objPedidoVenda.lCodigo
                    .lNumIntOrigem = objItemRomaneio.lNumIntDoc
                    .lCodigo = lCodigo
                    .sProduto = objItemRomaneio.sProduto
                    .iFilialEmpresa = objItem.iFilialEmpresa
                    .dQuantidade = objReservaItem.dQuantidade
                    .dtDataValidade = objReservaItem.dtDataValidade
                    .iAlmoxarifado = objReservaItem.iAlmoxarifado
                    .iFilialEmpresa = objReservaItem.iFilialEmpresa
                    .sAlmoxarifado = objReservaItem.sAlmoxarifado
                    .sResponsavel = objReservaItem.sResponsavel
                                       
                    'Insere a reserva no Banco de Dados
                    lErro = Comando_Executar(lComando1, "INSERT INTO Reserva (FilialEmpresa,Codigo,Produto,Almoxarifado,TipoDoc,DocOrigem,NumIntOrigem,Quantidade,DataReserva,DataValidade,CodUsuario,Responsavel,NumIntDoc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .lCodigo, .sProduto, .iAlmoxarifado, .iTipoDoc, .lDocOrigem, .lNumIntOrigem, .dQuantidade, .dtDataReserva, .dtDataValidade, .sCodUsuario, .sResponsavel, .lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 42970
                
                End With
                   
            Next
            
        Next

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ReservasPV_Grava = SUCESSO

    Exit Function

Erro_ReservasPV_Grava:

    ReservasPV_Grava = gErr

    Select Case gErr

        Case 42967
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42968, 42969

        Case 42970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RESERVASPV", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152525)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BloqueioCredito_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim dValorPedido As Double
Dim objItemPV As ClassItemPedido
Dim tCliente As typeCliente
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_BloqueioCredito_Adiciona

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 42974

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 42979

    dValorPedido = 0

    dValorPedido = objPedidoVenda.dValorTotal

    'Lê os saldos e o limite de credito do Cliente
    lErro = Comando_ExecutarPos(lComando, "SELECT LimiteCredito, Bloqueado FROM Clientes WHERE Codigo = ?", 0, tCliente.dLimiteCredito, tCliente.iBloqueado, objPedidoVenda.lCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 42975

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42976
    If lErro = AD_SQL_SEM_DADOS Then Error 42977 'Não encontrou

    'Faz o lock do Cliente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 42978
    
    objClienteEstatistica.lCodCliente = objPedidoVenda.lCliente
    objClienteEstatistica.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
    'Le dinamicamente o Saldo dos Titulos e dos Pedidos Liberados
    lErro = CF("Cliente_Le_Estatistica_Credito", objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52952
        
    'Verifica se a Soma dos Titulos, Pedidos, NFs não faturadas e valor do Pedido ultrapassa o limite do Cliente
    If tCliente.iBloqueado <> 0 Or (tCliente.dLimiteCredito < (objClienteEstatistica.dSaldoTitulos + objClienteEstatistica.dSaldoPedidosLiberados + objClienteEstatistica.dValorNFsNaoFaturadas + dValorPedido)) Then

        'Adiciona o bloqueio de crédito ao Pedido
        objPedidoVenda.colBloqueiosPV.Add giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_CREDITO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""

    End If

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    BloqueioCredito_Adiciona = SUCESSO

    Exit Function

Erro_BloqueioCredito_Adiciona:

    BloqueioCredito_Adiciona = Err

    Select Case Err

        Case 42974, 42979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 42975, 42976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 42977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objPedidoVenda.lCliente)

        Case 42978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", Err, objPedidoVenda.lCliente)

        Case 42980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", Err, objPedidoVenda.lCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152526)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ComissoesPV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui do BD as Comissões do Pedido de Venda passado

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iCodVendedor As Integer

On Error GoTo Erro_ComissoesPV_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46007

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 46008

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 46008

    'Busca as comissões vinculadas ao Pedido de Venda passado
    lErro = Comando_ExecutarPos(lComando, "SELECT CodVendedor FROM ComissoesPedVendas WHERE FilialEmpresa =? AND PedidoDeVendas = ?", 0, iCodVendedor, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46009

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46010

    'Para cada comissão encontrada
    Do While lErro = AD_SQL_SUCESSO

        'Exclui a comissão da tabela de Comissões de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ComissoesPedVendas", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 46011

        'Busca a próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46012

    Loop

    lErro = Comando_Executar(lComando2, "DELETE FROM ComissoesPorItem WHERE TipoDoc = ? AND NumIntDocItem IN (SELECT I.NumIntDoc FROM ItensPVTodos I, PedVenTodos P WHERE I.FilialEmpresa = P.FilialEmpresa AND I.CodPedido = P.Codigo AND P.FilialEmpresa = ? AND P.Codigo = ?)", _
        TIPO_PV, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46011
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    ComissoesPV_Exclui = SUCESSO

    Exit Function

Erro_ComissoesPV_Exclui:

    ComissoesPV_Exclui = Err
    
    Select Case Err

        Case 46007, 46008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43009, 46010, 46012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDVENDAS", Err)

        Case 46011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESPEDVENDAS", Err, objPedidoVenda.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152527)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function ParcelasPV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui do BD as Parcelas do Pedido

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iNumParcela As Integer

On Error GoTo Erro_ParcelasPV_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46018

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 46019

    'Busca as Parcelas do Pedido de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT NumParcela FROM ParcelasPedidoDeVenda WHERE FilialEmpresa = ? AND PedidodeVenda = ?", 0, iNumParcela, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46019

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46020

    'Para cada parcela encontrada
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui a parcela da tabela de parcelas de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasPedidoDeVenda", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 46021

        'Busca a Próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46022

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasPV_Exclui = SUCESSO

    Exit Function

Erro_ParcelasPV_Exclui:

    ParcelasPV_Exclui = Err

    Select Case Err

        Case 46018, 46019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46019, 46020, 46022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPEDIDODEVENDA", Err, objPedidoVenda.lCodigo)

        Case 46021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASPV", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152528)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BloqueioReserva_Adiciona2(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objBloqueioPV_BD As ClassBloqueioPV, objBloqueioPV As ClassBloqueioPV, iBloqueioEstoqueAdicionado As Integer, dPedidoFaltaReservar As Double) As Long
'Adiciona um bloqueio BLOQUEIO_PARCIAL ou BLOQUEIO_NAO_RESERVA
'ao Pedido de Venda

Dim lErro As Long
Dim bBloqueio_Total As Boolean
Dim dPedidoReservado As Double
Dim objItemPV As ClassItemPedido
Dim dQuantFaltaReservar As Double
Dim iTipoBloqueio As Integer
Dim iIndice As Integer

On Error GoTo Erro_BloqueioReserva_Adiciona2

    bBloqueio_Total = False
    iIndice = 0

    'Verifica se na coleção de Bloqueios há um bloqueio total
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV

        iIndice = iIndice + 1

        If objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
            bBloqueio_Total = True
            Exit For
        End If
    Next

    dPedidoFaltaReservar = 0

    'SE não houver
    If Not bBloqueio_Total Then
        'Para cada item do Pedido
        For Each objItemPV In objPedidoVenda.colItensPedido
            'Se o controle de estoque for de reserva
            If objItemPV.iControleEstoque = PRODUTO_CONTROLE_RESERVA Then
                'Verifica quanto falta reservar
                lErro = CF("ItemPedido_QuantFaltaReservar", objItemPV, dQuantFaltaReservar)
                If lErro <> SUCESSO Then Error 42981
                'Acumula a quantidade que falta reservar total do pedido
                dPedidoFaltaReservar = dPedidoFaltaReservar + dQuantFaltaReservar
                dPedidoReservado = dPedidoReservado + objItemPV.dQuantReservada
            End If
        Next

        'Cria novo objBloqueioPV
        Set objBloqueioPV = New ClassBloqueioPV

        'Se a quantidade que falta resrvar do Pedido for positiva
        If Not E_Zero(dPedidoFaltaReservar) Then

            'Se a quantidade reservada do Pedido é positiva
            If dPedidoReservado > 0 Then
                'o bloqueio será do tipo parcial
                objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_PARCIAL
            'se não
            Else
                'o bloqueio será do tipo total
                objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_NAO_RESERVA
            End If

            'Verifica se o Tipo do objBloqueioPV é diferente do tipo do objBloqueioPV_BD
            If objBloqueioPV.iTipoDeBloqueio <> objBloqueioPV_BD.iTipoDeBloqueio Then
                'Preenche o objBloqueioPV
                iBloqueioEstoqueAdicionado = 1
                objBloqueioPV.sResponsavel = BLOQUEIO_AUTO_RESP
                objBloqueioPV.dtDataLib = DATA_NULA
                objBloqueioPV.iFilialEmpresa = giFilialEmpresa
                objBloqueioPV.dtData = gdtDataHoje
                objBloqueioPV.sCodUsuario = gsUsuario
                objBloqueioPV.lPedidoDeVendas = objPedidoVenda.lCodigo
                objBloqueioPV.sResponsavelLib = ""

                With objBloqueioPV
                    objPedidoVenda.colBloqueiosPV.Add .iFilialEmpresa, .lPedidoDeVendas, .iSequencial, .iTipoDeBloqueio, .sCodUsuario, .sResponsavel, .dtData, .sCodUsuario, .sResponsavelLib, .dtDataLib, ""
                End With
            End If
        
        End If

    End If

    BloqueioReserva_Adiciona2 = SUCESSO

    Exit Function

Erro_BloqueioReserva_Adiciona2:

    BloqueioReserva_Adiciona2 = Err

    Select Case Err

        Case 42981

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152529)

    End Select

    Exit Function

End Function

Function BloqueioReserva_Adiciona1(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Adiciona um bloqueio BLOQUEIO_PARCIAL ou BLOQUEIO_NAO_RESERVA ao Pedido de Vennda

Dim lErro As Long
Dim objBloqueioPV As ClassBloqueioPV
Dim bBloqueio_Total As Boolean
Dim dPedidoFaltaReservar As Double
Dim dPedidoReservado As Double
Dim objItemPV As ClassItemPedido
Dim dQuantFaltaReservar As Double
Dim iTipoBloqueio As Integer

On Error GoTo Erro_BloqueioReserva_Adiciona1

    bBloqueio_Total = False

    'Para cada bloqueio do Pedido
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        'Verifica se existe um bloqueio do tipo BLOQUEIO_TOTAL
        If objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
            bBloqueio_Total = True
            Exit For
        End If
    Next

    'Se não foi encontrado nenhum bloqueio total
    If Not bBloqueio_Total Then

        dPedidoFaltaReservar = 0
        dPedidoReservado = 0
        'Para cada item do Pedido
        For Each objItemPV In objPedidoVenda.colItensPedido
            'Se o controle estoque do produto for PRODUTO_CONTROLE_RESERVA
            If objItemPV.iControleEstoque = PRODUTO_CONTROLE_RESERVA Then
                
                'Calcula a quantidade que falta reservar
                lErro = CF("ItemPedido_QuantFaltaReservar", objItemPV, dQuantFaltaReservar)
                If lErro <> SUCESSO Then Error 42981
                
                'Acumula a quantidade que falta reserva ao Pedido
                dPedidoFaltaReservar = dPedidoFaltaReservar + dQuantFaltaReservar
                dPedidoReservado = dPedidoReservado + objItemPV.dQuantReservada
            End If
        Next

        'Se a quantidade que falta a reservar do pedido for positiva
        If dPedidoFaltaReservar > 0 Then
            'Se a quantidade reservada do Pedido for positiva
            If dPedidoReservado > 0 Then
                'O tipo de bloqueio vai ser parcial
                iTipoBloqueio = BLOQUEIO_PARCIAL
            'se não
            Else
                'O tipode bloqueio vai ser não reserva
                iTipoBloqueio = BLOQUEIO_NAO_RESERVA
            End If

            'Adiciona o bloqueio ao Pedido
            objPedidoVenda.colBloqueiosPV.Add giFilialEmpresa, objPedidoVenda.lCodigo, 0, iTipoBloqueio, gsUsuario, BLOQUEIO_AUTO_RESP, Date, "", "", DATA_NULA, ""

        End If

    End If

    BloqueioReserva_Adiciona1 = SUCESSO

    Exit Function

Erro_BloqueioReserva_Adiciona1:

    BloqueioReserva_Adiciona1 = Err

    Select Case Err

        Case 42981

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152530)

    End Select

    Exit Function

End Function

Function BloqueiosPV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere no BD os Bloqueios do Pedido de Venda

Dim lComando As Long
Dim lErro As Long
Dim objBloqueioPV As ClassBloqueioPV
Dim iSequencial As Integer

On Error GoTo Erro_BloqueiosPV_Grava

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 42982

     'Para cada Bloqueio do Pedido
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV

        'Gera o sequencial automático para o Bloqueio
        lErro = CF("BloqueioPV_Automatico", objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, iSequencial)
        If lErro <> SUCESSO Then Error 42983

        'Atribui Sequencial ANTES de INSERIR
        objBloqueioPV.iSequencial = iSequencial

        'Insere o bloqueio na tabela de Bloqueios de Pedido de Venda
        lErro = Comando_Executar(lComando, "INSERT INTO BloqueiosPV(FilialEmpresa,PedidoDeVenda,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib,DataLib,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial, objBloqueioPV.iTipoDeBloqueio, objBloqueioPV.sCodUsuario, objBloqueioPV.sResponsavel, objBloqueioPV.dtData, objBloqueioPV.sCodUsuarioLib, objBloqueioPV.sResponsavelLib, objBloqueioPV.dtDataLib, objBloqueioPV.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then Error 42984

    Next
     
    'Fecha o comando
    Call Comando_Fechar(lComando)

    BloqueiosPV_Grava = SUCESSO

    Exit Function

Erro_BloqueiosPV_Grava:

    BloqueiosPV_Grava = Err

    Select Case Err

        Case 42982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 42983

        Case 42984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BLOQUEIOPV", Err, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, iSequencial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152531)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ReservasPV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Exclui as Reservas do pedido de Venda passado
'CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_ReservasPV_Exclui

    'Lê as reservas do Pedido
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM Reserva WHERE FilialEmpresa = ? AND TipoDoc IN (?,?) AND DocOrigem = ?", 0, iFilialEmpresa, objPedidoVenda.iFilialEmpresa, PEDIDO_DE_VENDA, PEDIDO_DE_VENDA_GRADE, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46063

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 46064

    'Para cada reserva lida
    Do While lErro = AD_SQL_SUCESSO
        
        'Exclui a Reserva do Pedido da tabela de Reservas
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Reserva", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 46065

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46066

    Loop

    ReservasPV_Exclui = SUCESSO

    Exit Function

Erro_ReservasPV_Exclui:

    ReservasPV_Exclui = Err

    Select Case Err

        Case 46063, 46064, 46066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", Err)

        Case 46065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RESERVASPV", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152532)

    End Select

    Exit Function

End Function

Function ItensPV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long
'Exclui os itens presentes na coleção de itens do Pedido de vanda passado

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_ItensPV_Exclui

    'Para cada item na coleção
    For Each objItemPV In objPedidoVenda.colItensPedido

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ItensPedidoDeVenda WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 46080

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46081
        If lErro = AD_SQL_SEM_DADOS Then gError 46082 'Não encontrou

        'Exclui o item da tabela de itens de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ItensPedidoDeVenda ", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 46083

        lErro = CF("ItensPVGrade_Exclui", objItemPV)
        If lErro <> SUCESSO Then gError 86351
        
        objItemPV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMPV
        objItemPV.objInfoAdicDocItem.lNumIntDocItem = objItemPV.lNumIntDoc

        lErro = CF("InfoAdicDocItem_Exclui", objItemPV.objInfoAdicDocItem)
        If lErro <> SUCESSO Then gError 86351

        lErro = CF("ItemPV_ExcluiTributacao", objItemPV, lComando2, lComando3)
        If lErro <> SUCESSO Then gError 46176

    Next

    ItensPV_Exclui = SUCESSO

    Exit Function

Erro_ItensPV_Exclui:

    ItensPV_Exclui = gErr

    Select Case gErr

        Case 46080, 46081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 46082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_ENCONTRADO", gErr, objPedidoVenda.lNumIntDoc)

        Case 46083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDIDODEVENDA", gErr, objPedidoVenda.lCodigo)

        Case 46176, 86351

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152533)

    End Select

    Exit Function

End Function

Function PedidoVendaBaixado_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda, faz o lock e o insere na tabela de Pedidos baixados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tPedidoVenda As typePedidoVenda

On Error GoTo Erro_PedidoVendaBaixado_Grava

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46098

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 46099

    tPedidoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tPedidoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tPedidoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tPedidoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
    tPedidoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
    tPedidoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
    tPedidoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    tPedidoVenda.sOBS = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPedidoVenda.sEmitente = String(STRING_USUARIO_CODIGO, 0)
    tPedidoVenda.sUsuarioUltAlteracao = String(STRING_USUARIO_CODIGO, 0)
    
    'Lê o Pedido de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,VolumeNumero,FreteRespons,ValorProdutos, CodTranspRedesp, DetPagFrete, DataEntrega, Moeda, TaxaMoeda, VolumeTotal,PedidoRepresentante, OBS, Andamento,ValorDescontoTit,ValorDescontoItens,ValorItens, CodigoBase, Parc, Emitente, UsuarioUltAlteracao, DataInclusao, DataAlteracao, HoraInclusao, HoraAlteracao FROM PedidosDeVenda WHERE FilialEmpresa = ? AND Codigo = ?", _
    0, tPedidoVenda.iFilialEmpresaFaturamento, tPedidoVenda.lCliente, tPedidoVenda.iFilial, tPedidoVenda.iFilialEntrega, tPedidoVenda.iCodTransportadora, tPedidoVenda.iCondicaoPagto, tPedidoVenda.dPercAcrescFinanceiro, tPedidoVenda.dtDataEmissao, tPedidoVenda.dtDataReferencia, tPedidoVenda.sMensagemNota, _
        tPedidoVenda.sNaturezaOp, tPedidoVenda.dValorTotal, tPedidoVenda.dValorFrete, tPedidoVenda.dValorDesconto, tPedidoVenda.dValorSeguro, tPedidoVenda.dValorOutrasDespesas, tPedidoVenda.sPedidoCliente, tPedidoVenda.iCanalVenda, tPedidoVenda.iTabelaPreco, tPedidoVenda.iProxSeqBloqueio, tPedidoVenda.iFaturaIntegral, _
        tPedidoVenda.iComissaoAutomatica, tPedidoVenda.iCobrancaAutomatica, tPedidoVenda.lNumIntDoc, tPedidoVenda.sPlaca, tPedidoVenda.sPlacaUF, tPedidoVenda.lVolumeQuant, tPedidoVenda.lVolumeEspecie, tPedidoVenda.lVolumeMarca, tPedidoVenda.dPesoLiq, tPedidoVenda.dPesoBruto, tPedidoVenda.sVolumeNumero, tPedidoVenda.iFreteRespons, tPedidoVenda.dValorProdutos, tPedidoVenda.iCodTranspRedesp, tPedidoVenda.iDetPagFrete, tPedidoVenda.dtDataEntrega, tPedidoVenda.iMoeda, tPedidoVenda.dTaxaMoeda, tPedidoVenda.dVolumeTotal, tPedidoVenda.sPedidoRepresentante, tPedidoVenda.sOBS, tPedidoVenda.iAndamento, _
        tPedidoVenda.dValorDescontoTit, tPedidoVenda.dValorDescontoItens, tPedidoVenda.dValorItens, tPedidoVenda.lCodigoBase, tPedidoVenda.iParc, tPedidoVenda.sEmitente, tPedidoVenda.sUsuarioUltAlteracao, tPedidoVenda.dtDataInclusao, tPedidoVenda.dtDataAlteracao, tPedidoVenda.dHoraInclusao, tPedidoVenda.dHoraAlteracao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46099

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46100
    If lErro = AD_SQL_SEM_DADOS Then Error 46101 'Não encontrou

    'Faz o "lock" no Pedido de Venda
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 46102

    objPedidoVenda.lNumIntDoc = tPedidoVenda.lNumIntDoc
    
    With tPedidoVenda
        tPedidoVenda.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
        tPedidoVenda.lCodigo = objPedidoVenda.lCodigo
        
        'Insere o Pedido na tabela de Pedidos Baixados
        lErro = Comando_Executar(lComando1, "INSERT INTO PedidosDeVendaBaixados (FilialEmpresa, Codigo,FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,ValorProdutos,VolumeNumero,FreteRespons, CodTranspRedesp, DetPagFrete, DataEntrega,Moeda, TaxaMoeda, VolumeTotal,PedidoRepresentante, OBS, Andamento,ValorDescontoTit,ValorDescontoItens,ValorItens, CodigoBase, Parc, Emitente, UsuarioUltAlteracao, DataInclusao, DataAlteracao, HoraInclusao, HoraAlteracao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .iFilialEmpresa, .lCodigo, .iFilialEmpresaFaturamento, .lCliente, .iFilial, .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataEmissao, .dtDataReferencia, .sMensagemNota, .sNaturezaOp, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iProxSeqBloqueio, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .lNumIntDoc, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .dValorProdutos, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .dtDataEntrega, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .sPedidoRepresentante, .sOBS, .iAndamento, .dValorDescontoTit, .dValorDescontoItens, .dValorItens, .lCodigoBase, .iParc, .sEmitente, .sUsuarioUltAlteracao, .dtDataInclusao, .dtDataAlteracao, .dHoraInclusao, .dHoraAlteracao)
        If lErro <> AD_SQL_SUCESSO Then Error 46103
    
    End With

    objPedidoVenda.lCodigo = tPedidoVenda.lCodigo
    objPedidoVenda.iFilialEmpresa = tPedidoVenda.iFilialEmpresa
    
    lErro = CF("PV_Grava_Customizado", objPedidoVenda)
    If lErro <> SUCESSO Then gError 180412
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PedidoVendaBaixado_Grava = SUCESSO

    Exit Function

Erro_PedidoVendaBaixado_Grava:

    PedidoVendaBaixado_Grava = Err

    Select Case Err

        Case 46098, 46099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46099, 46100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedidoVenda.lCodigo)

        Case 46101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoVenda.lCodigo)

        Case 46102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objPedidoVenda.lCodigo)

        Case 46103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDODEVENDABAIXADO", Err, objPedidoVenda.lCodigo)

        Case 180412

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152534)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ItensPVBaixados_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda, iFaturaTudo As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItemPV As typeItemPedido
Dim iTemItens As Integer 'Sinaliza se o Pedido já vem carregado com seus itens

On Error GoTo Erro_ItensPVBaixados_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 46104

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 46105

    tItemPV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPV.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemPV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ?", 0, _
        tItemPV.lNumIntDoc, tItemPV.sProduto, tItemPV.dQuantidade, tItemPV.dPrecoUnitario, tItemPV.dPrecoTotal, tItemPV.iClasseUM, tItemPV.sUnidadeMed, tItemPV.dValorDesconto, tItemPV.dtDataEntrega, tItemPV.sDescricao, tItemPV.sLote, tItemPV.iStatus, tItemPV.dValorAbatComissao, tItemPV.dQuantCancelada, tItemPV.dQuantReservada, tItemPV.dQuantFaturada, tItemPV.dQuantOP, tItemPV.dQuantSC, tItemPV.dPercentMaisReceb, tItemPV.dPercentMenosReceb, tItemPV.iRecebForaFaixa, tItemPV.dQuantFaturadaAMais, tItemPV.iPrioridade, tItemPV.dComissao, tItemPV.iTabelaPreco, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 46106

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46107

    iTemItens = False
    If objPedidoVenda.colItensPedido.Count > 0 Then iTemItens = True

    Do While lErro = AD_SQL_SUCESSO

        With tItemPV
            
            If Not iTemItens Then objPedidoVenda.colItensPedido.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantFaturada, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc
            
            .lCodPedido = objPedidoVenda.lCodigo
            .iFilialEmpresa = objPedidoVenda.iFilialEmpresa
            .dQuantReservada = 0

            lErro = Comando_Executar(lComando1, "INSERT INTO ItensPedidoDeVendaBaixados (FilialEmpresa, CodPedido, NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .iFilialEmpresa, .lCodPedido, .lNumIntDoc, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .dQuantFaturadaAMais, .iPrioridade, .dComissao, .iTabelaPreco)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46108

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46109

        End With

    Loop

'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("PVItens_Grava_Customizado", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 177047
'    '############################################

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensPVBaixados_Grava = SUCESSO

    Exit Function

Erro_ItensPVBaixados_Grava:

    ItensPVBaixados_Grava = gErr

    Select Case gErr

        Case 46104, 46105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 46106, 46107, 46109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA2", gErr, objPedidoVenda.lCodigo)

        Case 46108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDIDODEVENDABAIXADOS", gErr, objPedidoVenda.lCodigo)

        Case 177047

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152535)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ComissoesPV_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Faz a baixa das Comissões do Pedido de Venda passado

Dim lErro As Long
Dim tComissoesPV As typeComissaoPedVenda

On Error GoTo Erro_ComissoesPV_Baixa

    'Busca as Comissões do Pedido
    lErro = Comando_ExecutarPos(lComando, "SELECT CodVendedor,ValorBase,Percentual,Valor,PercentualEmissao,ValorEmissao, Indireta, Seq FROM ComissoesPedVendas WHERE FilialEmpresa= ? AND PedidoDeVendas = ?", 0, tComissoesPV.iCodVendedor, tComissoesPV.dValorBase, tComissoesPV.dPercentual, tComissoesPV.dValor, tComissoesPV.dPercentualEmissao, tComissoesPV.dValorEmissao, tComissoesPV.iIndireta, tComissoesPV.iSeq, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46110

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46111

    With tComissoesPV
        'Para cada comissão encontrada
        Do While lErro = AD_SQL_SUCESSO

            tComissoesPV.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
            tComissoesPV.lPedidoDeVendas = objPedidoVenda.lCodigo

            'Insere na tabela de Comissões baixadas
            lErro = Comando_Executar(lComando1, " INSERT INTO ComissoesPedVendasBaixados (FilialEmpresa,PedidoDeVendas,CodVendedor,ValorBase,Percentual,Valor,PercentualEmissao,ValorEmissao, Indireta, Seq) VALUES (?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .lPedidoDeVendas, .iCodVendedor, .dValorBase, .dPercentual, .dValor, .dPercentualEmissao, .dValorEmissao, .iIndireta, .iSeq)
            If lErro <> AD_SQL_SUCESSO Then Error 46112
            'Exclui da tabela de comissões não baixadas
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ComissoesPedVendas", lComando)
            If lErro <> SUCESSO Then Error 46113

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 46114
        Loop

    End With

    ComissoesPV_Baixa = SUCESSO

    Exit Function

Erro_ComissoesPV_Baixa:

    ComissoesPV_Baixa = Err

    Select Case Err

        Case 46110, 46111, 46114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDVENDAS", Err)

        Case 46112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESPV_BAIXADOS", Err, objPedidoVenda.lCodigo)

        Case 46113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESPEDVENDAS", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152536)

    End Select

    Exit Function

End Function

Function ParcelasPV_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Faz a baixa das Parcelas do Pedido de Venda

Dim lErro As Long
Dim tParcelaPV As typeParcelaPedidoVenda

On Error GoTo Erro_ParcelasPV_Baixa

    tParcelaPV.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaPV.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaPV.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaPV.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)

    'Lê as Parcelas do Pedido
    lErro = Comando_ExecutarPos(lComando, "SELECT NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao FROM ParcelasPedidoDeVenda WHERE FilialEmpresa=? AND PedidodeVenda=?", 0, _
    tParcelaPV.iNumParcela, tParcelaPV.dtDataVencimento, tParcelaPV.dValor, tParcelaPV.iDesconto1Codigo, tParcelaPV.dtDesconto1Ate, tParcelaPV.dDesconto1Valor, tParcelaPV.iDesconto2Codigo, tParcelaPV.dtDesconto2Ate, tParcelaPV.dDesconto2Valor, tParcelaPV.iDesconto3Codigo, tParcelaPV.dtDesconto3Ate, tParcelaPV.dDesconto3Valor, _
    tParcelaPV.iTipoPagto, tParcelaPV.iCodConta, tParcelaPV.dtDataCredito, tParcelaPV.dtDataEmissaoCheque, tParcelaPV.iBancoCheque, tParcelaPV.sAgenciaCheque, tParcelaPV.sContaCorrenteCheque, tParcelaPV.lNumeroCheque, tParcelaPV.dtDataDepositoCheque, tParcelaPV.iAdmMeioPagto, tParcelaPV.iParcelamento, tParcelaPV.sNumeroCartao, tParcelaPV.dtValidadeCartao, tParcelaPV.sAprovacaoCartao, tParcelaPV.dtDataTransacaoCartao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46115

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46116

    With tParcelaPV

        Do While lErro = AD_SQL_SUCESSO

            'Insere na tabela de Parcelas baixadas
            lErro = Comando_Executar(lComando1, "INSERT INTO ParcelasPedidoDeVendaBaixado (FilialEmpresa,PedidodeVenda,NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iNumParcela, .dtDataVencimento, .dValor, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, _
            .iTipoPagto, .iCodConta, .dtDataCredito, .dtDataEmissaoCheque, .iBancoCheque, .sAgenciaCheque, .sContaCorrenteCheque, .lNumeroCheque, .dtDataDepositoCheque, .iAdmMeioPagto, .iParcelamento, .sNumeroCartao, .dtValidadeCartao, .sAprovacaoCartao, .dtDataTransacaoCartao)
            If lErro <> AD_SQL_SUCESSO Then Error 46117
            'Execlui da tabela de Parcelas não baixadas
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ParcelasPedidoDeVenda", lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 46118

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 46119

        Loop

    End With

    ParcelasPV_Baixa = SUCESSO
    
    Exit Function

Erro_ParcelasPV_Baixa:

    ParcelasPV_Baixa = Err

    Select Case Err

        Case 46115, 16116, 46119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPEDIDODEVENDA", Err, objPedidoVenda.lCodigo)

        Case 46117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELASPV_BAIXADOS", Err, objPedidoVenda.lCodigo)

        Case 46118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASPV", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152537)

    End Select

    Exit Function

End Function

Function BloqueiosPV_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Faz a Baixa dos Bloqueios do Pedido de Venda passado

Dim lErro As Long
Dim tBloqueioPV As typeBloqueioPV

On Error GoTo Erro_BloqueiosPV_Baixa

    tBloqueioPV.sCodUsuario = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sCodUsuarioLib = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sResponsavel = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
    tBloqueioPV.sResponsavelLib = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
    tBloqueioPV.sObservacao = String(STRING_BLOQUEIOSPV_OBSERVACAO, 0)

    'Lê os Bloqueios
    lErro = Comando_ExecutarPos(lComando, "SELECT Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib,DataLib, Observacao FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ?", 0, tBloqueioPV.iSequencial, tBloqueioPV.iTipoDeBloqueio, tBloqueioPV.sCodUsuario, tBloqueioPV.sResponsavel, tBloqueioPV.dtData, tBloqueioPV.sCodUsuarioLib, tBloqueioPV.sResponsavelLib, tBloqueioPV.dtDataLib, tBloqueioPV.sObservacao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46120

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46122

    With tBloqueioPV

        'Para cada  Bloqueio encontrado
        Do While lErro = AD_SQL_SUCESSO

            .iFilialEmpresa = objPedidoVenda.iFilialEmpresa
            .lPedidoDeVendas = objPedidoVenda.lCodigo
            
            'Insere na tabela de Bloqueios Baixados
            If tBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_PARCIAL And tBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_NAO_RESERVA Then
                lErro = Comando_Executar(lComando1, "INSERT INTO BloqueiosPVBaixados (FilialEmpresa,PedidoDeVenda,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib,DataLib,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", tBloqueioPV.iFilialEmpresa, tBloqueioPV.lPedidoDeVendas, tBloqueioPV.iSequencial, tBloqueioPV.iTipoDeBloqueio, tBloqueioPV.sCodUsuario, tBloqueioPV.sResponsavel, tBloqueioPV.dtData, tBloqueioPV.sCodUsuarioLib, tBloqueioPV.sResponsavelLib, tBloqueioPV.dtDataLib, tBloqueioPV.sObservacao)
                If lErro <> AD_SQL_SUCESSO Then Error 46123
            End If
            
            'Exclui da Tabela de BloqueiosPV não baixados
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM BloqueiosPV", lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 46124

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46125

        Loop

    End With

    BloqueiosPV_Baixa = SUCESSO

    Exit Function

Erro_BloqueiosPV_Baixa:

    Select Case Err

        Case 46120, 46122, 46125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", Err, objPedidoVenda.lCodigo)

        Case 46123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_BLOQUEIOSPVBAIXADOS", Err, objPedidoVenda.lCodigo)

        Case 46124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOSPV", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152538)

    End Select

    Exit Function

End Function

Function PedidoVenda_Exclui_BD(ByVal objPedVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long) As Long
'Exclui o Pedido de Venda da tabela de Pedido de Vendas

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_PedidoVenda_Exclui_BD

    'Busca o Pedido no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM PedidosDeVenda WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46126

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46127
    If lErro = AD_SQL_SEM_DADOS Then Error 46128 'Não encontrou

    'Exclui o Pedido da Tabela de Pedidos
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PedidosDeVenda ", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 46129

    lErro = CF("PedidoVenda_ExcluiTributacao", objPedVenda, lComando2, lComando3, lComando4, lComando5)

    PedidoVenda_Exclui_BD = SUCESSO

    Exit Function

Erro_PedidoVenda_Exclui_BD:

    PedidoVenda_Exclui_BD = Err

    Select Case Err

        Case 46126, 46127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedVenda.lCodigo)

        Case 46128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedVenda.lCodigo)

        Case 46129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDODEVENDA", Err, objPedVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152539)

    End Select

    Exit Function

End Function

Function PedidoVenda_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui o Pedido de Venda passado

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PedidoVenda_Exclui

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 46140

    lErro = CF("PedidoVenda_Exclui_Trans", objPedidoVenda)
    If lErro <> SUCESSO Then gError 197628
    
    lErro = CF("PedidoVendaParc_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 46157

    lErro = CF("WorkFlow_Trata_Transacao", "FAT", objPedidoVenda.objTela, gsUsuario, objPedidoVenda, 1009)
    If lErro <> SUCESSO Then gError 197628
    
    PedidoVenda_Exclui = SUCESSO

    Exit Function

Erro_PedidoVenda_Exclui:

    PedidoVenda_Exclui = gErr

    Select Case gErr

        Case 46140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197628

        Case 46157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152540)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function PedidoVenda_Exclui_Trans(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui o Pedido de Venda passado

Dim lErro As Long
Dim alComando(0 To 14) As Long
Dim iIndice As Integer
Dim objItemPV As ClassItemPedido
Dim objItemOP As New ClassItemOP
Dim objMovEstoque As New ClassMovEstoque
Dim colobjRC As New Collection
Dim objRC As ClassRequisicaoCompras
Dim sRequisicoes As String, lCodigoPC As Long

On Error GoTo Erro_PedidoVenda_Exclui_Trans

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 46141
    Next

    'Faz lock no Pedido
    lErro = CF("PedidoVenda_Lock", objPedidoVenda, alComando(0))
    If lErro <> SUCESSO And lErro <> 42988 Then gError 46142

    If lErro <> SUCESSO Then gError 46143

    'Lê e faz lock nos itens do Pedido
    lErro = CF("ItensPV_Le_Lock", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46144
        
    lErro = ProdutoFilial_Subtrai_QuantPedida(objPedidoVenda, objPedidoVenda.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 51408
    
    'Para todos os itens de PedidoVenda
    For Each objItemPV In objPedidoVenda.colItensPedido

        'Verifica se item tem Quantidade Faturada positiva
        If objItemPV.dQuantFaturada > 0 Then gError 46145

        'Verifica se item está vinculado a item de ordem de produção
        lErro = CF("ItemOP_Le_ItemPV", objItemOP, objItemPV)
        If lErro <> SUCESSO And lErro <> 46074 Then gError 46146
        If lErro = SUCESSO Then gError 46147

    Next
    
    '?????? Fernando fazer chama funções, esta
    'função está neste módulo, passar para ClassPVSelect
    'William 05/04/01
    'Verifica se existe RequisicaoCompra vinculada ao Pedido de Venda
    lErro = PedidoVenda_Busca_ReqCompra(objPedidoVenda)
    If lErro <> SUCESSO Then gError 87531

    'Verifica se existe PedidoCompra vinculada ao Pedido de Venda
    lErro = CF("PedidoVenda_Busca_PedidoCompra", objPedidoVenda, lCodigoPC)
    If lErro <> SUCESSO And lErro <> 178835 Then gError 178831

    If lErro = SUCESSO Then gError 178837

    'Verifica se existe requisicao associado ao pedido de venda
    lErro = CF("ItemPVItemRC_Le_CodigoPV", objPedidoVenda, colobjRC)
    If lErro <> SUCESSO Then gError 178846

    If colobjRC.Count > 0 Then
        For Each objRC In colobjRC
            sRequisicoes = sRequisicoes & "Filial=" & Str(objRC.iFilialEmpresa) & " Requisicao=" & Str(objRC.lCodigo) & Chr(10)
        Next
        gError 178847
    End If
    
    'Lê as Reservas do pedido
    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46148
    
    'Cria um movimento de estoque para a exclusão
    lErro = MovEstoqueExclusao_Cria(objMovEstoque, objPedidoVenda)
    If lErro <> SUCESSO Then gError 46149

    If objMovEstoque.colItens.Count > 0 Then
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 46150
    End If

    'Exclui as Reservas do Pedido
    lErro = CF("ReservasPV_Exclui", objPedidoVenda, alComando(1), alComando(2))
    If lErro <> SUCESSO Then gError 46151

    'Exclui os bloqueios do Pedido
    lErro = CF("BloqueiosPV_Exclui", objPedidoVenda, alComando(3), alComando(4))
    If lErro <> SUCESSO Then gError 46152

    'Exclui as comissões do Pedido
    lErro = CF("ComissoesPV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46153

    'Exclui as Parcelas do Pedido
    lErro = CF("ParcelasPV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46154

    'Exclui as embalagens do pedido no BD
    lErro = PedidoVenda_EmbalagensExclui(objPedidoVenda)
    If lErro <> SUCESSO Then gError 42957
    
    lErro = CF("ItensPVEntrega_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 183216
    
    'Exclui os itens do Pedido
    lErro = CF("ItensPV_Exclui", objPedidoVenda, alComando(5), alComando(6), alComando(9), alComando(10))
    If lErro <> SUCESSO Then gError 46155

    'Caso exista vínculo, desvincula o Orçamento de Venda do Pedido a ser excluido.
    lErro = PedidoVenda_Desvincula_Orcamento(objPedidoVenda) 'Por Leo em 29/04/02
    If lErro <> SUCESSO Then gError 94484
    
    'Exclui o Pedido de Venda do BD
    lErro = CF("PedidoVenda_Exclui_BD", objPedidoVenda, alComando(7), alComando(8), alComando(11), alComando(12), alComando(13), alComando(14))
    If lErro <> SUCESSO Then gError 46156
    
    'Diminui a Estatistica para os Produtos e Data de Emissao do Pedido de Venda Gerados em SldDiaFat e SldMesFat
    lErro = PedVenda_Exclui_Estatisticas(objPedidoVenda)
    If lErro <> SUCESSO Then gError 64425
    
    lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_PV, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 64425
    
    lErro = CF("InfoAdicionais_Exclui", TIPODOC_INFOADIC_PV, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 64425
    
'    lErro = CF("WorkFlow_Trata_Transacao", "FAT", objPedidoVenda.objTela, gsUsuario, objPedidoVenda)
'    If lErro <> SUCESSO Then gError 178238
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoVenda_Exclui_Trans = SUCESSO

    Exit Function

Erro_PedidoVenda_Exclui_Trans:

    PedidoVenda_Exclui_Trans = gErr

    Select Case gErr

        Case 46141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 46142, 46144, 46146, 46148 To 46156, 51408, 64425, 178238, 178831, 178846, 183216

        Case 46143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case 46145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_FATURADO", gErr, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc, objItemPV.dQuantFaturada)

        Case 46147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_VINCULADO_ITEM_PV", gErr, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc, objItemOP.lNumIntDoc)

        Case 87531, 94484
        
        Case 178837
            Call Rotina_Erro(vbOKOnly, "ERRO_PV_VINCULADO_PEDIDOCOMPRA", gErr, lCodigoPC)
        
        Case 178847
            Call Rotina_Erro(vbOKOnly, "ERRO_PV_VINCULADO_REQCOMPRA", gErr, sRequisicoes)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152540)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueiosPV_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Exclui os Bloqueios do Pedido de Venda passado

Dim lErro As Long
Dim iSequencial As Integer

On Error GoTo Erro_BloqueiosPV_Exclui

    'Busca os Bloqueios do Pedido passado
    lErro = Comando_ExecutarPos(lComando, "SELECT Sequencial FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ?", 0, iSequencial, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46158

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46159

    'Para  cadabloqueio encontrado
    Do While lErro = AD_SQL_SUCESSO
        'Exclui o bloqueio do BD
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BloqueiosPV", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 46160

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46161

    Loop

    BloqueiosPV_Exclui = SUCESSO

    Exit Function

Erro_BloqueiosPV_Exclui:

    BloqueiosPV_Exclui = Err

    Select Case Err

        Case 46158, 46159, 46161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", Err, objPedidoVenda.lCodigo)

        Case 46160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOSPV", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152541)

    End Select

    Exit Function

End Function

Function ItemPV_GravaTributacao(ByVal objItemPV As ClassItemPedido, ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ItemPV_GravaTributacao

'    With objItemPV.objTributacaoItemPV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoItemPV (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemPV.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then Error 41540

    objItemPV.objTributacao.lNumIntDoc = objItemPV.lNumIntDocPV
    objItemPV.objTributacao.lNumIntDocItem = objItemPV.lNumIntDoc
    objItemPV.objTributacao.iItem = objItemPV.iItem
    objItemPV.objTributacao.dDescontoGrid = objItemPV.dValorDesconto
    objItemPV.objTributacao.dPrecoUnitario = objItemPV.dPrecoUnitario
    'objItemPV.objTributacao.dPrecoUnitarioMoeda = objItemPV.dPrecoUnitario
    objItemPV.objTributacao.dQuantidade = objItemPV.dQuantidade - objItemPV.dQuantCancelada
    objItemPV.objTributacao.sProduto = objItemPV.sProduto
    objItemPV.objTributacao.sProdutoDescricao = objItemPV.sDescricao
    objItemPV.objTributacao.sUnidadeMed = objItemPV.sUnidadeMed
    objItemPV.objTributacao.iTipoDoc = TIPODOC_TRIB_PV

    lErro = CF("TributacaoDocItem_Grava", objItemPV.objTributacao)
    If lErro <> SUCESSO Then Error 41540

    ItemPV_GravaTributacao = SUCESSO

    Exit Function

Erro_ItemPV_GravaTributacao:

    ItemPV_GravaTributacao = Err

    Select Case Err

        Case 41540
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMPV", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152542)

    End Select

    Exit Function

End Function

Function PedidoVenda_GravaTributacao(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_PedidoVenda_GravaTributacao

'    With objPedidoVenda.objTributacaoPV
'
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoPV (FilialEmpresa, CodPedido, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'           .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then Error 41534
'
'    'inserir os complementos
'    lErro = CF("PedidoVenda_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then Error 41535
'
'    lErro = CF("PedidoVenda_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then Error 41536
'
'    lErro = CF("PedidoVenda_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then Error 41537
'
'    lErro = CF("PedidoVenda_GravaTributacaoCompl", objPedidoVenda, objPedidoVenda.objTributacaoPV.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then Error 41538

    objPedidoVenda.objTributacao.lNumIntDoc = objPedidoVenda.lNumIntDoc
    objPedidoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_PV

    lErro = CF("TributacaoDoc_Grava", objPedidoVenda.objTributacao)
    If lErro <> SUCESSO Then Error 41538

    PedidoVenda_GravaTributacao = SUCESSO

    Exit Function

Erro_PedidoVenda_GravaTributacao:

    PedidoVenda_GravaTributacao = Err

    Select Case Err
'
'        Case 41534
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBPV", Err)

        Case 41535 To 41538

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152543)

    End Select

    Exit Function

End Function

'Function PedidoVenda_GravaTributacaoCompl(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objTribComplPV As ClassTributacaoItemPV, ByVal lComando As Long) As Long
'
'Dim lErro As Long, iItem As Integer
'
'On Error GoTo Erro_PedidoVenda_GravaTributacaoCompl
'
'    iItem = objTribComplPV.lNumIntDoc
'
'    With objTribComplPV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplPV (FilialEmpresa, CodPedido, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then Error 41539
'
'    PedidoVenda_GravaTributacaoCompl = SUCESSO
'
'    Exit Function
'
'Erro_PedidoVenda_GravaTributacaoCompl:
'
'    PedidoVenda_GravaTributacaoCompl = Err
'
'    Select Case Err
'
'        Case 41539
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152544)
'
'    End Select
'
'    Exit Function
'
'End Function

Function PedidoVenda_ExcluiTributacao(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long

Dim lErro As Long
Dim iItem As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PedidoVenda_ExcluiTributacao
'
'    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM TributacaoComplPV WHERE FilialEmpresa = ? AND CodPedido = ? ", 0, iItem, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then Error 46164
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46165
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoComplPV", lComando)
'        If lErro <> AD_SQL_SUCESSO Then Error 46166
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46167
'
'    Loop
'
'    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa FROM TributacaoPV WHERE FilialEmpresa = ? AND CodPedido = ? ", 0, iFilialEmpresa, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then Error 46172
'
'    lErro = Comando_BuscarPrimeiro(lComando2)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46173
'    If lErro <> AD_SQL_SUCESSO Then Error 46174
'
'    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM TributacaoPV", lComando2)
'    If lErro <> AD_SQL_SUCESSO Then Error 46175

    lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_PV, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 46174

    PedidoVenda_ExcluiTributacao = SUCESSO

    Exit Function

Erro_PedidoVenda_ExcluiTributacao:

    PedidoVenda_ExcluiTributacao = Err

    Select Case Err

'        Case 46164, 46165, 46167
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLPEDIDO", Err)
'
'        Case 46166
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLPEDIDO", Err, objPedidoVenda.lCodigo)
'
'        Case 46172, 46173
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBPEDIDO", Err)

        Case 46174
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBPEDIDO_NAO_ENCONTRADA", Err, objPedidoVenda.lCodigo)

        Case 46175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBPEDIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152545)

    End Select

    Exit Function

End Function

Function ItemPV_ExcluiTributacao(ByVal objItemPV As ClassItemPedido, ByVal lComando As Long, ByVal lComando1 As Long) As Long

Dim lErro As Long
Dim iNaturezaOpManual As Integer

On Error GoTo Erro_ItemPV_ExcluiTributacao

'    lErro = Comando_ExecutarPos(lComando, "SELECT NaturezaOpManual FROM TributacaoItemPV WHERE NumIntDoc = ? ", 0, iNaturezaOpManual, objItemPV.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then Error 46168
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46169
'    If lErro <> AD_SQL_SUCESSO Then Error 46170
'
'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoItemPV", lComando)
'    If lErro <> AD_SQL_SUCESSO Then Error 46171

    lErro = CF("TributacaoDocItem_Exclui", TIPODOC_TRIB_PV, objItemPV.lNumIntDoc)
    If lErro <> SUCESSO Then gError 46170

    ItemPV_ExcluiTributacao = SUCESSO

    Exit Function

Erro_ItemPV_ExcluiTributacao:

    ItemPV_ExcluiTributacao = gErr

    Select Case gErr
'
'        Case 46168, 46169, 46171
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMPEDIDO", Err)

        Case 46170
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMPEDIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152546)

    End Select

    Exit Function

End Function

Function ComissoesPV_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza as comissões do Pedido de Venda passado

Dim lErro As Long

On Error GoTo Erro_ComissoesPV_Altera

    'Exclui do BD as comissões do Pedido
    lErro = CF("ComissoesPV_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then Error 46006

    'Inclui no BD as comissões em objPedidoVenda
    lErro = CF("ComissoesPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then Error 46013

    ComissoesPV_Altera = SUCESSO

    Exit Function

Erro_ComissoesPV_Altera:

    ComissoesPV_Altera = Err

    Select Case Err

        Case 46006, 46013

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152547)

    End Select

    Exit Function

End Function

Function ParcelasPV_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza as pracelas do Pedido de Venda Passdados

Dim lErro As Long

On Error GoTo Erro_ParcelasPV_Altera

    'Exclui as Parcelas ligadas ao Pedido de Venda
    lErro = CF("ParcelasPV_Exclui", objPedidoVenda)
    If lErro <> AD_SQL_SUCESSO Then Error 46014

    'Inclui as novas parcelas do pedido de Venda
    lErro = CF("ParcelasPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then Error 46015

    ParcelasPV_Altera = SUCESSO

    Exit Function

Erro_ParcelasPV_Altera:

    ParcelasPV_Altera = Err

    Select Case Err

        Case 46014, 46015

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152548)

    End Select

    Exit Function

End Function

Function ReservasPV_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Atualiza as Reservas do Pedido de Venda passado

Dim lErro As Long

    'Exclui as reservas do Pedido de Venda
    lErro = CF("ReservasPV_Exclui", objPedidoVenda, lComando, lComando1)
    If lErro <> SUCESSO Then Error 46053

    'Grava as Reservas do Pedido de Venda
    lErro = CF("ReservasPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then Error 46067

    ReservasPV_Altera = SUCESSO

    Exit Function

Erro_ReservasPV_Altera:

    ReservasPV_Altera = Err

    Select Case Err

        Case 46053, 46067

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152549)

    End Select

    Exit Function

End Function

Private Function PedVenda_QuantNaoCancelada(ByVal objPedidoVenda As ClassPedidoDeVenda, dQuantNaoCancelada As Double) As Long
'Retorna a quantidade não cancelada em todos o Pedido

Dim lErro As Long
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_PedVenda_QuantNaoCancelada

    'Zera o acumulador
    dQuantNaoCancelada = 0

    'Para cada item acumula a Quantidade não cancelada
    For Each objItemPV In objPedidoVenda.colItensPedido
        dQuantNaoCancelada = dQuantNaoCancelada + objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada
    Next

    PedVenda_QuantNaoCancelada = SUCESSO

    Exit Function

Erro_PedVenda_QuantNaoCancelada:

    PedVenda_QuantNaoCancelada = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152550)

    End Select

    Exit Function

End Function

Private Function ItensPV_QuantFaturadaAlterada(ByVal objPedVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_BD As ClassPedidoDeVenda) As Long
'Para cada item verifica se a quantidade Faturada é a mesma que a do Banco de Dados

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objItemPV_BD As ClassItemPedido
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_ItensPV_QuantFaturadaAlterada

    iIndice = 0

    'Para cada item
    For Each objItemPV In objPedVenda.colItensPedido

        iIndice = iIndice + 1
        
        'Verifica se o controle é de reserva
        If objItemPV.iControleEstoque = PRODUTO_CONTROLE_RESERVA Then
            
            'Se for um item já existente
            If objItemPV.lNumIntDoc <> 0 Then
                'Pesquisa na colecao de itens do Pedido Venda lido no BD
                For Each objItemPV_BD In objPedidoVenda_BD.colItensPedido
                
                    If objItemPV_BD.lNumIntDoc = objItemPV.lNumIntDoc Then Exit For
                    
                Next

                'Se não achar erro
                If objItemPV_BD.lNumIntDoc <> objItemPV.lNumIntDoc Then Error 46001
                
                'Verifica se a quantidade faturada foi alterada
                If objItemPV.dQuantFaturada <> objItemPV_BD.dQuantFaturada Then Error 46002
            
            End If
        
        End If
    Next

    ItensPV_QuantFaturadaAlterada = SUCESSO

    Exit Function

Erro_ItensPV_QuantFaturadaAlterada:

    ItensPV_QuantFaturadaAlterada = Err

    Select Case Err

        Case 46001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_CADASTRADO", Err, objItemPV.lNumIntDoc)

        Case 46002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_ALTERADA", Err, iIndice, objPedVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152551)

    End Select

    Exit Function

End Function

Function PedidoVenda_Insere_BD(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere o novo Pedido de Venda no Banco de Dados

Dim lErro As Long
Dim dQuantNaoCancelada As Double
Dim lNumIntDoc As Long
Dim objMovEstoque As New ClassMovEstoque
Dim objCliente As New ClassCliente 'Inserido por Wagner
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemPV As ClassItemPedido
Dim objBloqueioPV As ClassBloqueioPV
Dim iIndice As Integer

On Error GoTo Erro_PedidoVenda_Insere_BD

    'zera os campos de quantfaturada, quantop, quantsc para garantir que quando
    'pedido e' novo, mesmo que ele seja oriundo de um pedido baixado, estas quantidades estejam zeradas
    For Each objItemPV In objPedidoVenda.colItensPedido

        objItemPV.dQuantFaturada = 0
        objItemPV.dQuantOP = 0
        objItemPV.dQuantSC = 0
        objItemPV.lNumIntDoc = 0
        
        If objItemPV.colItensRomaneioGrade.Count > 0 Then
        
            For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
                objItemRomaneio.dQuantFaturada = 0
                objItemRomaneio.dQuantOP = 0
                objItemRomaneio.dQuantSC = 0
            Next
            
        End If

    Next

     'Para cada Bloqueio do Pedido
    For iIndice = objPedidoVenda.colBloqueiosPV.Count To 1 Step -1
        Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Item(iIndice)
        If objBloqueioPV.dtDataLib <> DATA_NULA Then
            objPedidoVenda.colBloqueiosPV.Remove iIndice
        End If
    Next

    'Inclui no BD um bloqueio de pagamento antecipado, se necessario
    lErro = CF("BloqueioPagtoAnt_Adiciona", objPedidoVenda)
    If lErro <> SUCESSO Then gError 126281

    'Inclui no BD um bloqueio de dias de atraso, se necessario
    lErro = CF("BloqueioDiasAtraso_Adiciona", objPedidoVenda)
    If lErro <> SUCESSO Then gError 94368
    
    lErro = CF("BloqueiosPV_Cust_Adiciona", objPedidoVenda)
    If lErro <> SUCESSO Then gError 94368
    
    If gobjCRFAT.iBloquearPVPrecoBaixo <> 0 Then
        'Inclui no BD um bloqueio de preco "baixo"
        lErro = BloqueioPrecoBaixo_Adiciona(objPedidoVenda)
        If lErro <> SUCESSO Then gError 94368
    End If
    
    If gobjCRFAT.iBloquearPVMargemBaixa <> 0 Then
        'Inclui no BD um bloqueio de margem "baixa"
        lErro = BloqueioMargemBaixa_Adiciona(objPedidoVenda)
        If lErro <> SUCESSO Then gError 94368
    End If
    
    If gobjCRFAT.iBloquearPVPrecoDefasado <> 0 Then
        'Inclui no BD um bloqueio de preco "defasado"
        lErro = BloqueioPrecoDefasado_Adiciona(objPedidoVenda)
        If lErro <> SUCESSO Then gError 94368
    End If
    
    If gobjCRFAT.iCliAtrasoNaoGravaPV <> 0 Then
    
        lErro = CF("PV_Grava_AnalisaCliente", objPedidoVenda)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    'Inserido por Wagner
    '##################
    objCliente.lCodigo = objPedidoVenda.lCliente
    'le os dados do cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 129145
    
    If objCliente.iAtivo <> CLIENTE_ATIVO Then gError 129146
    '##################
    
    'tulio130103
    'verifica se eh para criar ou nao bloqueio de limite de credito
    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
        'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito
        lErro = CF("BloqueioCredito_Adiciona", objPedidoVenda)
        If lErro <> SUCESSO Then gError 42955
    End If

    'Adiciona um Bloqueio na coleção de Bloqueio de acordo com o a quantidade que falta reservar
    lErro = BloqueioReserva_Adiciona1(objPedidoVenda)
    If lErro <> SUCESSO Then gError 42956

    'Retorna a soma das quantidades não canceladas de todos os itens
    lErro = PedVenda_QuantNaoCancelada(objPedidoVenda, dQuantNaoCancelada)
    If lErro <> SUCESSO Then gError 42946
    
'''''    If lErro <> SUCESSO And lErro <> 42947 Then gError 42946
'''''    If lErro = 42947 Then gError 46000

    'Gera o número interno automático do Pedido de Venda
    lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_PEDIDO_VENDA", lNumIntDoc)
    If lErro <> SUCESSO Then gError 42948

    objPedidoVenda.lNumIntDoc = lNumIntDoc
    
    'Insere o Pedido na Tabela de Pedidos
    lErro = CF("PedidoVenda_Grava_BD", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42949

    'Insere as Parcelas do Pedido no BD
    lErro = CF("ParcelasPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42951

    lErro = ProdutoFilial_Adiciona_QuantPedida(objPedidoVenda)
    If lErro <> SUCESSO Then gError 51406

    'Insere os itens do Pedido no BD
    lErro = CF("ItensPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42952

    'Insere as comissões do Pedido no BD
    lErro = CF("ComissoesPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42950
    
    'Insere as reservas do Pedido no BD
    lErro = CF("ReservasPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 42953

    'Cria um Movimento de Estoque
    lErro = MovEstoque_Cria(objMovEstoque, objPedidoVenda)
    If lErro <> SUCESSO Then gError 42954

    'Grava o Movimento de Estoque
    lErro = CF("MovEstoque_Grava", objMovEstoque)
    If lErro <> SUCESSO Then gError 42971
    
'    If giTipoVersao = VERSAO_FULL Then
        
        'Insere no BD os Bloqueios do Pedido de Venda
        'Aqui o Pedido já está gravado e portanto acessivel no BD
        lErro = CF("BloqueiosPV_Grava", objPedidoVenda)
        If lErro <> SUCESSO Then gError 42957
    
'    End If
    
    lErro = PedidoVenda_EmbalagensGrava(objPedidoVenda)
    If lErro <> SUCESSO Then gError 42957
    
    PedidoVenda_Insere_BD = SUCESSO

    Exit Function

Erro_PedidoVenda_Insere_BD:

    PedidoVenda_Insere_BD = gErr

    Select Case gErr

        Case 42946 To 42957, 42971, 51406, 94368, 126281, 129145
        
        Case ERRO_SEM_MENSAGEM

        Case 46000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDVENDA_SEM_QUANTIDADE", gErr)

        Case 129146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objPedidoVenda.lCliente)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152552)

    End Select

    Exit Function

End Function

Function PedidoVenda_Altera_BD(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedVenda_BD As ClassPedidoDeVenda, alComando() As Long) As Long
'Faz as alterações no Pedido de Venda
'Passa os dados de ítens de PedidoVenda no BD para objPedVenda_BD

Dim lErro As Long
Dim objPedidoVenda_BD As New ClassPedidoDeVenda
Dim objItemPV As ClassItemPedido
Dim objBloqueioPV_BD As New ClassBloqueioPV
Dim objBloqueioPV1 As ClassBloqueioPV
Dim bAchouBloqueioBD As Boolean
Dim vbMsgRes As VbMsgBoxResult
Dim objBloqueioPV As New ClassBloqueioPV
Dim iIndice As Integer
Dim iBloqueioEstoqueAdicionado As Integer
Dim dPedidoFaltaReservar As Double
Dim bPossuiBloqueioTotal As Boolean
Dim objMovEstoque As New ClassMovEstoque
Dim dQuantNaoCancelada As Double
Dim iCreditoAprovado As Integer
Dim objItemPV_BD As ClassItemPedido
Dim bAdicionaBloqueio As Boolean
Dim iAtraso As Integer, colBloqBD As New Collection
Dim bPossuiBloqueioCredAtivo As Boolean
Dim objCliente As New ClassCliente 'Inserido Por Wagner

On Error GoTo Erro_PedidoVenda_Altera_BD

    'Inicialização de flags
    bPossuiBloqueioTotal = False

    objPedidoVenda_BD.lCodigo = objPedidoVenda.lCodigo
    objPedidoVenda_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa

    'Lê e faz o "lock" no Pedido
    lErro = CF("PedidoVenda_Lock", objPedidoVenda_BD, alComando(0))
    If lErro <> SUCESSO And lErro <> 42988 Then gError 42973

    If lErro = 42988 Then gError 46075 'Não Encontrou

    objPedidoVenda.lNumIntDoc = objPedidoVenda_BD.lNumIntDoc
    
    If objPedVenda_BD.lCliente <> objPedidoVenda.lCliente Or objPedVenda_BD.iFilial <> objPedidoVenda.iFilial Then gError 208098 'Não pode alterar o cliente e a filial de um pedido já salvo

    'Lê e locka os itens do Pedido
    lErro = CF("ItensPV_Le_Lock", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 42996
    
    'Exclui do BD as comissões do Pedido
    lErro = CF("ComissoesPV_Exclui", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objPedidoVenda_BD.iFilialEmpresaFaturamento = objPedVenda_BD.iFilialEmpresaFaturamento
    objPedidoVenda_BD.lCliente = objPedVenda_BD.lCliente
    objPedidoVenda_BD.iFilial = objPedVenda_BD.iFilial
    
    lErro = CF("PedidoVenda_Altera_Customizacoes", objPedidoVenda, objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 126505
    
'    If giTipoVersao = VERSAO_FULL Then
        '====BLOQUEIOS ==============================================
                        
        lErro = PedidoVenda_AtualizaObsBloq_Trans(objPedidoVenda)
        If lErro <> SUCESSO Then gError 124129
    
        'remove da colecao de bloqueios os que já estavam no BD
        'incluindo-os na colecao temporaria
        For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        
            If objBloqueioPV.iSequencial = 0 Then Exit For
        
            Call colBloqBD.Add(objBloqueioPV)
            Call objPedidoVenda.colBloqueiosPV.Remove(1)
            
        Next
    
        If gobjCRFAT.iBloquearPVMargemBaixa <> 0 Then
            
            lErro = CF("BloqueiosPV_Exclui_Tipo", objPedidoVenda, BLOQUEIO_MARGEM_BAIXA)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Inclui no BD um bloqueio de margem "baixa"
            lErro = BloqueioMargemBaixa_Adiciona(objPedidoVenda)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
    
    'Trecho inserido por Leo em 25/02/02
        
        '??? verificar se pedido tem bloqueio por atraso liberado
        
        'Verifica se o pedido tem credito aprovado
        lErro = CF("BloqueiosPV_Atraso_Testa", objPedidoVenda, iAtraso)
        If lErro <> SUCESSO Then gError 94382
        
        '??? se atraso ainda nao foi liberado, exclui bloqueios de atraso, se existirem
               
        'Se o bloqueio por atraso ainda não foi aprovado
        If iAtraso <> BLOQUEIO_POR_ATRASO_LIBERADO Then
        
            lErro = CF("BloqueiosPV_Exclui_Tipo", objPedidoVenda, BLOQUEIO_DIAS_ATRASO)
            If lErro <> SUCESSO Then gError 94383
            
            '??? chama BloqueioDiasAtraso_Adiciona
            
            lErro = CF("BloqueioDiasAtraso_Adiciona", objPedidoVenda)
            If lErro <> SUCESSO Then gError 94384
            
            'Inclui no BD um bloqueio de dias de atraso, se necessario
        
        End If
        
        '??? criar funcoes analogas a BloqueiosPV_Exclui_Credito,BloqueiosPV_Credito_Aprovado_Testa
        
        lErro = CF("BloqueiosPV_Le", objPedidoVenda_BD)
        If lErro <> SUCESSO Then gError 94382

        'Inserido por Wagner
        '##################
        objCliente.lCodigo = objPedidoVenda.lCliente
        'le os dados do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 129147
        
        If objCliente.iAtivo <> CLIENTE_ATIVO Then gError 129148
        '##################
         
        'tulio130103
        'se o cliente nao usar o bloqueio de credito, bypassar os testes
        If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
        
            bPossuiBloqueioCredAtivo = False
            
            For Each objBloqueioPV_BD In objPedidoVenda_BD.colBloqueiosPV
                If objBloqueioPV_BD.iTipoDeBloqueio = BLOQUEIO_CREDITO And objBloqueioPV_BD.dtDataLib = DATA_NULA Then
                    bPossuiBloqueioCredAtivo = True
                End If
            Next
            
            If bPossuiBloqueioCredAtivo Then
                'Verifica se o pedido tem credito aprovado
                lErro = CF("BloqueiosPV_Credito_Aprovado_Testa", objPedidoVenda, iCreditoAprovado)
                If lErro <> SUCESSO Then gError 44379
            
                bAdicionaBloqueio = False
            
                'Se o crédito ainda não foi aprovado
                If (iCreditoAprovado <> BLOQUEIO_CREDITO_LIBERADO) Then bAdicionaBloqueio = True
                    
            End If
            
            'Se Valor Total aumentou faz adição (condicional) de bloqueio de crédito
            If (Not bAdicionaBloqueio) And (objPedidoVenda.dValorTotal - objPedidoVenda_BD.dValorTotal) > DELTA_VALORMONETARIO Then bAdicionaBloqueio = True
            
            If bAdicionaBloqueio Then
                'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito
                lErro = CF("BloqueioCredito_Adiciona", objPedidoVenda)
                If lErro <> SUCESSO Then gError 46017
            
            End If
            If (iCreditoAprovado <> BLOQUEIO_CREDITO_LIBERADO) Then
                'Exclui bloqueios de Credito com data de liberacao = DATA_NULA
                lErro = CF("BloqueiosPV_Exclui_Credito", objPedidoVenda)
                If lErro <> SUCESSO Then gError 46016
                bAdicionaBloqueio = True
            End If

        End If
    
        objBloqueioPV_BD.lPedidoDeVendas = objPedidoVenda.lCodigo
        objBloqueioPV_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
        'Lê os bloqueio de estoque correspondente ao pedido que não esteja liberado
        lErro = CF("BloqueioEstoquePV_Le", objBloqueioPV_BD)
        If lErro <> SUCESSO And lErro <> 46032 Then gError 46033
    
        If lErro = SUCESSO Then bAchouBloqueioBD = True
        If lErro = 46032 Then bAchouBloqueioBD = False
    
        'Se encontrou e no BD existe bloqueio total
        If bAchouBloqueioBD And objBloqueioPV_BD.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
    
            'Avisa que as reservas serão canceladas
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BLOQUEIO_TOTAL2")
            If vbMsgRes = vbNo Then gError 46034
    
            'Limpa as reservas
            For Each objItemPV In objPedidoVenda.colItensPedido
                Set objItemPV.ColReserva = New ColReserva
            Next
    
            iIndice = 0
    
            'Exclui o Bloqueio total da coleção de bloqueios vinda da Tela
            For Each objBloqueioPV1 In objPedidoVenda.colBloqueiosPV
                iIndice = iIndice + 1
                If objBloqueioPV1.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
                    objPedidoVenda.colBloqueiosPV.Remove iIndice
                    Exit For
                End If
            Next
    
        Else 'nao encontrou bloqueio TOTAL de reserva
    
            'Adiciona um Bloqueio na coleção de Bloqueio de acordo com o a quantidade que falta reservar
            lErro = BloqueioReserva_Adiciona2(objPedidoVenda, objBloqueioPV_BD, objBloqueioPV, iBloqueioEstoqueAdicionado, dPedidoFaltaReservar)
            If lErro <> SUCESSO Then gError 46035
    
            'Verifica se na colecao de bloqueios existe algum bloqueio total
            For Each objBloqueioPV1 In objPedidoVenda.colBloqueiosPV
                If objBloqueioPV1.iTipoDeBloqueio = BLOQUEIO_TOTAL Then
                    bPossuiBloqueioTotal = True
                    Exit For
                End If
            Next
    
            If bAchouBloqueioBD And (dPedidoFaltaReservar = 0 Or objBloqueioPV.iTipoDeBloqueio <> objBloqueioPV_BD.iTipoDeBloqueio Or bPossuiBloqueioTotal) Then
                'Libera o Bloqueios de Estoque de Pedido
                lErro = CF("BloqueioEstoquePV_Libera", objBloqueioPV_BD)
                If lErro <> SUCESSO Then gError 46036
    
            End If
        
        End If
    
        lErro = CF("BloqueioPagtoAnt_Altera", objPedidoVenda, objPedidoVenda_BD)
        If lErro <> SUCESSO Then gError 126743
    
        'copiar observacoes de bloqueios automaticos que foram recriados
        Call PedidoVenda_BloqAutoAtualiza(objPedidoVenda, colBloqBD)
        
        'Grava os bloqueios do Pedido de Venda
        lErro = CF("BloqueiosPV_Grava", objPedidoVenda)
        If lErro <> SUCESSO Then gError 46042
        '=========FIM BLOQUEIOS=============================================
        
    'End If
    
    'Exclui as embalagens do Pedido no BD
    lErro = PedidoVenda_EmbalagensExclui(objPedidoVenda)
    If lErro <> SUCESSO Then gError 42957
    
    'Altera estatística de quantidade pedida do Produto
    lErro = ProdutoFilial_Altera_QuantPedida(objPedidoVenda, objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 51407

    'Passa as informações de objPedidoVenda_BD para objPedVenda_BD
    For Each objItemPV_BD In objPedidoVenda_BD.colItensPedido
        With objItemPV_BD
        objPedVenda_BD.colItensPedido.Add .iFilialEmpresa, .lCodPedido, .iItem, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .ColReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc
        End With
    Next
    
    'Verifica se a quantidade faturada do Item do BD é a mesma da tela
    lErro = ItensPV_QuantFaturadaAlterada(objPedidoVenda, objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 46003

    'Altera os dados do Pedido no BD
    lErro = CF("PedidoVenda_Grava_BD", objPedidoVenda, alComando(0))
    If lErro <> SUCESSO Then gError 46005

    'Atualiza as parcelas do Pedido de Venda
    lErro = CF("ParcelasPV_Altera", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46163
    
    'Insere ou Altera os itens do Pedido de Venda
    lErro = CF("ItensPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46043

    'Atualiza as comissões do Pedido passado
    lErro = CF("ComissoesPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46162

    'Lê as reservas do Pedido existentes no BD
    lErro = CF("ReservasPedido_Le", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 46050

    'Cria um Movimento de Estoque de Cancelamento de Reserva
    lErro = MovEstoqueExclusao_Cria(objMovEstoque, objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 46051

    'Grava o Movimento de Estoque de exclusao de reservas
    lErro = CF("MovEstoque_Grava", objMovEstoque)
    If lErro <> SUCESSO Then gError 46052

    'Exclui reservas da tabela de reservas e inclui as vindas da tela
    lErro = CF("ReservasPV_Altera", objPedidoVenda, alComando(1), alComando(2))
    If lErro <> SUCESSO Then gError 46163

    'Cria o Movimento de Estoque para inclusão das reservas
    lErro = MovEstoque_Cria(objMovEstoque, objPedidoVenda)
    If lErro <> SUCESSO Then gError 46068

    'Grava o Movimento de Estoque
    lErro = CF("MovEstoque_Grava", objMovEstoque)
    If lErro <> SUCESSO Then gError 46069

    'Remove de colItens do BD os itens presente na colItens da tela
    For Each objItemPV In objPedidoVenda.colItensPedido
        For iIndice = objPedidoVenda_BD.colItensPedido.Count To 1 Step -1
            If objItemPV.lNumIntDoc = objPedidoVenda_BD.colItensPedido(iIndice).lNumIntDoc Then
                objPedidoVenda_BD.colItensPedido.Remove iIndice
                Exit For
            End If
        Next
    Next

    'Verifica para cada item se existe vínculo com um item de Ordem de Produção
    lErro = CF("ItensPV_ItensOP_Vinculados", objPedidoVenda_BD)
    If lErro <> SUCESSO Then gError 46070

    'Exclui os ítens de Pedido de venda que foram excluídos na tela
    lErro = CF("ItensPV_Exclui", objPedidoVenda_BD, alComando(3), alComando(4), alComando(23), alComando(24))
    If lErro <> SUCESSO Then gError 46078

    lErro = PedidoVenda_EmbalagensGrava(objPedidoVenda)
    If lErro <> SUCESSO Then gError 42957
    
    PedidoVenda_Altera_BD = SUCESSO

    Exit Function

Erro_PedidoVenda_Altera_BD:

    PedidoVenda_Altera_BD = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 42973, 42996, 44379, 46003, 46005, 46162, 46163, 46016, 46017, 46033, 46034, 46035, 46036, 46042, 46043, 46050, 46051, 46052, 46163, 46068, 46069, 46070, 46078, 51407, 96128, 96129, 94369, 94382 To 94384, 124129, 126505, 126743, 129147

        Case 46075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case 129148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objPedidoVenda.lCliente)

        Case 208098
            Call Rotina_Erro(vbOKOnly, "ERRO_PV_TROCA_CLIENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152553)

    End Select

    Exit Function

End Function

Function MovEstoque_Cria(objMovEstoque As ClassMovEstoque, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Cria para o Pedido de Venda passado um Movimento de Estoque do tipo MOV_EST_INCLUI_RESERVA

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objReserva As ClassReserva
Dim lCodigo As Long, objItemMovEstoque As ClassItemMovEstoque
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_MovEstoque_Cria

    Set objMovEstoque = New ClassMovEstoque

    'Gera o número interno do movimento
    lErro = CF("MovEstoque_Automatico_EmTransacao", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then Error 46059

    'Preenche o objMovEstoque
    objMovEstoque.dtData = gdtDataHoje
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    objMovEstoque.iTipoMov = MOV_EST_INCLUI_RESERVA
    objMovEstoque.lCodigo = lCodigo

    'Para cada item do Pedido
    For Each objItemPV In objPedidoVenda.colItensPedido

        If objItemPV.colItensRomaneioGrade.Count = 0 Then

            'Para cada reserva do item
            For Each objReserva In objItemPV.ColReserva
    
                'Adiciona ítem de MovEstoque
                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, 0, 0, objItemPV.sProduto, "", objItemPV.sUMEstoque, objReserva.dQuantidade, objReserva.iAlmoxarifado, objReserva.sAlmoxarifado, objReserva.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
    
                objItemMovEstoque.iClasseUM = objItemPV.iClasseUM
                objItemMovEstoque.sSiglaUMEst = objItemPV.sUMEstoque
                objItemMovEstoque.iControleEstoque = objItemPV.iControleEstoque
                objItemMovEstoque.iApropriacaoProd = objItemPV.iApropriacaoProd
                
            Next
        
        End If
        
        For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
            For Each objReservaItem In objItemRomaneio.colLocalizacao
            
            'Adiciona item de MovEstoque
            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, 0, 0, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objReservaItem.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)

            objItemMovEstoque.iClasseUM = objItemPV.iClasseUM
            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
            
            Next
        Next
        
    Next

    MovEstoque_Cria = SUCESSO

    Exit Function

Erro_MovEstoque_Cria:

    MovEstoque_Cria = Err

    Select Case Err

        Case 46059

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152554)

    End Select

    Exit Function

End Function

Function MovEstoqueExclusao_Cria(objMovEstoque As ClassMovEstoque, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Cria um Movimento de Estoque de Cancelamento de Reserva para o pedido passado

Dim lErro As Long
Dim lCodigo As Long
Dim objReserva As ClassReserva
Dim objItemPV As ClassItemPedido
Dim objProduto As New ClassProduto
Dim bProdutoLeu As Boolean
Dim sUMEstoque As String
Dim objReservaItem  As ClassReservaItem
Dim objItemRomaneio  As ClassItemRomaneioGrade

On Error GoTo Erro_MovEstoqueExclusao_Cria

    Set objMovEstoque = New ClassMovEstoque

    'Gera o código do movimento
    lErro = CF("MovEstoque_Automatico_EmTransacao", giFilialEmpresa, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46060

    'Carrega o objMovEstoque
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    objMovEstoque.dtData = gdtDataHoje
    objMovEstoque.iTipoMov = MOV_EST_CANCELA_RESERVA
    objMovEstoque.lCodigo = lCodigo

    'Para casa item do Pedido de Venda
    For Each objItemPV In objPedidoVenda.colItensPedido
        'Para cada reserva do item do Pedido de Venda
        For Each objReserva In objItemPV.ColReserva

            objProduto.sCodigo = objItemPV.sProduto
            'Lê o produto do item
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 46061
            If lErro <> SUCESSO Then Error 46062
            
            'Adiciona ítem de MovEstoque
            objMovEstoque.colItens.Add 0, 0, 0, 0, objReserva.sProduto, "", objProduto.sSiglaUMEstoque, objReserva.dQuantidade, objReserva.iAlmoxarifado, "", objReserva.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA

        Next

        For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
            
            objProduto.sCodigo = objItemRomaneio.sProduto
            'Lê o produto do item
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 46061
            If lErro <> SUCESSO Then Error 46062
            
            For Each objReservaItem In objItemRomaneio.colLocalizacao
            'Adiciona item de MovEstoque
            objMovEstoque.colItens.Add 0, 0, 0, 0, objItemRomaneio.sProduto, "", objProduto.sSiglaUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, "", objReservaItem.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
            
            Next
        Next

    Next

    MovEstoqueExclusao_Cria = SUCESSO

    Exit Function

Erro_MovEstoqueExclusao_Cria:

    MovEstoqueExclusao_Cria = Err

    Select Case Err

        Case 46060, 46061

        Case 46062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152555)

    End Select

    Exit Function

End Function

Function PedidoDeVenda_Automatico(lNumPedido As Long) As Long
'Gera o próximo Número de um Pedido De Venda

Dim lErro As Long

On Error GoTo Erro_PedidoDeVenda_Automatico

    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_CODIGO_PEDVENDA", "PedVenTodos", "Codigo", lNumPedido)
    If lErro <> SUCESSO Then Error 26797

    PedidoDeVenda_Automatico = SUCESSO

    Exit Function

Erro_PedidoDeVenda_Automatico:

    PedidoDeVenda_Automatico = Err

    Select Case Err

        Case 26797

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152556)

    End Select

    Exit Function

End Function

Function BloqueiosPV_Exclui_Credito(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui os Bloqueios de Creditos não liberados ligados ao Pedido de Venda passado

'Função alterada por leo em 26/02/02. Agora só chama uma função genérica que exclui os bloqueios do tipo passado como parâmetro
Dim lErro As Long

On Error GoTo Erro_BloqueiosPV_Exclui_Credito
    
    'Função que exclui os Bloqueios do Tipo passada como parâmetro.
    lErro = CF("BloqueiosPV_Exclui_Tipo", objPedidoVenda, BLOQUEIO_CREDITO)
    If lErro <> SUCESSO Then gError 94388
    
    BloqueiosPV_Exclui_Credito = SUCESSO

    Exit Function

Erro_BloqueiosPV_Exclui_Credito:

    BloqueiosPV_Exclui_Credito = gErr

    Select Case gErr

        Case 94388
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152557)

    End Select
    
    Exit Function

End Function

Function PedVenda_Grava_Estatisticas(ByVal objPedVenda As ClassPedidoDeVenda, Optional ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística dos produtos Pedidos

Dim lErro As Long

On Error GoTo Erro_PedVenda_Grava_Estatisticas
           
    If objPedVenda_BD Is Nothing Then
        
        'Grava a estatistica do Dia do produto
        lErro = SldDiaFat_Grava_EstPV(objPedVenda)
        If lErro <> SUCESSO Then Error 52749
    
        'Grava a estatística do Mês do produto
        lErro = SldMesFat_Grava_EstPV(objPedVenda)
        If lErro <> SUCESSO Then Error 52750

    Else
            
        'Grava a estatística do Dia do produto
        lErro = SldDiaFat_Grava_EstPV(objPedVenda, objPedVenda_BD)
        If lErro <> SUCESSO Then Error 52751
    
        'Grava a estatística do Mês do produto
        lErro = SldMesFat_Grava_EstPV(objPedVenda, objPedVenda_BD)
        If lErro <> SUCESSO Then Error 52752
    
    End If
            
    PedVenda_Grava_Estatisticas = SUCESSO
    
    Exit Function
    
Erro_PedVenda_Grava_Estatisticas:
    
    PedVenda_Grava_Estatisticas = Err

    Select Case Err
        
        Case 52749, 52750, 52751, 52752 'Tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152558)

    End Select
    
    Exit Function

End Function

Function PedVenda_Exclui_Estatisticas(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Exclui a estatística dos Produtos Pedidos

Dim lErro As Long

On Error GoTo Erro_PedVenda_Exclui_Estatisticas
           
    'Exclui a estatistica do Dia do produto
    lErro = SldDiaFat_Exclui_EstPV(objPedVenda)
    If lErro <> SUCESSO Then gError 64423
    
    'Exclui a estatística do Mês do produto
    lErro = SldMesFat_Exclui_EstPV(objPedVenda)
    If lErro <> SUCESSO Then gError 64424

    PedVenda_Exclui_Estatisticas = SUCESSO
    
    Exit Function
    
Erro_PedVenda_Exclui_Estatisticas:
    
    PedVenda_Exclui_Estatisticas = gErr

    Select Case gErr
        
        Case 64423, 64424 'Tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152559)

    End Select
    
    Exit Function

End Function

Private Function SldDiaFat_Grava_EstPV(ByVal objPedVenda As ClassPedidoDeVenda, Optional ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objItemPV As New ClassItemPedido
Dim objItemPV_BD As New ClassItemPedido
Dim objSldDiaFat As New ClassSldDiaFat
Dim dQuantUMVenda As Double
Dim dFator As Double
Dim alComando(1 To 5) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldDiaFat_Grava_EstPV
        
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52753
    Next
          
    'Para cada item do PedidoVenda
    For Each objItemPV In objPedVenda.colItensPedido
        
        objProduto.sCodigo = objItemPV.sProduto
        
        'Lê o produto para pegar a UM_Venda
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 52802
    
        If lErro = 28030 Then Error 52803
            
        'Lê o fator de conversão da UM_Venda
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then Error 52804

        'Converte a UM da quantidade
        dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
        
        objSldDiaFat.iFilialEmpresa = giFilialEmpresa
        objSldDiaFat.dtData = objPedVenda.dtDataEmissao
        objSldDiaFat.sProduto = objItemPV.sProduto
        
        'Lê todos os atributos na tabela SldDiaFat e faz lock exclusive
        lErro = CF("SldDiaFat_Le_Lock", alComando(1), objSldDiaFat)
        If lErro <> SUCESSO And lErro <> 52729 Then Error 52754
                    
        'Se não encontrou
        If lErro = 52729 Then
            
            'Insere na tabela de SldDiaFat um novo registro
            lErro = Comando_Executar(alComando(5), "INSERT INTO SldDiaFat (FilialEmpresa, Data, Produto, QuantPedida) VALUES (?,?,?,?)", giFilialEmpresa, objPedVenda.dtDataEmissao, objItemPV.sProduto, dQuantUMVenda)
            If lErro <> AD_SQL_SUCESSO Then Error 52805
        
        Else
            'Se não atualiza
            
            objSldDiaFat.dQuantPedida = objSldDiaFat.dQuantPedida + dQuantUMVenda
            
            'Atualiza a tabela de SldDiaFat
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldDiaFat SET QuantPedida =?", alComando(1), objSldDiaFat.dQuantPedida)
            If lErro <> AD_SQL_SUCESSO Then Error 52758
        
        End If
    
    Next
    
    'Se o parametro foi preenchido é um pedido alteração
    If Not (objPedVenda_BD Is Nothing) Then
            
        For Each objItemPV_BD In objPedVenda_BD.colItensPedido
        
            objSldDiaFat.iFilialEmpresa = giFilialEmpresa
            objSldDiaFat.dtData = objPedVenda_BD.dtDataEmissao
            objSldDiaFat.sProduto = objItemPV_BD.sProduto
            
            'Lê todos os atributos na tabela SldDiaFat e faz lock exclusive
            lErro = CF("SldDiaFat_Le_Lock", alComando(3), objSldDiaFat)
            If lErro <> SUCESSO Then Error 52759
                        
            objProduto.sCodigo = objItemPV_BD.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 52760
            
            If lErro = 28030 Then Error 52761
                    
            'Calcula o fator de conversão da UM_Venda
            lErro = CF("UM_Conversao_Trans", objItemPV_BD.iClasseUM, objItemPV_BD.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then Error 52762
        
            'Converte a quantidade
            dQuantUMVenda = ((objItemPV_BD.dQuantidade - objItemPV_BD.dQuantCancelada) * dFator)
                
            objSldDiaFat.dQuantPedida = objSldDiaFat.dQuantPedida - dQuantUMVenda
            
            'Atualiza a tabela de SldDiaFat para alteração
            lErro = Comando_ExecutarPos(alComando(4), "UPDATE SldDiaFat SET QuantPedida =?", alComando(3), objSldDiaFat.dQuantPedida)
            If lErro <> AD_SQL_SUCESSO Then Error 52763
            
        Next
        
    End If
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    SldDiaFat_Grava_EstPV = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Grava_EstPV:
    
    SldDiaFat_Grava_EstPV = Err

    Select Case Err
    
        Case 52753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52754, 52755, 52757, 52759, 52762, 52760, 52802, 52804 'Tratados nas rotinas chamadas
        
        Case 52756, 52761, 52803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 52758, 52763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFAT", Err, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case 52805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDDIAFAT", Err, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152560)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
        
Private Function SldDiaFat_Exclui_EstPV(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Grava a estatística do dia do Produto Pedido

Dim lErro As Long
Dim objItemPV As New ClassItemPedido
Dim objSldDiaFat As New ClassSldDiaFat
Dim dQuantUMVenda As Double
Dim dFator As Double
Dim alComando(1 To 2) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldDiaFat_Exclui_EstPV
        
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 64411
    Next
          
    'Para cada item do PedidoVenda
    For Each objItemPV In objPedVenda.colItensPedido
        
        objProduto.sCodigo = objItemPV.sProduto
        
        'Lê o produto para pegar a UM_Venda
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 64412
    
        If lErro = 28030 Then gError 64413
            
        'Lê o fator de conversão da UM_Venda
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 64414

        'Converte a UM da quantidade
        dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
        
        objSldDiaFat.iFilialEmpresa = giFilialEmpresa
        objSldDiaFat.dtData = objPedVenda.dtDataEmissao
        objSldDiaFat.sProduto = objItemPV.sProduto
        
        'Lê todos os atributos na tabela SldDiaFat e faz lock exclusive
        lErro = CF("SldDiaFat_Le_Lock", alComando(1), objSldDiaFat)
        If lErro <> SUCESSO And lErro <> 52729 Then gError 64415
                    
        'Se encontrou
        If lErro = SUCESSO Then
            
            objSldDiaFat.dQuantPedida = objSldDiaFat.dQuantPedida - dQuantUMVenda
            
            'Atualiza a tabela de SldDiaFat
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldDiaFat SET QuantPedida =?", alComando(1), objSldDiaFat.dQuantPedida)
            If lErro <> AD_SQL_SUCESSO Then gError 64416
        
        End If
    
    Next
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    SldDiaFat_Exclui_EstPV = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Exclui_EstPV:
    
    SldDiaFat_Exclui_EstPV = gErr

    Select Case gErr
    
        Case 64411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 64412, 64414, 64415 'Tratados nas rotinas chamadas
        
        Case 64413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 64416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFAT", gErr, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152561)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
        
Private Function SldMesFat_Grava_EstPV(ByVal objPedVenda As ClassPedidoDeVenda, Optional ByVal objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim objItemPV_BD As ClassItemPedido
Dim alComando(1 To 5) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldMesFat_Grava_EstPV

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52764
    Next

    'Para cada item do Pedido de Venda
    For Each objItemPV In objPedVenda.colItensPedido
        
        objSldMesFat.iFilialEmpresa = giFilialEmpresa
        objSldMesFat.iAno = Year(objPedVenda.dtDataEmissao)
        objSldMesFat.sProduto = objItemPV.sProduto
        
        'Lê todos os atributos na tabela SldMesFat e faz lock exclusive
        lErro = CF("SldMesFat_Le_Lock", alComando(1), objSldMesFat)
        If lErro <> SUCESSO And lErro <> 52733 Then Error 52765
                
        If lErro = 52733 Then
            
            iMes = Month(objPedVenda.dtDataEmissao)
        
            objProduto.sCodigo = objItemPV.sProduto
            
            'Lê o produto para pegar a UMVenda
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 52806
            
            If lErro = 28030 Then Error 52807
                    
            'Calcula o fator da UM_Venda
            lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then Error 52808
        
            'Converte a quantidade
            dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
                        
            sMes = CStr(iMes)
    
            'Atualiza a tabela de SldMesFat para inserção
            lErro = Comando_Executar(alComando(2), "INSERT INTO SldMesFat (FilialEmpresa, Ano, Produto, QuantPedida" + sMes + ") VALUES (?,?,?,?)", giFilialEmpresa, objSldMesFat.iAno, objSldMesFat.sProduto, dQuantUMVenda)
            If lErro <> AD_SQL_SUCESSO Then Error 52809

        Else
                    
            iMes = Month(objPedVenda.dtDataEmissao)
            
            objProduto.sCodigo = objItemPV.sProduto
            
            'Lê o produto para pegar a UMVenda
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 52766
            
            If lErro = 28030 Then Error 52767
                    
            'Calcula o fator para a UM_Venda
            lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then Error 52768
        
            'converte a quantidade
            dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
                
            objSldMesFat.dQuantPedida(iMes) = objSldMesFat.dQuantPedida(iMes) + dQuantUMVenda
            
            sMes = CStr(iMes)
    
            'Atualiza a tabela de SldMesFat para inserção
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldMesFat SET QuantPedida" + sMes + "=?", alComando(1), objSldMesFat.dQuantPedida(iMes))
            If lErro <> AD_SQL_SUCESSO Then Error 52769
        
        End If
        
    Next
    
    'Se o parametro foi preenchido é um pedido alteração
    If Not (objPedVenda_BD Is Nothing) Then
    
        'Para cada objItemPV
        For Each objItemPV_BD In objPedVenda_BD.colItensPedido
            
            objSldMesFat.iFilialEmpresa = giFilialEmpresa
            objSldMesFat.iAno = Year(objPedVenda_BD.dtDataEmissao)
            objSldMesFat.sProduto = objItemPV_BD.sProduto
            
            'Lê todos os atributos na tabela SldMesFat e faza lock exclusive
            lErro = CF("SldMesFat_Le_Lock", alComando(3), objSldMesFat)
            If lErro <> SUCESSO Then Error 52770
            
            iMes = Month(objPedVenda_BD.dtDataEmissao)
            
            objProduto.sCodigo = objItemPV_BD.sProduto
            
            'Lê o produto para pegar a UMVenda
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 52771
            
            If lErro = 28030 Then Error 52772
                    
            'Calcula o fator da UM_Venda
            lErro = CF("UM_Conversao_Trans", objItemPV_BD.iClasseUM, objItemPV_BD.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then Error 52773
        
            'converte a quantidade
            dQuantUMVenda = ((objItemPV_BD.dQuantidade - objItemPV_BD.dQuantCancelada) * dFator)
                
            objSldMesFat.dQuantPedida(iMes) = objSldMesFat.dQuantPedida(iMes) - dQuantUMVenda
            
            sMes = CStr(iMes)
    
            'Atualiza a tabela de SldMesFat para alteração
            lErro = Comando_ExecutarPos(alComando(4), "UPDATE SldMesFat SET QuantPedida" + sMes + "=?", alComando(3), objSldMesFat.dQuantPedida(iMes))
            If lErro <> AD_SQL_SUCESSO Then Error 52774
    
        Next
        
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    SldMesFat_Grava_EstPV = SUCESSO

    Exit Function

Erro_SldMesFat_Grava_EstPV:

    SldMesFat_Grava_EstPV = Err

    Select Case Err
        
        Case 52764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52765, 52766, 52768, 52770, 52771, 52773, 52806, 52808 'Tratados nas rotinas chamadoras
        
        Case 52767, 52772, 52807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 52769, 52774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFAT", Err, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case 52809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFAT", Err, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152562)

    End Select

    Exit Function
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

End Function

Private Function SldMesFat_Exclui_EstPV(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Grava a estatística do Mês do produto

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objSldMesFat As New ClassSldMesFat
Dim iMes As Integer
Dim dFator As Double
Dim dQuantUMVenda As Double
Dim sMes As String
Dim alComando(1 To 2) As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_SldMesFat_Exclui_EstPV

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 64417
    Next

    'Para cada item do Pedido de Venda
    For Each objItemPV In objPedVenda.colItensPedido
        
        objSldMesFat.iFilialEmpresa = giFilialEmpresa
        objSldMesFat.iAno = Year(objPedVenda.dtDataEmissao)
        objSldMesFat.sProduto = objItemPV.sProduto
        
        'Lê todos os atributos na tabela SldMesFat e faz lock exclusive
        lErro = CF("SldMesFat_Le_Lock", alComando(1), objSldMesFat)
        If lErro <> SUCESSO And lErro <> 52733 Then gError 64418
                
        If lErro = SUCESSO Then
            
            iMes = Month(objPedVenda.dtDataEmissao)
            
            objProduto.sCodigo = objItemPV.sProduto
            
            'Lê o produto para pegar a UMVenda
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 64419
            
            If lErro = 28030 Then gError 64420
                    
            'Calcula o fator para a UM_Venda
            lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 64421
        
            'converte a quantidade
            dQuantUMVenda = ((objItemPV.dQuantidade - objItemPV.dQuantCancelada) * dFator)
                
            objSldMesFat.dQuantPedida(iMes) = objSldMesFat.dQuantPedida(iMes) - dQuantUMVenda
            
            sMes = CStr(iMes)
    
            'Atualiza a tabela de SldMesFat para inserção
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE SldMesFat SET QuantPedida" + sMes + "=?", alComando(1), objSldMesFat.dQuantPedida(iMes))
            If lErro <> AD_SQL_SUCESSO Then gError 64422
        
        End If
        
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    SldMesFat_Exclui_EstPV = SUCESSO

    Exit Function

Erro_SldMesFat_Exclui_EstPV:

    SldMesFat_Exclui_EstPV = gErr

    Select Case gErr
        
        Case 64417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 64418, 64419, 64421 'Tratados nas rotinas chamadoras
        
        Case 64420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 64422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFAT", gErr, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152563)

    End Select

    Exit Function
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

End Function

'Function FilialFornecedor_Grava_EstatisticasDev(objNFiscal As ClassNFiscal) As Long
''Tem que ser chamado dentro de Transação
''Atualiza a gravação da parte de estatística de devolução de Fornecedor
'
'Dim lErro As Long
'Dim objFilFornFilEmp As New ClassFilFornFilEmp
'Dim lComando As Long
'Dim lComando2 As Long
'Dim lComando3 As Long
'
'On Error GoTo Erro_FilialFornecedor_Grava_EstatisticasDev
'
'    'Inicializa comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 52644
'
'    lComando2 = Comando_Abrir()
'    If lComando2 = 0 Then Error 52645
'
'    lComando3 = Comando_Abrir()
'    If lComando3 = 0 Then Error 52646
'
'    If objNFiscal.iTipoNFiscal = DOCINFO_NFISDCM Then
'
'        objFilFornFilEmp.lCodFornecedor = objNFiscal.lFornecedor
'        objFilFornFilEmp.iFilialEmpresa = objNFiscal.iFilialEmpresa
'        objFilFornFilEmp.iCodFilial = objNFiscal.iFilialForn
'
'        'Lê e locka a Tabela FilialFornFilEmp
'        lErro = CF("FilFornFilEmp_Le_Lock",lComando, objFilFornFilEmp)
'        If lErro <> SUCESSO And lErro <> 52642 Then Error 52647
'
'        'se não encontrou
'        If lErro = 52642 Then
'
'            'Insere na tabela de FilialFornFilEmp
'            lErro = Comando_Executar(lComando3, "INSERT INTO FilialFornFilEmp (FilialEmpresa,CodFornecedor, CodFilial, ValorAcumuladoDevolucoes, NumTotalDevolucoes) VALUES (?,?,?,?,?)", giFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dValorTotal, 1)
'            If lErro <> AD_SQL_SUCESSO Then Error 52648
'
'        'se encontrou
'        Else
'            'recalcula os valores
'            objFilFornFilEmp.dValorAcumuladoDevolucoes = objFilFornFilEmp.dValorAcumuladoDevolucoes + objNFiscal.dValorTotal
'            objFilFornFilEmp.lNumTotalDevolucoes = objFilFornFilEmp.lNumTotalDevolucoes + 1
'
'            'Atualiza a tabela de FilialFornFilEmp
'            lErro = Comando_ExecutarPos(lComando2, "UPDATE FilialFornFilEmp SET ValorAcumuladoDevolucoes = ?, NumTotalDevolucoes = ?", lComando, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lNumTotalDevolucoes)
'            If lErro <> AD_SQL_SUCESSO Then Error 52649
'
'        End If
'
'    End If
'
'    'Libera comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando2)
'    Call Comando_Fechar(lComando3)
'
'    FilialFornecedor_Grava_EstatisticasDev = SUCESSO
'
'    Exit Function
'
'Erro_FilialFornecedor_Grava_EstatisticasDev:
'
'    FilialFornecedor_Grava_EstatisticasDev = Err
'
'    Select Case Err
'
'        Case 52644, 52645, 52646
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 52647
'
'        Case 52648
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERIR_FILIALFORNFILEMP", Err)
'
'        Case 52649
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALFORNFILEMP", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152564)
'
'    End Select
'
'    'Libera comandos
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando2)
'    Call Comando_Fechar(lComando3)
'
'    Exit Function
'
'End Function

Function PedidoDeVenda_Baixar(alComando() As Long, ByVal objPedidoDeVenda As ClassPedidoDeVenda, Optional ByVal bLerReservasDoBD As Boolean = True) As Long
'transfere os dados do pedido para as tabelas de pedidos baixados
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim lErro As Long
Dim lNumIntDoc As Long
Dim tPedidoVenda As typePedidoVenda
Dim tItemPedido As typeItemPedido
Dim objBloqueioPV As New ClassBloqueioPV
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_PedidoDeVenda_Baixar
    
    If bLerReservasDoBD Then
    
        'Lê as reservas dos ítens do Pedido existentes no BD
        lErro = CF("ReservasPedido_Le", objPedidoDeVenda)
        If lErro <> SUCESSO Then Error 25643

    End If

    'Cria um Movimento de Estoque de Cancelamento de Reserva
    lErro = MovEstoqueExclusao_Cria(objMovEstoque, objPedidoDeVenda)
    If lErro <> SUCESSO Then Error 25644

    If objMovEstoque.colItens.Count > 0 Then
        'Grava o Movimento de Estoque de exclusao de reservas
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then Error 25645
    End If
    
    lErro = CF("PedidoDeVenda_Baixar1", alComando(), objPedidoDeVenda, NAO_FATURA_TUDO)
    If lErro <> SUCESSO Then Error 44296
           
    PedidoDeVenda_Baixar = SUCESSO

    Exit Function

Erro_PedidoDeVenda_Baixar:

    PedidoDeVenda_Baixar = Err

    Select Case Err
            
        Case 25643, 25644, 25645, 44296
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152565)
        
    End Select

    Exit Function

End Function

Function PedidoDeVenda_Baixar1(alComando() As Long, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal iFaturaTudo As Integer) As Long
'transfere os dados do pedido para as tabelas de pedidos baixados
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lPedidoVenda As Long
Dim objBloqueioPV As New ClassBloqueioPV

On Error GoTo Erro_PedidoDeVenda_Baixar1
        
    'Lê o Pedido de Venda, faz o lock e o insere na tabela de Pedidos baixados
    lErro = CF("PedidoVendaBaixado_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then Error 46088

    'Lê e faz o lock nos itens de PV e os insere na tabela de Itens PV baixados
    lErro = CF("ItensPVBaixados_Grava", objPedidoVenda, iFaturaTudo)
    If lErro <> SUCESSO Then Error 46089

    'Faz a baixa das comissões do Pedido
    lErro = CF("ComissoesPV_Baixa", objPedidoVenda, alComando(5), alComando(6), alComando(7))
    If lErro <> SUCESSO Then Error 46090

    'Faz a Baixa das Parcelas do Pedido de Venda
    lErro = CF("ParcelasPV_Baixa", objPedidoVenda, alComando(8), alComando(9), alComando(10))
    If lErro <> SUCESSO Then Error 46091

    'Faz a baixa dos Bloqueios do Pedido de Venda
    lErro = CF("BloqueiosPV_Baixa", objPedidoVenda, alComando(11), alComando(12), alComando(13))
    If lErro <> SUCESSO Then Error 46092
    
    'Exclui as reservas vinculadas ao Pedido de Venda
    lErro = CF("ReservasPV_Exclui", objPedidoVenda, alComando(14), alComando(15))
    If lErro <> SUCESSO Then Error 46095

    'Exclui os Itens do Pedido de Venda
    lErro = ItensPV_Exclui1(objPedidoVenda, alComando(16), alComando(17), alComando(20), alComando(21))
    If lErro <> SUCESSO Then Error 46096
    
    'Exclui o Pedido de Venda do BD
    lErro = PedidoVenda_Exclui_BD1(objPedidoVenda, alComando(18), alComando(19), alComando(24), alComando(25), alComando(26), alComando(27))
    If lErro <> SUCESSO Then Error 46097
    
'    'Baixa Tributação do Pedido de Vendas
'    lErro = TributacaoPV_Baixa(objPedidoVenda)
'    If lErro <> SUCESSO Then Error 57778
    
    PedidoDeVenda_Baixar1 = SUCESSO

    Exit Function

Erro_PedidoDeVenda_Baixar1:

    PedidoDeVenda_Baixar1 = Err

    Select Case Err

        Case 46088 To 46097, 57778

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152566)
        
    End Select

    Exit Function

End Function

'Function TributacaoPV_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Baixa Tributação relacionada ao pedido de vendas passado como parâmetro
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim lErro As Long
'
'On Error GoTo Erro_TributacaoPV_Baixa
'
'    lErro = TributacaoPV_Baixa_BD(objPedidoVenda)
'    If lErro <> SUCESSO Then Error 57786
'
'    lErro = TributacaoItemPV_Baixa(objPedidoVenda)
'    If lErro <> SUCESSO Then Error 57787
'
'    lErro = TributacaoComplPV_Baixa(objPedidoVenda)
'    If lErro <> SUCESSO Then Error 57788
'
'    TributacaoPV_Baixa = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Baixa:
'
'    TributacaoPV_Baixa = Err
'
'    Select Case Err
'
'        Case 57786, 57787, 57788
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152567)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function TributacaoPV_Baixa_BD(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Baixa Tributação do pedido de vendas passado
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim alComando(0 To 2) As Long
'Dim lErro As Long
'Dim iIndice As Integer
'Dim typeTributacaoPV As typeTributacaoPV
'
'On Error GoTo Erro_TributacaoPV_Baixa_BD
'
'    'Abertura dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 57789
'    Next
'
'    'Pesquisa no BD a Tributação Relacionada ao Pedido de Vendas
'    With typeTributacaoPV
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoPV WHERE FilialEmpresa = ? AND CodPedido = ?", 0, _
'            .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual, _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then Error 57790
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57791
'
'    'Se não encontrou --> Erro
'    If lErro <> AD_SQL_SUCESSO Then Error 57792
'
'    'Exclui do BD a Tributação Relacionado ao Pedido de Vendas passado
'    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoPV", alComando(0))
'    If lErro <> AD_SQL_SUCESSO Then Error 57793
'
'    'Insere em TributacaoPVBaixado a Tributação do Pedido de Vendas passado
'    With typeTributacaoPV
'        lErro = Comando_Executar(alComando(2), "INSERT INTO TributacaoPVBaixado (FilialEmpresa, CodPedido, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'    If lErro <> SUCESSO Then Error 57794
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    TributacaoPV_Baixa_BD = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Baixa_BD:
'
'    TributacaoPV_Baixa_BD = Err
'
'    Select Case Err
'
'        Case 57789
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 57790, 57791
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBPEDIDO", Err)
'
'        Case 57792
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBPEDIDO_NAO_ENCONTRADA", Err, objPedidoVenda.lCodigo)
'
'        Case 57793
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBPEDIDO", Err)
'
'        Case 57794
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152568)
'
'    End Select
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Function TributacaoItemPV_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Baixa os Itens da Tributação do pedido de vendas passado
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim alComando(0 To 2) As Long
'Dim lErro As Long
'Dim iNaturezaOpManual As Integer
'Dim iIndice As Integer
'Dim objItemPV As ClassItemPedido
'Dim tTributacaoItemPV As typeTributacaoItemPV
'
'On Error GoTo Erro_TributacaoItemPV_Baixa
'
'    'Abertura dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 57795
'    Next
'
'    'Para cada item da coleção de pedido de vendas
'    For Each objItemPV In objPedidoVenda.colItensPedido
'
'        tTributacaoItemPV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        'Pesquisa no BD o item da Tributação relacionado ao item do pedido de vendas passado
'        With tTributacaoItemPV
'            lErro = Comando_ExecutarPos(alComando(0), "SELECT NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoItemPV WHERE NumIntDoc = ? ", 0, _
'            .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito, objItemPV.lNumIntDoc)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then Error 57796
'
'        lErro = Comando_BuscarPrimeiro(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57797
'
'        'Se não encontrou --> Erro
'        If lErro <> AD_SQL_SUCESSO Then Error 57798
'
'        'Exclui do BD o item da Tributação relacionado ao item do pedido de vendas passado
'        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoItemPV", alComando(0))
'        If lErro <> AD_SQL_SUCESSO Then Error 57799
'
'        'Insere em TributacaoItemPVBaixado o item da Tributação relacionado ao item do pedido de vendas passado
'        With tTributacaoItemPV
'            lErro = Comando_Executar(alComando(2), "INSERT INTO TributacaoItemPVBaixado (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemPV.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then Error 57800
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    TributacaoItemPV_Baixa = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItemPV_Baixa:
'
'    TributacaoItemPV_Baixa = Err
'
'    Select Case Err
'
'        Case 57795
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 57796, 57797
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMPEDIDO", Err)
'
'        Case 57798
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBITEMPEDIDO_NAO_ENCONTRADA", Err, objItemPV.lNumIntDoc)
'
'        Case 57799
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMPEDIDO", Err)
'
'        Case 57800
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152569)
'
'    End Select
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Function TributacaoComplPV_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Baixa a TributaçãoCompl relacionada ao pedido de vendas passado
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim alComando(0 To 3) As Long
'Dim lErro As Long
'Dim iIndice As Integer
'Dim tTribComplPV As typeTributacaoComplPV
'Dim sCodigo As String
'
'On Error GoTo Erro_TributacaoComplPV_Baixa
'
'    'Abertura dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 57801
'    Next
'
'    tTribComplPV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'    sCodigo = String(STRING_IPICODIGOPRODUTO_TRIBUTACAO, 0)
'
'    'Pesquisa no BD a TributaçãoCompl relacionada ao Pedido de Vendas passado
'    With tTribComplPV
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual FROM TributacaoComplPV WHERE FilialEmpresa = ? AND CodPedido = ? ", 0, _
'        .iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, sCodigo, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then Error 57802
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57803
'
'    'Enquanto encontrar relacionamento
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Exclui do BD a TributaçãoCompl
'        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoComplPV", alComando(0))
'        If lErro <> AD_SQL_SUCESSO Then Error 57804
'
'        'Insere em TributaçãoComplBaixado a tributação relacionada ao pedido de vendas passado
'        With tTribComplPV
'            lErro = Comando_Executar(alComando(3), "INSERT INTO TributacaoComplPVBaixado (FilialEmpresa, CodPedido, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'                objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then Error 57806
'
'        'Busca próxima TributaçãoCompl
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57805
'
'    Loop
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    TributacaoComplPV_Baixa = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoComplPV_Baixa:
'
'    TributacaoComplPV_Baixa = Err
'
'    Select Case Err
'
'        Case 57801
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 57802, 57803, 57805
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLPEDIDO", Err)
'
'        Case 57804
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLPEDIDO", Err, objPedidoVenda.lCodigo)
'
'        Case 57806
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152570)
'
'    End Select
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function PedidosBaixar_Batch(ByVal colPVInfo As Collection) As Long
'recebe colecao de ClassPVInfo
'transfere pedidos de venda p/pedidos de venda baixados

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 29) As Long
Dim lTransacao As Long
Dim objPVInfo As ClassPVInfo
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_PedidosBaixar_Batch

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33461
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33462

    'Percorre a coleção de PedVendas
    For Each objPVInfo In colPVInfo
    
        If objPVInfo.iMarcada = MARCADO Then
        
            Set objPedidoVenda = New ClassPedidoDeVenda
        
            'Passa o identificador do Pedido
            objPedidoVenda.iFilialEmpresa = giFilialEmpresa
            objPedidoVenda.lCodigo = objPVInfo.lCodPedido
            
            'Le o Pedido
            
            lErro = CF("PedidoVenda_Lock", objPedidoVenda, alComando(29))
            If lErro <> SUCESSO Then Error 64016
            
            lErro = CF("PV_Le_Customizado", objPedidoVenda)
            If lErro <> SUCESSO Then Error 64016
            
            'Lê os ítens do Pedido de Vendas
            lErro = CF("PedidoDeVenda_Le_Itens", objPedidoVenda)
            If lErro <> SUCESSO Then Error 25642
            
            'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
            lErro = ProdutoFilial_Subtrai_QuantPedida(objPedidoVenda, objPedidoVenda.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 58578
            
            'baixa o pedido
            lErro = CF("PedidoDeVenda_Baixar", alComando(), objPedidoVenda)
            If lErro <> SUCESSO Then Error 36724
            
        End If
            
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33463

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    PedidosBaixar_Batch = SUCESSO

    Exit Function

Erro_PedidosBaixar_Batch:

    PedidosBaixar_Batch = Err

    Select Case Err

        Case 25642, 58578, 64016  'tratado na rotina chamada
        
        Case 33461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 33462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 36724
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152571)

    End Select
    
    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ProdutoFilial_Adiciona_QuantPedida(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza adicionando a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim dQuantPedida As Double
Dim dQuantItem As Double
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemPedVenda As ClassItemPedido
Dim objProduto As New ClassProduto
Dim objProdutoIPV As New ClassProduto
Dim dFator As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iContador As Integer
Dim iQuantTotal As Integer

Dim sProduto As String
Dim sUnidadeMed As String
 
On Error GoTo Erro_ProdutoFilial_Adiciona_QuantPedida

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51387
    Next
    
    For Each objItemPedVenda In objPedidoVenda.colItensPedido
    
        objProdutoIPV.sCodigo = objItemPedVenda.sProduto
        
        lErro = CF("Produto_Le", objProdutoIPV)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 51392
        If lErro <> SUCESSO Then Error 51393
            
        For iContador = 0 To objItemPedVenda.colItensRomaneioGrade.Count
            
            If objItemPedVenda.colItensRomaneioGrade.Count = 0 Then
                sProduto = objItemPedVenda.sProduto
                dQuantItem = (objItemPedVenda.dQuantidade - objItemPedVenda.dQuantCancelada - objItemPedVenda.dQuantFaturada)
            Else
                If iContador = 0 Then iContador = 1
                sProduto = objItemPedVenda.colItensRomaneioGrade(iContador).sProduto
                dQuantItem = (objItemPedVenda.colItensRomaneioGrade(iContador).dQuantidade - objItemPedVenda.colItensRomaneioGrade(iContador).dQuantCancelada - objItemPedVenda.colItensRomaneioGrade(iContador).dQuantFaturada)
            End If
            
            objProduto.sCodigo = sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 51392
            If lErro <> SUCESSO Then Error 51393
        
            If objProdutoIPV.iKitVendaComp <> MARCADO Then
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedVenda.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            Else
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProduto.sSiglaUMVenda, dFator)
            End If
            If lErro <> SUCESSO Then Error 51394
            
            'Calcula a quantidade a ser adicionada a quantidade pedida do produto
            dQuantItem = dQuantItem * dFator
                        
            'Busca o Registro de Produto do Item e FiliaisProduto na Filial de Faturamento
            lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPedida FROM ProdutosFilial WHERE Produto = ? AND FilialEmpresa = ?", 0, dQuantPedida, sProduto, objPedidoVenda.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then Error 51388
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51389
            If lErro <> AD_SQL_SUCESSO Then Error 51390 'Não encontrou
                          
            'Atualiza a quantidade pedida em ProdutosFilial para o Produto e Filial
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ProdutosFilial SET QuantPedida = QuantPedida + ?", alComando(0), dQuantItem)
            If lErro <> AD_SQL_SUCESSO Then Error 51391
        
        Next
    
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ProdutoFilial_Adiciona_QuantPedida = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Adiciona_QuantPedida:

    ProdutoFilial_Adiciona_QuantPedida = Err
    
    Select Case Err
    
        Case 51387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51388, 51389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objItemPedVenda.iFilialEmpresa, objItemPedVenda.sProduto)
        
        Case 51390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE_FILIALFATURAMENTO", Err, objItemPedVenda.sProduto, objItemPedVenda.iFilialEmpresa)
        
        Case 51391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", Err, objItemPedVenda.iFilialEmpresa, objItemPedVenda.sProduto)

        Case 51392, 51394
        
        Case 51393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152572)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ProdutoFilial_Subtrai_QuantPedida(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal iFilialEmpresa As Integer) As Long
'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
 
Dim lErro As Long
Dim dQuantPedida As Double
Dim dQuantItem As Double
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemPedVenda As ClassItemPedido
Dim objProduto As New ClassProduto
Dim objProdutoIPV As New ClassProduto
Dim dFator As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iContador As Integer
Dim iQuantTotal As Integer
Dim sProduto As String
 
On Error GoTo Erro_ProdutoFilial_Subtrai_QuantPedida

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 51395
    Next
    
    For Each objItemPedVenda In objPedidoVenda.colItensPedido
    
        objProdutoIPV.sCodigo = objItemPedVenda.sProduto
        
        lErro = CF("Produto_Le", objProdutoIPV)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 51392
        If lErro <> SUCESSO Then Error 51393
    
        For iContador = 0 To objItemPedVenda.colItensRomaneioGrade.Count
        
            If objItemPedVenda.colItensRomaneioGrade.Count = 0 Then
                sProduto = objItemPedVenda.sProduto
                dQuantItem = (objItemPedVenda.dQuantidade - objItemPedVenda.dQuantCancelada - objItemPedVenda.dQuantFaturada)
            Else
                If iContador = 0 Then iContador = 1
                sProduto = objItemPedVenda.colItensRomaneioGrade(iContador).sProduto
                dQuantItem = (objItemPedVenda.colItensRomaneioGrade(iContador).dQuantidade - objItemPedVenda.colItensRomaneioGrade(iContador).dQuantCancelada - objItemPedVenda.colItensRomaneioGrade(iContador).dQuantFaturada)
            End If
    
            objProduto.sCodigo = sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 51400
            If lErro <> SUCESSO Then Error 51402
        
            If objProdutoIPV.iKitVendaComp <> MARCADO Then
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedVenda.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            Else
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProduto.sSiglaUMVenda, dFator)
            End If
            If lErro <> SUCESSO Then Error 51401
            
            'Calcula a quantidade a ser Subtraida a quantidade pedida do produto
            dQuantItem = dQuantItem * dFator
                        
            'Busca o Registro de Produto do Item e FiliaisProduto
            lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPedida FROM ProdutosFilial WHERE Produto = ? AND FilialEmpresa = ?", 0, dQuantPedida, sProduto, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then Error 51396
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51397
            If lErro <> AD_SQL_SUCESSO Then Error 51398 'Não encontrou
                          
            'Atualiza a quantidade pedida em ProdutosFilial para o Produto
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ProdutosFilial SET QuantPedida = QuantPedida - ?", alComando(0), dQuantItem)
            If lErro <> AD_SQL_SUCESSO Then Error 51399
                        
        Next
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ProdutoFilial_Subtrai_QuantPedida = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Subtrai_QuantPedida:

    ProdutoFilial_Subtrai_QuantPedida = Err
    
    Select Case Err
    
        Case 51395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51396, 51397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objItemPedVenda.iFilialEmpresa, objItemPedVenda.sProduto)
        
        Case 51398
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", Err, objItemPedVenda.sProduto, objItemPedVenda.iFilialEmpresa)
        
        Case 51399
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", Err, objItemPedVenda.iFilialEmpresa, objItemPedVenda.sProduto)

        Case 51400, 51401
        
        Case 51402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152573)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ProdutoFilial_Altera_QuantPedida(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_Anterior As ClassPedidoDeVenda) As Long
'Atualiza a quantidade pedida em produtos filial na alteração do pedido de venda passado

Dim lErro As Long
Dim objPedidoVenda_BD As New ClassPedidoDeVenda

On Error GoTo Erro_ProdutoFilial_Altera_QuantPedida

    objPedidoVenda_BD.lNumIntDoc = objPedidoVenda.lNumIntDoc
    objPedidoVenda_BD.lCodigo = objPedidoVenda.lCodigo
    objPedidoVenda_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
    'Lê o Pedido de Venda no BD
    lErro = CF("PedidoDeVenda_Le_Itens", objPedidoVenda_BD)
    If lErro <> SUCESSO Then Error 51403
    
    'Subtrai a quantidade pedida dos itens lidos do BD
    lErro = ProdutoFilial_Subtrai_QuantPedida(objPedidoVenda_BD, objPedidoVenda_Anterior.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 51404
    
    'Adiciona a quantidade pedida dos itens da tela
    lErro = ProdutoFilial_Adiciona_QuantPedida(objPedidoVenda)
    If lErro <> SUCESSO Then Error 51405
    
    ProdutoFilial_Altera_QuantPedida = SUCESSO
    
    Exit Function
    
Erro_ProdutoFilial_Altera_QuantPedida:

    ProdutoFilial_Altera_QuantPedida = Err
    
    Select Case Err
    
        Case 51403, 51404, 51405
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152574)
            
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Exclui_BD1(ByVal objPedVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long) As Long
'Exclui o Pedido de Venda da tabela de Pedido de Vendas mas não exclui a Tributação nem a contabilidade relacionado a esse pedido

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_PedidoVenda_Exclui_BD1

    'Busca o Pedido no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM PedidosDeVenda WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 57815

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57816
    If lErro = AD_SQL_SEM_DADOS Then Error 57817 'Não encontrou

    'Exclui o Pedido da Tabela de Pedidos
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PedidosDeVenda ", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 57818

    PedidoVenda_Exclui_BD1 = SUCESSO

    Exit Function

Erro_PedidoVenda_Exclui_BD1:

    PedidoVenda_Exclui_BD1 = Err

    Select Case Err

        Case 57815, 57816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedVenda.lCodigo)

        Case 57817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedVenda.lCodigo)

        Case 57818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDODEVENDA", Err, objPedVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152575)

    End Select

    Exit Function

End Function

Function ItensPV_Exclui1(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long
'Exclui os itens presentes na coleção de itens do Pedido de vendas passado, mas não
'exclui a tributação nem a contabilidade associada a esse Item e nem infoadicdocitem

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_ItensPV_Exclui1

    'Para cada item na coleção
    For Each objItemPV In objPedidoVenda.colItensPedido

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ItensPedidoDeVenda WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 57811

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57814
        If lErro = AD_SQL_SEM_DADOS Then Error 57812 'Não encontrou

        'Exclui o item da tabela de itens de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ItensPedidoDeVenda ", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 57813

    Next

    ItensPV_Exclui1 = SUCESSO

    Exit Function

Erro_ItensPV_Exclui1:

    ItensPV_Exclui1 = Err

    Select Case Err

        Case 57811, 57814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 57812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_ENCONTRADO", Err, objPedidoVenda.lNumIntDoc)

        Case 57813
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDIDODEVENDA", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152576)

    End Select

    Exit Function

End Function

Function PedidoDeVenda_Baixar_Pedido(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
    
Dim lErro As Long
Dim alComando(4 To 27) As Long
Dim lTransacao As Long
Dim iIndice As Integer

On Error GoTo Erro_PedidoDeVenda_Baixar_Pedido
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 62121
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 62122
    
    'Atualiza Subtraindo a quantPedida do Produto em UM de Venda na tabela ProdutosFilial
    lErro = ProdutoFilial_Subtrai_QuantPedida(objPedidoVenda, objPedidoVenda.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 62119
    
    'baixa o pedido
    lErro = CF("PedidoDeVenda_Baixar", alComando(), objPedidoVenda)
    If lErro <> SUCESSO Then Error 62120
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 62123
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoDeVenda_Baixar_Pedido = SUCESSO
    
    Exit Function
    
Erro_PedidoDeVenda_Baixar_Pedido:

    PedidoDeVenda_Baixar_Pedido = Err
    
    Select Case Err
        
        Case 62119, 62120
    
        Case 62121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 62123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152577)
            
    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'Function TributacaoPV_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Baixa Tributação relacionada ao pedido de vendas passado como parâmetro
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim lErro As Long
'
'On Error GoTo Erro_TributacaoPV_Desfaz_Baixa
'
''    'Desfaz a baixa da tributação PV
''    lErro = TributacaoPV_Desfaz_Baixa_BD(objPedidoVenda)
''    If lErro <> SUCESSO Then Error 62224
''
''    'Desfaz a baixa da tributação dos itens PV
''    lErro = TributacaoItemPV_Desfaz_Baixa(objPedidoVenda)
''    If lErro <> SUCESSO Then Error 62225
''
''    'Desfaz a baixa da tributação dos compolementos do PV
''    lErro = TributacaoComplPV_Desfaz_Baixa(objPedidoVenda)
''    If lErro <> SUCESSO Then Error 62226
'
'    TributacaoPV_Desfaz_Baixa = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Desfaz_Baixa:
'
'    TributacaoPV_Desfaz_Baixa = Err
'
'    Select Case Err
'
'        Case 62224, 62225, 62226
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152578)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function TributacaoPV_Desfaz_Baixa_BD(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Desfaz a Baixa Tributação do pedido de vendas passado
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim alComando(0 To 2) As Long
'Dim lErro As Long
'Dim iIndice As Integer
'Dim typeTributacaoPV As typeTributacaoPV
'
'On Error GoTo Erro_TributacaoPV_Desfaz_Baixa_BD
'
'    'Abertura dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 62227
'    Next
'
'    'Pesquisa no BD a Tributação Relacionada ao Pedido de Vendas
'    With typeTributacaoPV
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoPVBaixado WHERE FilialEmpresa = ? AND CodPedido = ?", 0, _
'            .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual, _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then Error 62228
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62229
'
'    'Se não encontrou --> Erro
'    If lErro <> AD_SQL_SUCESSO Then Error 62230
'
'    'Insere em TributacaoPVBaixado a Tributação do Pedido de Vendas passado
'    With typeTributacaoPV
'        lErro = Comando_Executar(alComando(2), "INSERT INTO TributacaoPV (FilialEmpresa, CodPedido, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'            "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
'            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'    If lErro <> SUCESSO Then Error 62232
'
'    'Exclui do BD a Tributação Relacionado ao Pedido de Vendas passado
'    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoPVBaixado", alComando(0))
'    If lErro <> AD_SQL_SUCESSO Then Error 62231
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    TributacaoPV_Desfaz_Baixa_BD = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Desfaz_Baixa_BD:
'
'    TributacaoPV_Desfaz_Baixa_BD = Err
'
'    Select Case Err
'
'        Case 62227
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 62228, 62229
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBPEDIDOBAIXADO", Err)
'
'        Case 62230
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBPEDIDO_NAO_ENCONTRADA", Err, objPedidoVenda.lCodigo)
'
'        Case 62231
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBPEDIDOBAIXADO", Err)
'
'        Case 62232
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152579)
'
'    End Select
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function TributacaoItemPV_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Desfaz a Baixa os Itens da Tributação do pedido de vendas passado
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim alComando(0 To 2) As Long
'Dim lErro As Long
'Dim iNaturezaOpManual As Integer
'Dim iIndice As Integer
'Dim objItemPV As ClassItemPedido
'Dim tTributacaoItemPV As typeTributacaoItemPV
'
'On Error GoTo Erro_TributacaoItemPV_Desfaz_Baixa
'
'    'Abertura dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 62233
'    Next
'
'    'Para cada item da coleção de pedido de vendas
'    For Each objItemPV In objPedidoVenda.colItensPedido
'
'        tTributacaoItemPV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        'Pesquisa no BD o item da Tributação relacionado ao item do pedido de vendas passado
'        With tTributacaoItemPV
'            lErro = Comando_ExecutarPos(alComando(0), "SELECT NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoItemPVBaixado WHERE NumIntDoc = ? ", 0, _
'            .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito, objItemPV.lNumIntDoc)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then Error 62234
'
'        lErro = Comando_BuscarPrimeiro(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62235
'
'        'Se não encontrou --> Erro
'        If lErro <> AD_SQL_SUCESSO Then Error 62236
'
'        'Insere em TributacaoItemPVBaixado o item da Tributação relacionado ao item do pedido de vendas passado
'        With tTributacaoItemPV
'            lErro = Comando_Executar(alComando(2), "INSERT INTO TributacaoItemPV (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ICMSSubstPercRedBase, ICMSSubstPercMVA, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemPV.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dICMSSubstPercRedBase, .dICMSSubstPercMVA, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then Error 62238
'
'        'Exclui do BD o item da Tributação relacionado ao item do pedido de vendas passado
'        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoItemPVBaixado", alComando(0))
'        If lErro <> AD_SQL_SUCESSO Then Error 62237
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    TributacaoItemPV_Desfaz_Baixa = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItemPV_Desfaz_Baixa:
'
'    TributacaoItemPV_Desfaz_Baixa = Err
'
'    Select Case Err
'
'        Case 62233
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 62234, 62235
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMPEDIDOBAIXADO", Err)
'
'        Case 62236
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBITEMPEDIDO_NAO_ENCONTRADA", Err, objItemPV.lNumIntDoc)
'
'        Case 62237
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMPEDIDOBAIXADO", Err)
'
'        Case 62238
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152580)
'
'    End Select
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function TributacaoComplPV_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
''Desfaz a Baixa da TributaçãoCompl relacionada ao pedido de vendas passado
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim alComando(0 To 3) As Long
'Dim lErro As Long
'Dim iIndice As Integer
'Dim tTribComplPV As typeTributacaoComplPV
'Dim sCodigo As String
'
'On Error GoTo Erro_TributacaoComplPV_Desfaz_Baixa
'
'    'Abertura dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 62239
'    Next
'
'    tTribComplPV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'    sCodigo = String(STRING_IPICODIGOPRODUTO_TRIBUTACAO, 0)
'
'    'Pesquisa no BD a TributaçãoComplBaixado relacionada ao Pedido de Vendas passado
'    With tTribComplPV
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual FROM TributacaoComplPVBaixado WHERE FilialEmpresa = ? AND CodPedido = ? ", 0, _
'        .iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, sCodigo, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then Error 62240
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62241
'
'    'Enquanto encontrar relacionamento
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Insere em TributaçãoCompl a tributação relacionada ao pedido de vendas passado
'        With tTribComplPV
'            lErro = Comando_Executar(alComando(3), "INSERT INTO TributacaoComplPV(FilialEmpresa, CodPedido, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'                objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then Error 62244
'
'        'Exclui do BD a TributaçãoComplBaixado
'        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TributacaoComplPVBaixado", alComando(0))
'        If lErro <> AD_SQL_SUCESSO Then Error 62243
'
'        'Busca próxima TributaçãoComplBaixado
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62242
'
'    Loop
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    TributacaoComplPV_Desfaz_Baixa = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoComplPV_Desfaz_Baixa:
'
'    TributacaoComplPV_Desfaz_Baixa = Err
'
'    Select Case Err
'
'        Case 62239
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 62240, 62241, 62242
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLPEDIDOBAIXADO", Err)
'
'        Case 62243
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLPEDIDOBAIXADO", Err, objPedidoVenda.lCodigo)
'
'        Case 62244
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLPV", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152581)
'
'    End Select
'
'    'Fechamento dos comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function ParcelasPV_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Desfaz a baixa das parcelas de pedido de venda

Dim lErro As Long

On Error GoTo Erro_ParcelasPV_Desfaz_Baixa

    'Lê as Parcelas baixados do pedido
    lErro = CF("ParcelasPV_Baixados_Le", objPedidoVenda)
    If lErro <> SUCESSO Then Error 62251
    
    'Inclui as parcelas na tabela das não baixadas
    lErro = CF("ParcelasPV_Grava", objPedidoVenda)
    If lErro <> AD_SQL_SUCESSO Then Error 62252
    
    'Exclui as parcelas do pedido da tabela das parc baixadas
    lErro = ParcelasPVBaixado_Exclui(objPedidoVenda, lComando, lComando1)
    If lErro <> SUCESSO Then Error 62253
    
    ParcelasPV_Desfaz_Baixa = SUCESSO
    
    Exit Function

Erro_ParcelasPV_Desfaz_Baixa:

    ParcelasPV_Desfaz_Baixa = Err
    
    Select Case Err
    
        Case 62251, 62252, 62253
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152582)
            
    End Select
    
    Exit Function

End Function

Private Function ParcelasPVBaixado_Exclui(ByVal objPedidoDeVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Exclui as Parcelas Baixadas do pedido passado

Dim lErro As Long
Dim lCodPedido As Long

On Error GoTo Erro_ParcelasPVBaixado_Exclui

    'BUsca as Parcelas do PV passado
    lErro = Comando_ExecutarPos(lComando, "SELECT PedidoDeVenda FROM ParcelasPedidoDeVendaBaixado WHERE PedidoDeVenda = ? AND FilialEmpresa = ?", 0, lCodPedido, objPedidoDeVenda.lCodigo, objPedidoDeVenda.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 62254
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62255
    'Para cada parcela encontrada
    Do While lErro = AD_SQL_SUCESSO
        'Exclui a parcela do BD
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasPedidoDeVendaBaixado", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 62257
        
        'Busca a próxima parcela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62256
    Loop
    
    ParcelasPVBaixado_Exclui = SUCESSO
    
    Exit Function
    
Erro_ParcelasPVBaixado_Exclui:

    ParcelasPVBaixado_Exclui = Err
    
    Select Case Err
    
        Case 62254, 62255, 62256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPEDIDODEVENDA_BAIXADAS", Err, objPedidoDeVenda.lCodigo)
        
        Case 62257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASPEDIDODEVENDA_BAIXADAS", Err, objPedidoDeVenda.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152583)
            
    End Select
    
    Exit Function
    
End Function

Function ComissoesPV_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Desfaz a Baixa das comissoes do PV

Dim lErro As Long

On Error GoTo Erro_ComissoesPV_Desfaz_Baixa
    
    'Lê as comissões baixadas do PV
    lErro = CF("ComissoesPV_Baixadas_Le", objPedidoVenda)
    If lErro <> SUCESSO Then Error 62258
    
    'Inclui as comissões na tabela das comiss. não baixadas
    lErro = CF("ComissoesPV_Grava", objPedidoVenda, True)
    If lErro <> SUCESSO Then Error 62259
    
    'Exclui as comissões baixadas do PV
    lErro = ComissoesPVBaixado_Exclui(objPedidoVenda, lComando, lComando1)
    If lErro <> SUCESSO Then Error 62260
    
    ComissoesPV_Desfaz_Baixa = SUCESSO
    
    Exit Function
    
Erro_ComissoesPV_Desfaz_Baixa:

    ComissoesPV_Desfaz_Baixa = Err
    
    Select Case Err
    
        Case 62258, 62259, 62260
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152584)
            
    End Select
            
    Exit Function

End Function

Private Function ComissoesPVBaixado_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Exclui as comissoes baixadados do pedido passado

Dim lErro As Long
Dim lCodPedido As Long

On Error GoTo Erro_ComissoesPVBaixado_Exclui

    'BUsca as comissoes do pedido passado
    lErro = Comando_ExecutarPos(lComando, "SELECT PedidoDeVendas FROM ComissoesPedVendasbaixados WHERE PedidodeVendas =? AND FilialEmpresa = ? ", 0, lCodPedido, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 62261
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62262
    
    'Para cada comissão encontrada
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Exclui a comissão do BD
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ComissoesPedVendasbaixados", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 62263
        
        'BUsca a próxima comissão
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62264
    
    Loop
    
    ComissoesPVBaixado_Exclui = SUCESSO
    
    Exit Function
    
Erro_ComissoesPVBaixado_Exclui:

    ComissoesPVBaixado_Exclui = Err
    
    Select Case Err
    
        Case 62261, 62262, 62264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESPEDVENDASBAIXADOS", Err, objPedidoVenda.lCodigo)
            
        Case 62263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESPEDVENDASBAIXADOS", Err, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152585)
            
    End Select
        
    Exit Function

End Function

Function BloqueiosPV_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Desfaz a baixa nos Bloqueios de pedido de Venda

Dim lErro As Long
Dim lCodPedido As Long
Dim objBloqueioPV As ClassBloqueioPV
Dim bJaPossuiBloq As Boolean

On Error GoTo Erro_BloqueiosPV_Desfaz_Baixa

    'Lê os Bloqueios PV já baixados
    lErro = CF("BloqueiosPV_Baixados_Le", objPedidoVenda)
    If lErro <> SUCESSO Then Error 62265
    
    'Inclui os bloqueios lidos na tabela dos bloqueios não baixados
    lErro = CF("BloqueiosPV_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then Error 62266
        
    'Exclui os bloqueios da tabela dos baixados
    lErro = BloqueiosPVBaixado_Exclui(objPedidoVenda, lComando, lComando1)
    If lErro <> SUCESSO Then Error 62267
    
    BloqueiosPV_Desfaz_Baixa = SUCESSO
    
    Exit Function
    
Erro_BloqueiosPV_Desfaz_Baixa:

    BloqueiosPV_Desfaz_Baixa = Err
    
    Select Case Err
        
        Case 62265, 62266, 62267
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152586)
            
    End Select
        
    Exit Function

End Function

Private Function BloqueiosPVBaixado_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long, ByVal lComando1 As Long) As Long
'Exclui os bloqueios PV baixados

Dim lErro As Long
Dim lCodPedido As Long

On Error GoTo Erro_BloqueiosPVBaixado_Exclui

    'BUsca os Bloqueios baixados do PV passado
    lErro = Comando_ExecutarPos(lComando, "SELECT PedidoDeVenda FROM BloqueiosPVBaixados WHERE PedidoDeVenda = ? AND FilialEmpresa = ? ", 0, lCodPedido, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 62268
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62269
    
    'Para cada bloqueio lido
    Do While lErro <> AD_SQL_SEM_DADOS
        'Exclui o bloqueio do BD
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BloqueiosPVBaixados", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 62271
        'Busca o próximo bloqueio
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62270
        
    Loop

    BloqueiosPVBaixado_Exclui = SUCESSO
    
    Exit Function
    
Erro_BloqueiosPVBaixado_Exclui:

    BloqueiosPVBaixado_Exclui = Err
    
    Select Case Err
    
        Case 62268, 62269, 62270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPVBAIXADOS", Err, objPedidoVenda.lCodigo)

        Case 62271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOSPVBAIXADOS", Err, objPedidoVenda.lCodigo)
            
    End Select
    
    Exit Function

End Function

Public Function Verifica_Situacao(ByVal iSituacaoAnterior As Integer, ByVal iSituacaoNova As Integer) As Long
'Verifica se a nova situacao pode ser efetivada, em função do que era a situação anterior

Dim lErro As Long

On Error GoTo Erro_Verifica_Situacao

    Select Case iSituacaoNova
    
        Case ITEMOP_SITUACAO_NORMAL
        
            If (iSituacaoAnterior = ITEMOP_SITUACAO_SACR) Then Error 41650
        
        Case ITEMOP_SITUACAO_DESAB

            If (iSituacaoAnterior = ITEMOP_SITUACAO_SACR) Or (iSituacaoAnterior = ITEMOP_SITUACAO_BAIXADA) Then Error 41651

        Case ITEMOP_SITUACAO_SACR

            If (iSituacaoAnterior = ITEMOP_SITUACAO_BAIXADA) Then Error 41652
        
    End Select

    Verifica_Situacao = SUCESSO
    
    Exit Function
    
Erro_Verifica_Situacao:

    Verifica_Situacao = Err
    
    Select Case Err
        
        Case 41650, 41651, 41652
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152587)

    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Grava(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ou atualiza uma ordem de producao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OrdemDeProducao_Grava

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 31522

    lErro = CF("OrdemDeProducao_Grava0", objOrdemDeProducao)
    If lErro <> AD_SQL_SUCESSO Then Error 62560

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 31532

    OrdemDeProducao_Grava = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava:

    OrdemDeProducao_Grava = Err

    Select Case Err

        Case 31522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 62560

        Case 31532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152588)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function OrdemDeProducao_Grava_EmTrans(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ou atualiza uma ordem de producao

Dim lErro As Long
Dim iIndice As Integer, lNumIntDocItem As Long, lEmpenho As Long
Dim sDescricao As String, lCodPrestador As Long
Dim dtDataEmissao As Date
Dim iNumItens As Integer
Dim iNumItensBaixados As Integer
Dim objItemOP As New ClassItemOP
Dim alComando(16) As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objOP As New ClassOrdemDeProducao
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemOPGrade As ClassItemOP
Dim bGerarOPs As Boolean
'Dim X As New ClassPVGrava

'Inserido por Jorge Specian
'---------------------------
Dim lCodigoNumerico As Long
'---------------------------

On Error GoTo Erro_OrdemDeProducao_Grava_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 31521
    Next
    
    bGerarOPs = False

    'Leitura da tabela de Ordens de Produção Baixadas
    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumItens FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ?", dtDataEmissao, iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 31523

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31524

    'se já existe uma OP baixada com este codigo => reativa a OP
    If lErro = AD_SQL_SUCESSO Then

        'Se o status da Ordem de Producao trazido da tela não for baixado ==> erro
        If objOrdemDeProducao.iStatusOP <> ITEMOP_SITUACAO_BAIXADA Then gError 82826

        lErro = OrdemDeProducaoBaixada_Reativar(objOrdemDeProducao, alComando)
        If lErro <> SUCESSO Then gError 82817

    Else

        'Leitura da tabela de Ordens de Produção
        lErro = Comando_ExecutarPos(alComando(1), "SELECT DataEmissao, NumItens, NumItensBaixados, CodPrestador FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, dtDataEmissao, iNumItens, iNumItensBaixados, lCodPrestador, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 31526

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31527

        'Existe a OP no BD
        If lErro = AD_SQL_SUCESSO Then

            '??catharine(apenas a linha do if abaixo)
            If objOrdemDeProducao.iStatusOP = ITEMOP_SITUACAO_BAIXADA Then gError 82827

            'Lock na OP
            lErro = Comando_LockExclusive(alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 22968

            'atualiza os itens pré-existentes da OP
            lErro = OrdemDeProducao_Grava4(objOrdemDeProducao)
            If lErro <> SUCESSO Then gError 22969

            'se trocou a DataEmissao, NumItens, NumItensBaixados ou CodPrestador
            'If objOrdemDeProducao.dtDataEmissao <> dtDataEmissao Or objOrdemDeProducao.iNumItens <> iNumItens Or objOrdemDeProducao.iNumItensBaixados <> iNumItensBaixados Or objOrdemDeProducao.lCodPrestador <> lCodPrestador Then

                If objOrdemDeProducao.lCodPrestador <> lCodPrestador And objOrdemDeProducao.lCodPrestador <> 0 Then
                
                    lErro = CF("PrestServ_Lock", objOrdemDeProducao.lCodPrestador)
                    If lErro <> SUCESSO Then gError 124247
                    
                End If
                
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE OrdensDeProducao SET DataEmissao = ?, NumItens = ?, NumItensBaixados = ?, CodPrestador = ?, TipoTerc = ?, CodTerc = ?, FilialTerc = ?, EscaninhoTerc = ?, IgnoraEst = ?", alComando(1), objOrdemDeProducao.dtDataEmissao, objOrdemDeProducao.iNumItens, objOrdemDeProducao.iNumItensBaixados, objOrdemDeProducao.lCodPrestador, objOrdemDeProducao.iTipoTerc, objOrdemDeProducao.lCodTerc, objOrdemDeProducao.iFilialTerc, objOrdemDeProducao.iEscaninhoTerc, objOrdemDeProducao.iIgnoraEst)
                If lErro <> AD_SQL_SUCESSO Then gError 22970

            'End If

        Else 'Nao existe a OP no BD

            If objOrdemDeProducao.lCodPrestador <> 0 Then
            
                lErro = CF("PrestServ_Lock", objOrdemDeProducao.lCodPrestador)
                If lErro <> SUCESSO Then gError 124248
            
            End If
            
            'Inserido por Jorge Specian
            '--------------------------
            If IsNumeric(objOrdemDeProducao.sCodigo) Then
                lCodigoNumerico = StrParaLong(objOrdemDeProducao.sCodigo)
            End If
            '--------------------------
            
            'Insere nova OP
            lErro = Comando_Executar(alComando(2), "INSERT INTO OrdensDeProducao (Codigo, FilialEmpresa, DataEmissao, NumItens, NumItensBaixados, OPGeradora, CodPrestador, Tipo, CodigoNumerico, Usuario, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.dtDataEmissao, objOrdemDeProducao.iNumItens, 0, objOrdemDeProducao.sOPGeradora, objOrdemDeProducao.lCodPrestador, objOrdemDeProducao.iTipo, lCodigoNumerico, gsUsuario, objOrdemDeProducao.iTipoTerc, objOrdemDeProducao.lCodTerc, objOrdemDeProducao.iFilialTerc, objOrdemDeProducao.iEscaninhoTerc, objOrdemDeProducao.iIgnoraEst) 'Alterado por Jorge Specian
            If lErro <> AD_SQL_SUCESSO Then gError 31528
    
        End If
        
        'Se a OP for gerar outra OP
        '#############################
        'ALTERADO POR WAGNER
        If objOrdemDeProducao.iGeraOPsArvore <> MARCADO And objOrdemDeProducao.iGeraOP = MARCADO Then
        '#############################

            objOP.sCodigo = objOrdemDeProducao.sCodOPGerar
            objOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa

            lErro = CF("OrdemProducao_Le", objOP)
            If lErro <> SUCESSO And lErro <> 30368 Then gError 62631
            If lErro = SUCESSO Then gError 62632

        End If
        
        'Se a OP for gerar RC ou OP p/subprodutos
        '#############################
        'ALTERADO POR WAGNER
        If objOrdemDeProducao.iGeraOPsArvore <> MARCADO And (objOrdemDeProducao.iGeraOP = MARCADO Or objOrdemDeProducao.iGeraReqCompra = MARCADO) Then
        '#############################
        
            bGerarOPs = True
'            lErro = OrdemProducao_Grava_OPSub_ReqCom(objOrdemDeProducao)
'            If lErro <> SUCESSO Then gError 62559
        
        End If

    End If

    'obtem o proximo codigo para empenhos e para NumIntItemOP
    lErro = OrdemDeProducao_Grava3(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
    If lErro <> AD_SQL_SUCESSO Then gError 22962

    'inclui os novos itens da OP
    For Each objItemOP In objOrdemDeProducao.colItens
        
        If objItemOP.lNumIntDoc = 0 Then

            If objItemOP.iPossuiGrade = DESMARCADO Then
                
                'Verifica o estoque máximo do produto utilizado
                lErro = Verifica_EstoqueMaximo(objItemOP)
                If lErro <> SUCESSO And lErro <> 41335 Then gError 40663
    
                If lErro = 41335 Then
    
                    'Confirmar se vai gravar com a Quantidade maior que a quantidade maxima do estoque
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ESTOQUEMAXIMO", objItemOP.sProduto)
    
                    If vbMsgRes = vbNo Then gError 40622
    
                End If

            End If

            objItemOP.lNumIntDoc = lNumIntDocItem
            
            lErro = CF("OrdemDeProducao_Grava1", objItemOP, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
            If lErro <> SUCESSO Then gError 31531

            lNumIntDocItem = lNumIntDocItem + 1

        End If
        
        For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
           
            'se ainda nao estiver cadastrado ==> insere
            If objItemRomaneioGrade.lNumIntDoc = 0 Then

                Set objItemOPGrade = New ClassItemOP
    
                objItemOPGrade.dQuantidade = objItemRomaneioGrade.dQuantidade
                objItemOPGrade.iAlmoxarifado = objItemRomaneioGrade.iAlmoxarifado
                objItemOPGrade.sProdutoPai = objItemOP.sProduto
                objItemOPGrade.sProduto = objItemRomaneioGrade.sProduto
                objItemOPGrade.sVersao = objItemRomaneioGrade.sVersao
                objItemOPGrade.lNumIntDoc = lNumIntDocItem
                objItemOPGrade.lNumItemOP = objItemOP.lNumIntDoc
                objItemOPGrade.iFilialPedido = objItemOP.iFilialPedido
                objItemOPGrade.lCodPedido = objItemOP.lCodPedido
                objItemOPGrade.lNumIntOrigem = objItemOP.lNumIntOrigem
                objItemOPGrade.iOrigemPedido = objItemOP.iOrigemPedido
                objItemOPGrade.iClasseUM = objItemOP.iClasseUM
                objItemOPGrade.sSiglaUMEstoque = objItemOP.sSiglaUMEstoque
                objItemOPGrade.sSiglaUM = objItemOP.sSiglaUM
                objItemOPGrade.iFilialEmpresa = objItemOP.iFilialEmpresa
                objItemOPGrade.sCodigo = objItemOP.sCodigo
                objItemOPGrade.iItem = objItemOP.iItem
                objItemOPGrade.sCcl = objItemOP.sCcl
                objItemOPGrade.dtDataInicioProd = objItemOP.dtDataInicioProd
                objItemOPGrade.dtDataFimProd = objItemOP.dtDataFimProd
                objItemOPGrade.iPrioridade = objItemOP.iPrioridade
                objItemOPGrade.iSituacao = objItemOP.iSituacao
                objItemOPGrade.iDestinacao = objItemOP.iDestinacao
                objItemOPGrade.iBeneficiamento = objItemOP.iBeneficiamento
                objItemOPGrade.lNumIntEquipamento = objItemOP.lNumIntEquipamento
                objItemOPGrade.lNumIntEquipamento2 = objItemOP.lNumIntEquipamento2
                objItemOPGrade.iPossuiGrade = DESMARCADO
                objItemOPGrade.iTipo = objItemOP.iTipo
    
                'Verifica o estoque máximo do produto utilizado
                lErro = Verifica_EstoqueMaximo(objItemOPGrade)
                If lErro <> SUCESSO And lErro <> 41335 Then gError 126705
    
                If lErro = 41335 Then
    
                    'Confirmar se vai gravar com a Quantidade maior que a quantidade maxima do estoque
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ESTOQUEMAXIMO", objItemOPGrade.sProduto)
    
                    If vbMsgRes = vbNo Then gError 126706
    
                End If
    
                lErro = CF("OrdemDeProducao_Grava1", objItemOPGrade, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
                If lErro <> SUCESSO Then gError 126521
    
                lNumIntDocItem = lNumIntDocItem + 1

            End If

        Next

    Next

    'atualiza o proximo codigo para lNumIntDocItem e lEmpenho
    lErro = OrdemDeProducao_Grava5(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
    If lErro <> AD_SQL_SUCESSO Then gError 19471

    If objOrdemDeProducao.iNumItens = objOrdemDeProducao.iNumItensBaixados Then

        'Baixa a ordem de producao, seus itens e seus empenhos
        lErro = CF("OrdemDeProducao_Baixa", objOrdemDeProducao.sCodigo)
        If lErro <> SUCESSO Then gError 55273

    End If

    'Inserido por Jorge Specian - 10/05/2005
    '---------------------------------------
    lErro = CF("OrdemDeProducao_Grava_Operacoes_EmTrans", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 134978
    '---------------------------------------
    
    '##############################################################
    'Alterado por Wagner 09/05/2006
    'Teve de mudar de posição para já ter o NumIntDoc do item de OP
    If bGerarOPs Then
    
        lErro = OrdemProducao_Grava_OPSub_ReqCom(objOrdemDeProducao)
        If lErro <> SUCESSO Then gError 62559
            
    End If
    '##############################################################
    
    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Grava", objOrdemDeProducao.objProjetoInfo, PRJ_CR_TIPO_OP, 0, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 185840
    '############################################

    lErro = CF("OrdemDeProducao_Grava_Customizado", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 185840

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava_EmTrans:

    OrdemDeProducao_Grava_EmTrans = gErr

    Select Case gErr

        Case 22968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORDENSDEPRODUCAO", gErr)

        Case 22970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ORDENSDEPRODUCAO", gErr, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)

        Case 27509, 27510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", gErr)

        Case 31521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 31523, 31524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS", gErr)

        Case 31526, 31527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)

        Case 31528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDEPRODUCAO", gErr)

        Case 31530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG", gErr)

        Case 19471, 22962, 22969, 31531, 40622, 40663, 55273, 62558, 62559, 62631, 82817, 126521, 124127, 124128, 126705, 126706

        Case 134977, 134978, 185840

        Case 62632
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOP.sCodigo)

        Case 82826
            Call Rotina_Erro(vbOKOnly, "ERRO_STATUS_OP_ABERTO", gErr)

        Case 82827
            Call Rotina_Erro(vbOKOnly, "ERRO_STATUS_OP_BAIXADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152589)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrdemDeProducao_Grava1(ByVal objItemOP As ClassItemOP, alComando() As Long, ByVal dtDataEmissao As Date, lEmpenho As Long) As Long
'Inclui um item numa OP
'Tem que ser chamada dentro de transação

Dim lErro As Long
Dim iPCP As Integer
Dim iCodigo As Integer
Dim iItem As Integer
Dim iControleEstoque As Integer
Dim iNivel As Integer
Dim objProdutoKit As New ClassProdutoKit
Dim colProdutos As New Collection
Dim dQuantidadeRaiz As Double
Dim objItemPV As New ClassItemPedido
Dim sCcl As String, objEstoqueProduto As New ClassEstoqueProduto
Dim iClasseUM As Integer
Dim sSiglaUMEstoque As String
Dim iCompras As Integer

On Error GoTo Erro_OrdemDeProducao_Grava1

    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    'Leitura da tabela Produtos
    lErro = Comando_ExecutarLockado(alComando(4), "SELECT PCP, SiglaUMEstoque, ClasseUM, Compras FROM Produtos WHERE Codigo = ? ", iPCP, sSiglaUMEstoque, iClasseUM, iCompras, objItemOP.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 31533

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31534

    If lErro = AD_SQL_SEM_DADOS Then gError 31535

    lErro = Comando_LockShared(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then gError 31536

    If iPCP = PRODUTO_PCP_NAOPODE Or iCompras <> PRODUTO_PRODUZIVEL Then gError 55277

    objItemOP.sSiglaUMEstoque = sSiglaUMEstoque
    objItemOP.iClasseUM = iClasseUM

    If objItemOP.iPossuiGrade = DESMARCADO Then

        'garantir existencia do almoxarifado
        lErro = Comando_ExecutarLockado(alComando(5), "SELECT Codigo FROM Almoxarifado WHERE Codigo = ? ", iCodigo, objItemOP.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 31537
    
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31538
    
        If lErro = AD_SQL_SEM_DADOS Then gError 31539
    
        lErro = Comando_LockShared(alComando(5))
        If lErro <> AD_SQL_SUCESSO Then gError 31540

    End If

    'garantir a existencia da Ccl
    If objItemOP.sCcl <> "" Then

        sCcl = String(STRING_CCL, 0)

        lErro = Comando_ExecutarLockado(alComando(6), "SELECT Ccl FROM Ccl WHERE Ccl = ? ", sCcl, objItemOP.sCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 31541

        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31542

        If lErro = AD_SQL_SEM_DADOS Then gError 31543

        lErro = Comando_LockShared(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 31544

    End If

    If objItemOP.lCodPedido <> 0 And objItemOP.iPossuiGrade = DESMARCADO Then

        objItemPV.iFilialEmpresa = objItemOP.iFilialPedido
        objItemPV.lCodPedido = objItemOP.lCodPedido
        objItemPV.sProduto = objItemOP.sProduto
        objItemPV.iTipoOP = objItemOP.iTipo

        'aumenta a quantidade em OP do Pedido de Venda
        lErro = OrdemDeProducao_Grava6(alComando(14), alComando(15), objItemPV, objItemOP.dQuantidade)
        If lErro <> SUCESSO And lErro <> 41346 Then gError 40664

        If lErro = 41346 Then gError 40666

        objItemOP.lNumIntOrigem = objItemPV.lNumIntDoc
        objItemOP.iOrigemPedido = objItemPV.iOrigemPedido

    End If

    'se é um item de grade
    If objItemOP.iPossuiGrade = MARCADO Then

        lErro = Comando_Executar(alComando(7), "INSERT INTO ItensOPGrade (NumIntDoc, FilialEmpresa, Codigo, Item, Produto, FilialPedido, " & _
        "CodPedido, NumIntOrigem, OrigemPedido, Almoxarifado, Beneficiamento, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, " & _
        "Situacao, Destinacao, Versao, NumIntEquipamento, NumIntEquipamento2, Tipo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
        objItemOP.lNumIntDoc, objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.iItem, objItemOP.sProduto, objItemOP.iFilialPedido, objItemOP.lCodPedido, objItemOP.lNumIntOrigem, objItemOP.iOrigemPedido, _
        objItemOP.iAlmoxarifado, objItemOP.iBeneficiamento, objItemOP.sCcl, objItemOP.sSiglaUM, objItemOP.dQuantidade, objItemOP.dtDataInicioProd, objItemOP.dtDataFimProd, objItemOP.iPrioridade, _
        objItemOP.dQuantidadeProd, objItemOP.iSituacao, objItemOP.iDestinacao, objItemOP.sVersao, objItemOP.lNumIntEquipamento, objItemOP.lNumIntEquipamento2, objItemOP.iTipo)
        If lErro <> SUCESSO Then gError 126692

    Else
    
        'Insere item da OP
        '##############################################
        'ALTERADO POR WAGNER
        lErro = Comando_Executar(alComando(7), "INSERT INTO ItensOrdemProducao (NumIntItemOPGrade, NumIntDoc, FilialEmpresa, Codigo, Item, Produto, FilialPedido, " & _
        "CodPedido, NumIntOrigem, OrigemPedido, Almoxarifado, Beneficiamento, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, " & _
        "Situacao, Destinacao, Versao, NumIntEquipamento, NumIntEquipamento2, Tipo, NumIntDocPai,ProduzLogo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
        objItemOP.lNumItemOP, objItemOP.lNumIntDoc, objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.iItem, objItemOP.sProduto, objItemOP.iFilialPedido, objItemOP.lCodPedido, objItemOP.lNumIntOrigem, objItemOP.iOrigemPedido, _
        objItemOP.iAlmoxarifado, objItemOP.iBeneficiamento, objItemOP.sCcl, objItemOP.sSiglaUM, objItemOP.dQuantidade, objItemOP.dtDataInicioProd, objItemOP.dtDataFimProd, objItemOP.iPrioridade, _
        objItemOP.dQuantidadeProd, objItemOP.iSituacao, objItemOP.iDestinacao, objItemOP.sVersao, objItemOP.lNumIntEquipamento, objItemOP.lNumIntEquipamento2, objItemOP.iTipo, objItemOP.lNumIntDocPai, objItemOP.iProduzLogo)
        If lErro <> SUCESSO Then gError 31551
        '##############################################

    End If

    If objItemOP.iPossuiGrade = DESMARCADO And objItemOP.iTipo = OP_TIPO_OP Then

        'cria os empenhos
        lErro = CF("OrdemDeProducao_Grava2", objItemOP, alComando, dtDataEmissao, lEmpenho)
        If lErro <> SUCESSO Then gError 31552
    
        'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
        objEstoqueProduto.sProduto = objItemOP.sProduto
        objEstoqueProduto.iAlmoxarifado = objItemOP.iAlmoxarifado
    
        lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, objItemOP.dQuantidade)
        If lErro <> SUCESSO Then gError 41552

    End If

    OrdemDeProducao_Grava1 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava1:

    OrdemDeProducao_Grava1 = gErr

    Select Case gErr

        Case 31533, 31534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 31535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemOP.sProduto)

        Case 31536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", gErr)

        Case 31537, 31538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", gErr)

        Case 31539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objItemOP.iAlmoxarifado)

        Case 31540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO", gErr)

        Case 31541, 31542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)

        Case 31543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemOP.iAlmoxarifado)

        Case 31544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr)

        Case 31549, 31550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)

       Case 31551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSORDENSDEPRODUCAO", gErr)

        Case 31552, 40664, 41552

        Case 40666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemOP.lCodPedido, objItemOP.sProduto)

        Case 55277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, objItemOP.sProduto)

        Case 126692
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSOPGRADE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152590)

    End Select

    Exit Function

End Function

Function OrdemDeProducao_Grava2(ByVal objItemOP As ClassItemOP, alComando() As Long, ByVal dtDataEmissao As Date, lEmpenho As Long) As Long
'cria empenhos correspondentes ao objItemOP
'IMPORTANTE: Tem que ser chamada dentro de transação

Dim lErro As Long
Dim iNivel As Integer
Dim iClasseUMRaiz As Integer, iAlmoxPadraoProd As Integer
Dim dFator As Double, dFatorRaiz As Double, dFatorItemOP As Double
Dim dtDataTEMP As Date
Dim sProduto As String
Dim sSiglaUMEstoqueRaiz As String, sSiglaUMKit As String
Dim objProdutoKit As New ClassProdutoKit
Dim tProdutoKit As typeProdutoKit, objEmpenho As New ClassEmpenho
Dim tProdutoKitRaiz As typeProdutoKit
Dim colItemOP As New Collection
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos
Dim dQuantidadeArr As Double

On Error GoTo Erro_OrdemDeProducao_Grava2

    lErro = Comando_ExecutarLockado(alComando(8), "SELECT Data FROM Kit WHERE ProdutoRaiz = ? AND Versao = ?", dtDataTEMP, objItemOP.sProduto, objItemOP.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 31553

    lErro = Comando_BuscarPrimeiro(alComando(8))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31554

    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_LockShared(alComando(8))
        If lErro <> AD_SQL_SUCESSO Then gError 31555

        sSiglaUMEstoqueRaiz = String(STRING_UM_SIGLA, 0)
        tProdutoKitRaiz.sUnidadeMed = String(STRING_UM_SIGLA, 0)

        'pega dados do raiz (nivel = 0)
        lErro = Comando_Executar(alComando(10), "SELECT Quantidade, ProdutoKit.UnidadeMed, SiglaUMEstoque, ClasseUM FROM ProdutoKit, Produtos WHERE ProdutoKit.Produto = ? AND ProdutoKit.Versao = ? AND ProdutoKit.Nivel = ? AND ProdutoKit.Produto = Produtos.Codigo", tProdutoKitRaiz.dQuantidade, tProdutoKitRaiz.sUnidadeMed, sSiglaUMEstoqueRaiz, iClasseUMRaiz, objItemOP.sProduto, objItemOP.sVersao, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 31556

        lErro = Comando_BuscarPrimeiro(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 31557

        'descobre o fator de conversao da UM p/UM de estoque do produto raiz
        lErro = CF("UM_Conversao_Trans", iClasseUMRaiz, tProdutoKitRaiz.sUnidadeMed, sSiglaUMEstoqueRaiz, dFatorRaiz)
        If lErro <> SUCESSO Then gError 55248

        'descobre o fator de conversao da UM p/UM de estoque do produto ao qual o item da OP se refere
        lErro = CF("UM_Conversao_Trans", objItemOP.iClasseUM, objItemOP.sSiglaUM, objItemOP.sSiglaUMEstoque, dFatorItemOP)
        If lErro <> SUCESSO Then gError 55249

        'Alteracao Daniel em 29/07/2002
        objProdutoKitProdutos.sProdutoRaiz = objItemOP.sProduto
        objProdutoKitProdutos.sVersao = objItemOP.sVersao
        objProdutoKitProdutos.dQuantidade = objItemOP.dQuantidade
        objProdutoKitProdutos.iClasseUM = objItemOP.iClasseUM
        objProdutoKitProdutos.sUnidadeMed = objItemOP.sSiglaUM
        
        'Obtem uma Colecao com os itens do kit para empenho (a embalagem do ProdutoRaiz será o ultimo elemento da colecao)
        lErro = CF("OrdemProducao_Le_Col_Empenho", objProdutoKitProdutos, colItemOP)
        If lErro <> SUCESSO And lErro <> 106393 Then gError 106399
        
        'Se nao Encontrou => Erro
        If lErro = 106393 Then gError 106400

        'Para cada elemento da Colecao...
        For Each objProdutoKitProdutos In colItemOP

            If objProdutoKitProdutos.iControleEstoque = PRODUTO_CONTROLE_RESERVA Or objProdutoKitProdutos.iControleEstoque = PRODUTO_CONTROLE_ESTOQUE Then

                'le a tabela produtofilial para tentar descobrir o almoxarifado padrao associado ao produto
                lErro = Comando_ExecutarLockado(alComando(16), "SELECT Almoxarifado FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ? ", iAlmoxPadraoProd, objItemOP.iFilialEmpresa, objProdutoKitProdutos.sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 55209

                lErro = Comando_BuscarPrimeiro(alComando(16))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 55210

                lErro = Comando_LockExclusive(alComando(16))
                If lErro <> AD_SQL_SUCESSO Then gError 55211

                'se achou o almoxarifado padrão do produto ==> cria o empenho neste almoxarifado, senao nao cria o empenho
                If lErro = SUCESSO And iAlmoxPadraoProd <> 0 Then

                    If objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then

                        objProdutoKitProdutos.dQuantidade = objProdutoKitProdutos.dQuantidade * ((objItemOP.dQuantidade * dFatorItemOP) / (tProdutoKitRaiz.dQuantidade * dFatorRaiz))

                    End If

                    lErro = CF("UM_Conversao_Trans", objProdutoKitProdutos.iClasseUM, objProdutoKitProdutos.sUnidadeMed, objProdutoKitProdutos.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 31564

                    Call CF("Arredonda_Insumos", objProdutoKitProdutos.sUnidadeMed, objProdutoKitProdutos.dQuantidade * dFator, dQuantidadeArr)

                    objEmpenho.dQuantidade = dQuantidadeArr
                    objEmpenho.dtData = dtDataEmissao
                    objEmpenho.iAlmoxarifado = iAlmoxPadraoProd
                    objEmpenho.iFilialEmpresa = objItemOP.iFilialEmpresa
                    objEmpenho.iItemOP = objItemOP.iItem
                    objEmpenho.lCodigo = lEmpenho
                    objEmpenho.lNumIntDocItemOP = objItemOP.lNumIntDoc
                    objEmpenho.sCodigoOP = objItemOP.sCodigo
                    objEmpenho.sProduto = objProdutoKitProdutos.sProduto
                    objEmpenho.sUnidadeMed = objProdutoKitProdutos.sSiglaUMEstoque

                    lEmpenho = lEmpenho + 1

                    lErro = CF("Empenho_Inserir_Batch", objEmpenho, alComando(11))
                    If lErro <> SUCESSO Then gError 31565

                End If

            End If

        Next
        '**********************************************************
        
        lErro = Comando_Unlock(alComando(8))
        If lErro <> SUCESSO Then gError 31568

    End If

    OrdemDeProducao_Grava2 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava2:

    OrdemDeProducao_Grava2 = gErr

    Select Case gErr

        Case 31553, 31554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 31556, 31557, 31558, 31559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 31564, 31565, 55248, 55249, 106399

        Case 55209, 55210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, iAlmoxPadraoProd, objItemOP.sProduto)

        Case 55211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOSFILIAL", gErr)
            
        Case 106400
        '??? Dar mensagem de Erro ?
        'Sim, senão sai sem dar nenhuma mensagem em kit sem insumos e ninguém fica sabendo o motivo
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_APENAS_COM_PRODUTORAIZ2", gErr, objProdutoKitProdutos.sProdutoRaiz, objProdutoKitProdutos.sVersao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152591)

    End Select

End Function

Private Function OrdemDeProducao_Grava5(ByVal lComando As Long, ByVal lComando1 As Long, lNumIntDocItem As Long, lEmpenho As Long) As Long
'atualiza numeração de lNumIntDocItem e lEmpenho

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer

On Error GoTo Erro_OrdemDeProducao_Grava5

    'Abre novos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 22943
    Next

    lErro = Comando_ExecutarPos(alComando(0), "UPDATE MATConfig SET Conteudo = ?", lComando, lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then Error 21954

    lErro = OrdemDeProducao_Grava51(lComando1, lEmpenho)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava5 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava5:

    OrdemDeProducao_Grava5 = Err

    Select Case Err

        Case ERRO_SEM_MENSAGEM

        Case 22943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", Err, "NUM_PROX_ITEM_OP")

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152592)

    End Select

    'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OrdemDeProducao_Grava3(ByVal lComando As Long, ByVal lComando1 As Long, lNumIntDocItem As Long, lEmpenho As Long) As Long
'obtem o proximo codigo para lNumIntDocItem e lEmpenho

Dim lErro As Long, sConteudo As String
Dim iIndice As Integer

On Error GoTo Erro_OrdemDeProducao_Grava3

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ?", 0, sConteudo, "NUM_PROX_ITEM_OP")
    If lErro <> AD_SQL_SUCESSO Then Error 22955

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 22956

    lErro = Comando_LockExclusive(lComando)

    If lErro <> AD_SQL_SUCESSO Then Error 22957

    lNumIntDocItem = CLng(sConteudo)

    lErro = OrdemDeProducao_Grava31(lComando1, lEmpenho)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    OrdemDeProducao_Grava3 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava3:

    OrdemDeProducao_Grava3 = Err

    Select Case Err

        Case ERRO_SEM_MENSAGEM
        
        Case 22955, 22956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", Err, "NUM_PROX_ITEM_OP")

        Case 22957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG", Err, "NUM_PROX_ITEM_OP")

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152593)

    End Select

    Exit Function

End Function

Public Function OrdemDeProducao_Grava4(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'atualiza os itens pré-existentes da OP

Dim alComando(1 To 5) As Long, iIndice As Integer
Dim tItemOP As typeItemOP, objItemOP As ClassItemOP
Dim iAchou As Integer, sCcl As String, lErro As Long
Dim sParam1 As String, sParam2 As String
Dim objItemOPGrade As ClassItemOP
Dim tItemOPGrade As typeItemOP
Dim iAchouGrade As Integer
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_OrdemDeProducao_Grava4

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 22971
    Next

    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
    
    '###########################################
    'INSERIDO POR WAGNER
    tItemOP.sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
    '###########################################


    'tenta selecionar item pre-existente da OP
    '################################################
    'ALTERADO POR WAGNER
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntItemOPGrade, Produto, NumIntDoc, Item, FilialPedido, CodPedido, Almoxarifado, Ccl, Quantidade, DataInicioProd, DataFimProd, Prioridade, Situacao, Destinacao, QuantidadeProd, Versao, NumIntEquipamento, NumIntEquipamento2, Tipo, NumIntDocPai, JustificativaBloqueio, ProduzLogo FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Codigo = ? ORDER BY Item", 0, _
    tItemOP.lNumIntItemOP, tItemOP.sProduto, tItemOP.lNumIntDoc, tItemOP.iItem, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.dQuantidadeProd, tItemOP.sVersao, tItemOP.lNumIntEquipamento, tItemOP.lNumIntEquipamento2, tItemOP.iTipo, tItemOP.lNumIntDocPai, tItemOP.sJustificativaBloqueio, tItemOP.iProduzLogo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 22972
    '################################################

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 22973

    Do While lErro <> AD_SQL_SEM_DADOS

        'se nao for item de grade
        If tItemOP.lNumIntItemOP = 0 Then

            iAchou = 0
    
            'localiza item na coleção
            For Each objItemOP In objOrdemDeProducao.colItens
            'Acrescentado o produto e versão porque na Alacer dois usuários criaram uma OP com o mesmo código
            'mesma quantidade de itens e produto diferentes. Gravou a 1a e ao gravar a 2o não alterou
            'o produto da OP mas mudou toda parte do MPR com operações
            'Como na tela não deixa excluir ou alterar o produto o usuário só pode acrescentar ou mexer na quantidade,
            'fora isso vai dar erro
                If objItemOP.iItem = tItemOP.iItem And tItemOP.sProduto = objItemOP.sProduto And tItemOP.sVersao = objItemOP.sVersao Then
                    iAchou = 1
                    objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
                    Exit For
                End If
            Next
    
            If iAchou = 0 Then gError 22975
    
            If tItemOP.sCcl <> objItemOP.sCcl Then
    
                'bloquear ccl
                sCcl = String(STRING_CCL, 0)
                lErro = Comando_ExecutarLockado(alComando(3), "SELECT Ccl FROM Ccl WHERE Ccl = ?", sCcl, objItemOP.sCcl)
                If lErro <> AD_SQL_SUCESSO Then gError 41338
    
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 41339
    
                If lErro = AD_SQL_SEM_DADOS Then gError 41340
    
                lErro = Comando_LockShared(alComando(3))
                If lErro <> AD_SQL_SUCESSO Then gError 41341
    
            End If
    
            lErro = CF("Verifica_Situacao", tItemOP.iSituacao, objItemOP.iSituacao)
            If lErro <> SUCESSO Then gError 41342
    
            'se houve alguma alteracao , atualiza o Item da OP
            '###########################################
            'ALTERADO POR WAGNER
            If objItemOP.iProduzLogo <> tItemOP.iProduzLogo Or objItemOP.sJustificativaBloqueio <> tItemOP.sJustificativaBloqueio Or objItemOP.iFilialPedido <> tItemOP.iFilialPedido Or objItemOP.lCodPedido <> tItemOP.lCodPedido Or objItemOP.dQuantidade <> tItemOP.dQuantidade Or tItemOP.iAlmoxarifado <> objItemOP.iAlmoxarifado Or tItemOP.dtDataInicioProd <> objItemOP.dtDataInicioProd Or tItemOP.dtDataFimProd <> objItemOP.dtDataFimProd Or tItemOP.iPrioridade <> objItemOP.iPrioridade Or tItemOP.iSituacao <> objItemOP.iSituacao Or tItemOP.iDestinacao <> objItemOP.iDestinacao Or tItemOP.lNumIntEquipamento <> objItemOP.lNumIntEquipamento _
                Or tItemOP.lNumIntEquipamento2 <> objItemOP.lNumIntEquipamento2 Or tItemOP.sCcl <> objItemOP.sCcl Then
                
                lErro = OrdemDeProducao_Grava41(objOrdemDeProducao, objItemOP, tItemOP)
                If lErro <> SUCESSO Then gError 126547
    
                With objItemOP
                    lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensOrdemProducao SET FilialPedido = ?, CodPedido = ?, Almoxarifado = ?, Beneficiamento = ?, Ccl = ?, Quantidade = ?, DataInicioProd = ?, DataFimProd = ?, Prioridade = ?, Situacao = ?, Destinacao = ?, NumIntEquipamento = ?, NumIntEquipamento2 = ?, NumIntOrigem=?, OrigemPedido=?, JustificativaBloqueio = ?, ProduzLogo = ?", alComando(1), _
                        .iFilialPedido, .lCodPedido, .iAlmoxarifado, .iBeneficiamento, .sCcl, .dQuantidade, .dtDataInicioProd, .dtDataFimProd, .iPrioridade, .iSituacao, .iDestinacao, .lNumIntEquipamento, .lNumIntEquipamento2, .lNumIntOrigem, .iOrigemPedido, .sJustificativaBloqueio, .iProduzLogo)
                End With
    
                If lErro <> AD_SQL_SUCESSO Then gError 22979
    
            End If
            '###########################################

        Else
            
            iAchouGrade = 0
            
            'se for um item de grade
            For Each objItemOP In objOrdemDeProducao.colItens
        
                Set objItemOPGrade = New ClassItemOP
                
                objItemOPGrade.dQuantidade = objItemOP.dQuantidade
                objItemOPGrade.dQuantidadeProd = objItemOP.dQuantidadeProd
                objItemOPGrade.dtDataFimProd = objItemOP.dtDataFimProd
                objItemOPGrade.dtDataInicioProd = objItemOP.dtDataInicioProd
                objItemOPGrade.iAlmoxarifado = objItemOP.iAlmoxarifado
                objItemOPGrade.iBeneficiamento = objItemOP.iBeneficiamento
                objItemOPGrade.iClasseUM = objItemOP.iClasseUM
                objItemOPGrade.iDestinacao = objItemOP.iDestinacao
                objItemOPGrade.iFilialEmpresa = objItemOP.iFilialEmpresa
                objItemOPGrade.iFilialPedido = objItemOP.iFilialPedido
                objItemOPGrade.iItem = objItemOP.iItem
                objItemOPGrade.iOrigemPedido = objItemOP.iOrigemPedido
                objItemOPGrade.iPossuiGrade = objItemOP.iPossuiGrade
                objItemOPGrade.iPrioridade = objItemOP.iPrioridade
                objItemOPGrade.iSituacao = objItemOP.iSituacao
                objItemOPGrade.lCodPedido = objItemOP.lCodPedido
                objItemOPGrade.lNumIntDoc = objItemOP.lNumIntDoc
                objItemOPGrade.lNumIntEquipamento = objItemOP.lNumIntEquipamento
                objItemOPGrade.lNumIntEquipamento2 = objItemOP.lNumIntEquipamento2
                objItemOPGrade.lNumIntOrigem = objItemOP.lNumIntOrigem
                objItemOPGrade.lNumItemOP = objItemOP.lNumItemOP
                objItemOPGrade.sAlmoxarifadoNomeRed = objItemOP.sAlmoxarifadoNomeRed
                objItemOPGrade.sCcl = objItemOP.sCcl
                objItemOPGrade.sCodigo = objItemOP.sCodigo
                objItemOPGrade.sDescricao = objItemOP.sDescricao
                objItemOPGrade.sProduto = objItemOP.sProduto
                objItemOPGrade.sProdutoPai = objItemOP.sProdutoPai
                objItemOPGrade.sSiglaUM = objItemOP.sSiglaUM
                objItemOPGrade.sSiglaUMEstoque = objItemOP.sSiglaUMEstoque
                objItemOPGrade.sVersao = objItemOP.sVersao
                objItemOPGrade.iTipo = objItemOP.iTipo
        
                'localiza item na coleção
                For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
                    If objItemRomaneioGrade.lNumIntDoc = tItemOP.lNumIntDoc Then
                        
                        iAchouGrade = 1
                        
                        objItemOPGrade.lNumIntDoc = tItemOP.lNumIntDoc
                        objItemOPGrade.iAlmoxarifado = objItemRomaneioGrade.iAlmoxarifado
                        objItemOPGrade.dQuantidade = objItemRomaneioGrade.dQuantidade
                        Exit For
                    End If
                Next
        
                If iAchouGrade = 0 Then gError 126550
        
                lErro = CF("Verifica_Situacao", tItemOP.iSituacao, objItemOPGrade.iSituacao)
                If lErro <> SUCESSO Then gError 126694
        
                'se houve alguma alteracao , atualiza o Item da OP
                If objItemOPGrade.iFilialPedido <> tItemOP.iFilialPedido Or objItemOPGrade.lCodPedido <> tItemOP.lCodPedido Or objItemOPGrade.dQuantidade <> tItemOP.dQuantidade Or tItemOP.iAlmoxarifado <> objItemOPGrade.iAlmoxarifado Or tItemOP.dtDataInicioProd <> objItemOPGrade.dtDataInicioProd Or tItemOP.dtDataFimProd <> objItemOPGrade.dtDataFimProd Or tItemOP.iPrioridade <> objItemOPGrade.iPrioridade Or tItemOP.iSituacao <> objItemOPGrade.iSituacao Or tItemOP.iDestinacao <> objItemOPGrade.iDestinacao _
                    Or tItemOP.lNumIntEquipamento <> objItemOPGrade.lNumIntEquipamento Or tItemOP.lNumIntEquipamento2 <> objItemOPGrade.lNumIntEquipamento2 Or tItemOP.sCcl <> objItemOPGrade.sCcl Then
                    
                    lErro = OrdemDeProducao_Grava41(objOrdemDeProducao, objItemOPGrade, tItemOP)
                    If lErro <> SUCESSO Then gError 126547
        
                    With objItemOPGrade
                        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensOrdemProducao SET FilialPedido = ?, CodPedido = ?, Almoxarifado = ?, Beneficiamento = ?, Ccl = ?, Quantidade = ?, DataInicioProd = ?, DataFimProd = ?, Prioridade = ?, Situacao = ?, Destinacao = ?, NumIntEquipamento = ?, NumIntEquipamento2 = ?, NumIntOrigem=?, OrigemPedido=?", alComando(1), _
                            .iFilialPedido, .lCodPedido, .iAlmoxarifado, .iBeneficiamento, .sCcl, .dQuantidade, .dtDataInicioProd, .dtDataFimProd, .iPrioridade, .iSituacao, .iDestinacao, .lNumIntEquipamento, .lNumIntEquipamento2, .lNumIntOrigem, .iOrigemPedido)
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 126551
        
                End If
        
            Next
        
        End If

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 22974

    Loop

    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)

    'tenta selecionar item pre-existente da OP
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Produto, NumIntDoc, Item, FilialPedido, CodPedido, Almoxarifado, Ccl, Quantidade, DataInicioProd, DataFimProd, Prioridade, Situacao, Destinacao, QuantidadeProd, Versao, NumIntEquipamento, NumIntEquipamento2 FROM ItensOPGrade WHERE FilialEmpresa = ? AND Codigo = ? ORDER BY Item", 0, _
    tItemOP.sProduto, tItemOP.lNumIntDoc, tItemOP.iItem, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.dQuantidadeProd, tItemOP.sVersao, tItemOP.lNumIntEquipamento, tItemOP.lNumIntEquipamento2, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126695

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126696

    Do While lErro <> AD_SQL_SEM_DADOS

        iAchou = 0

        'localiza item na coleção
        For Each objItemOP In objOrdemDeProducao.colItens
            If objItemOP.iItem = tItemOP.iItem Then
                iAchou = 1
                objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
                Exit For
            End If
        Next

        If iAchou = 0 Then gError 126697

        If tItemOP.sCcl <> objItemOP.sCcl Then

            'bloquear ccl
            sCcl = String(STRING_CCL, 0)
            lErro = Comando_ExecutarLockado(alComando(3), "SELECT Ccl FROM Ccl WHERE Ccl = ?", sCcl, objItemOP.sCcl)
            If lErro <> AD_SQL_SUCESSO Then gError 126698

            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126699

            If lErro = AD_SQL_SEM_DADOS Then gError 126700

            lErro = Comando_LockShared(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 126701

        End If

        lErro = CF("Verifica_Situacao", tItemOP.iSituacao, objItemOP.iSituacao)
        If lErro <> SUCESSO Then gError 126702

        'se houve alguma alteracao , atualiza o Item da OP
        If objItemOP.iFilialPedido <> tItemOP.iFilialPedido Or objItemOP.lCodPedido <> tItemOP.lCodPedido Or objItemOP.dQuantidade <> tItemOP.dQuantidade Or tItemOP.iAlmoxarifado <> objItemOP.iAlmoxarifado Or tItemOP.dtDataInicioProd <> objItemOP.dtDataInicioProd Or tItemOP.dtDataFimProd <> objItemOP.dtDataFimProd Or tItemOP.iPrioridade <> objItemOP.iPrioridade Or tItemOP.iSituacao <> objItemOP.iSituacao Or tItemOP.iDestinacao <> objItemOP.iDestinacao _
            Or tItemOP.lNumIntEquipamento <> objItemOP.lNumIntEquipamento Or tItemOP.lNumIntEquipamento2 <> objItemOP.lNumIntEquipamento2 Or tItemOP.sCcl <> objItemOP.sCcl Then
            
            With objItemOP
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensOPGrade SET FilialPedido = ?, CodPedido = ?, Almoxarifado = ?, Beneficiamento = ?, Ccl = ?, Quantidade = ?, DataInicioProd = ?, DataFimProd = ?, Prioridade = ?, Situacao = ?, Destinacao = ?, NumIntEquipamento = ?, NumIntEquipamento2 = ?, NumIntOrigem=?, OrigemPedido=?", alComando(1), _
                    .iFilialPedido, .lCodPedido, .iAlmoxarifado, .iBeneficiamento, .sCcl, .dQuantidade, .dtDataInicioProd, .dtDataFimProd, .iPrioridade, .iSituacao, .iDestinacao, .lNumIntEquipamento, .lNumIntEquipamento2, .lNumIntOrigem, .iOrigemPedido)
            End With

            If lErro <> AD_SQL_SUCESSO Then gError 126703

        End If

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126704

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava4 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava4:

    OrdemDeProducao_Grava4 = gErr

    Select Case gErr

        Case 22971
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 22972 To 22974
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)

        Case 22975, 126697
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_NAO_EXCLUIVEL", gErr, tItemOP.iItem, objOrdemDeProducao.sCodigo)

        Case 22977, 41429, 126547

        Case 22979
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSORDENSDEPRODUCAO", gErr, objItemOP.iItem, objItemOP.sCodigo, objItemOP.iFilialEmpresa)

        Case 41338, 41339, 126698, 126699
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)

        Case 41340, 126700
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemOP.iAlmoxarifado)

        Case 41341, 126701
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr)

        Case 41342, 126694, 126702
            If tItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL Then sParam1 = STRING_NORMAL
            If tItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then sParam1 = STRING_DESABILITADA
            If tItemOP.iSituacao = ITEMOP_SITUACAO_SACR Then sParam1 = STRING_SACRAMENTADA
            If tItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then sParam1 = STRING_BAIXADA
            If objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL Then sParam2 = STRING_NORMAL
            If objItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then sParam2 = STRING_DESABILITADA
            If objItemOP.iSituacao = ITEMOP_SITUACAO_SACR Then sParam2 = STRING_SACRAMENTADA
            If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then sParam2 = STRING_BAIXADA

            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SITUACAO", gErr, sParam1, sParam2)

        Case 126550
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOPGRADE_NAO_EXCLUIVEL", gErr, objItemOP.iItem, objOrdemDeProducao.sCodigo, objItemOPGrade.sProduto)

        Case 126551, 126703
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSOPGRADE", gErr, objItemOP.iItem, objItemOP.sCodigo, objItemOPGrade.sProduto)

        Case 126695, 126696, 126704
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPGRADE1", gErr, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152594)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OrdemDeProducao_Grava41(ByVal objOrdemDeProducao As ClassOrdemDeProducao, ByVal objItemOP As ClassItemOP, tItemOP As typeItemOP) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim objItemPV As New ClassItemPedido
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantOPAnterior As Double, dQuantOPNova As Double

On Error GoTo Erro_OrdemDeProducao_Grava41


'        'Quantidade (só pode mudar p/maior ou igual a produzida)
'        If objItemOp.dQuantidade < tItemOP.dQuantidadeProd Then gError 22976

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126545
    Next

    'se trocou FilialPedido ou CodPedido ou a quantidade
    If objItemOP.iFilialPedido <> tItemOP.iFilialPedido Or objItemOP.lCodPedido <> tItemOP.lCodPedido Or objItemOP.dQuantidade <> tItemOP.dQuantidade Then

        'se trocou o pedido
        If tItemOP.lCodPedido <> objItemOP.lCodPedido Then

            If tItemOP.lCodPedido <> 0 Then

                'chamar funcao que faz update tirando a qtde anterior do pedido anterior
                objItemPV.iFilialEmpresa = tItemOP.iFilialPedido
                objItemPV.lCodPedido = tItemOP.lCodPedido
                objItemPV.sProduto = tItemOP.sProduto
                objItemPV.iTipoOP = tItemOP.iTipo

                'altera a quantidade em OP do Pedido de Venda
                lErro = OrdemDeProducao_Grava6(alComando(1), alComando(2), objItemPV, -tItemOP.dQuantidade)
                If lErro <> SUCESSO And lErro <> 41346 Then gError 41601

                If lErro = 41346 Then gError 41599
                
            End If

            If objItemOP.lCodPedido <> 0 Then

                'chamar funcao que faz update incluindo qtde nova no pedido novo
                objItemPV.iFilialEmpresa = objItemOP.iFilialPedido
                objItemPV.lCodPedido = objItemOP.lCodPedido
                objItemPV.sProduto = objItemOP.sProduto
                objItemPV.iTipoOP = objItemOP.iTipo

                'altera a quantidade em OP do Pedido de Venda
                lErro = OrdemDeProducao_Grava6(alComando(3), alComando(4), objItemPV, objItemOP.dQuantidade)
                If lErro <> SUCESSO And lErro <> 41346 Then gError 41602

                If lErro = 41346 Then gError 41436

                objItemOP.lNumIntOrigem = objItemPV.lNumIntDoc
                objItemOP.iOrigemPedido = objItemPV.iOrigemPedido

            End If

        Else

            If objItemOP.lCodPedido <> 0 Then

                'se trocou a qtde
                If tItemOP.dQuantidade <> objItemOP.dQuantidade Then

                    objItemPV.iFilialEmpresa = objItemOP.iFilialPedido
                    objItemPV.lCodPedido = objItemOP.lCodPedido
                    objItemPV.sProduto = objItemOP.sProduto
                    objItemPV.iTipoOP = objItemOP.iTipo

                   'chamar funcao que faz update pela diferenca
                   lErro = OrdemDeProducao_Grava6(alComando(5), alComando(6), objItemPV, objItemOP.dQuantidade - tItemOP.dQuantidade)
                   If lErro <> SUCESSO And lErro <> 41346 Then gError 41603

                   If lErro = 41346 Then gError 41600

                    objItemOP.lNumIntOrigem = objItemPV.lNumIntDoc
                    objItemOP.iOrigemPedido = objItemPV.iOrigemPedido
                
                End If

            End If

        End If

    End If

    'se se tratar de um OP (nao uma ordem de corte)
    If tItemOP.iTipo = OP_TIPO_OP Then
    
        'se alterou a qtde a ser produzida
        If objItemOP.dQuantidade <> tItemOP.dQuantidade And objItemOP.iPossuiGrade = DESMARCADO And objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL And tItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL Then
    
            If tItemOP.dQuantidade > tItemOP.dQuantidadeProd Then
                dQuantOPAnterior = Arredonda_Estoque(tItemOP.dQuantidade - tItemOP.dQuantidadeProd)
            Else
                dQuantOPAnterior = 0
            End If
            
            If objItemOP.dQuantidade > tItemOP.dQuantidadeProd Then
                dQuantOPNova = Arredonda_Estoque(objItemOP.dQuantidade - tItemOP.dQuantidadeProd)
            Else
                dQuantOPNova = 0
            End If
            
            If dQuantOPNova <> dQuantOPAnterior Then
            
                'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
        
                objEstoqueProduto.sProduto = tItemOP.sProduto
                objEstoqueProduto.iAlmoxarifado = tItemOP.iAlmoxarifado
        
                lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, dQuantOPNova - dQuantOPAnterior)
                If lErro <> SUCESSO Then gError 41551
    
            End If
    
        End If
    
        'se alterou a situacao para baixado ==> diminuir a quantidade em ordem de producao da quantidade que ainda faltava ser produzido
        If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA And objItemOP.iSituacao <> tItemOP.iSituacao And objItemOP.iPossuiGrade = DESMARCADO Then
    
            'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
            objEstoqueProduto.sProduto = tItemOP.sProduto
            objEstoqueProduto.iAlmoxarifado = tItemOP.iAlmoxarifado
    
            If tItemOP.dQuantidade > tItemOP.dQuantidadeProd Then
    
                lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, -(tItemOP.dQuantidade - tItemOP.dQuantidadeProd))
                If lErro <> SUCESSO Then gError 55254
    
            End If
    
            objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
            objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
    
            'exclui empenhos associados ao itemOP
            lErro = CF("Empenhos_Excluir_ItemOP1", objItemOP, alComando(7), alComando(8))
            If lErro <> SUCESSO Then gError 55259
    
        End If
    
        ' Se a situacao era 'baixado' e foi alterada para 'normal'
        If objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL And tItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA And objItemOP.iPossuiGrade = DESMARCADO Then
    
            If objItemOP.dQuantidade > tItemOP.dQuantidadeProd Then
                dQuantOPNova = Arredonda_Estoque(objItemOP.dQuantidade - tItemOP.dQuantidadeProd)
            Else
                dQuantOPNova = 0
            End If
            
            If dQuantOPNova <> 0 Then
            
                'altera a qtde em OP para o produto-almoxarifado na tabela EstoqueProduto
                objEstoqueProduto.sProduto = tItemOP.sProduto
                objEstoqueProduto.iAlmoxarifado = tItemOP.iAlmoxarifado
        
                If tItemOP.dQuantidade > tItemOP.dQuantidadeProd Then
        
                    lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, dQuantOPNova)
                    If lErro <> SUCESSO Then gError 126546
        
                End If
        
            End If
        
        End If

    End If
    
    lErro = OrdemDeProducao_Grava42(objOrdemDeProducao, objItemOP, tItemOP)
    If lErro <> SUCESSO Then gError 126546

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava41 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava41:

    OrdemDeProducao_Grava41 = gErr

    Select Case gErr

'        Case 22976
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_QTDE_MENOR_PROD", gErr)

        Case 41436, 41600
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemOP.lCodPedido, objItemOP.sProduto)

        Case 41551, 41601, 41602, 41603, 55254, 55259, 126546
        
        Case 41599
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemOP.lCodPedido, objItemOP.sProduto)

        Case 126545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152595)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrdemDeProducao_Grava6(ByVal lComando As Long, ByVal lComando1 As Long, _
                 ByVal objItemPV As ClassItemPedido, ByVal dQuantidade As Double) As Long
'Atualiza QuantOP em ItensPedidoDeVenda ou ItensPedidoDeVendaGrade

Dim lErro As Long
Dim sComando_SQL As String
Dim lNumIntDoc As Long
Dim lNumIntDoc1 As Long
Dim dQuantOP As Double
Dim dQuant As Double
Dim dQuantCanc As Double

On Error GoTo Erro_OrdemDeProducao_Grava6
    
    'Seleciona QuantOp e  NumIntDoc
    sComando_SQL = "SELECT NumIntDoc, QuantOP, Quantidade, QuantCancelada FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ? AND Produto = ?"

    lErro = Comando_ExecutarPos(lComando, sComando_SQL, 0, lNumIntDoc, dQuantOP, dQuant, dQuantCanc, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 41344

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 41345

    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Lock no registro
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 41347
    
        If objItemPV.iTipoOP = OP_TIPO_OP Then
        
            'Se está produzindo mais que o necessário é porque o restante vai para o estoque, não tem que marcar a quantop no pedido
            If ((dQuantCanc + dQuantOP + dQuantidade) - dQuant) > QTDE_ESTOQUE_DELTA Then
                'A quantidade a ser acertada é o que falta a ser produzido
                'para atender o pedido completamente
                dQuantidade = dQuant - dQuantCanc - dQuantOP
            End If
    
            'Atualiza a quantOP
            sComando_SQL = "UPDATE ItensPedidoDeVenda SET QuantOP = QuantOP + ?"
        
            lErro = Comando_ExecutarPos(lComando1, sComando_SQL, lComando, dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 41348
        
        End If
        
        objItemPV.lNumIntDoc = lNumIntDoc
        objItemPV.iOrigemPedido = ORIGEM_PEDIDO_ITEM_PV
        
    Else
    
        sComando_SQL = "SELECT NumIntDoc, QuantOP, Quantidade, QuantCancelada FROM ItensPedidoDeVendaBaixados WHERE FilialEmpresa = ? AND CodPedido = ? AND Produto = ?"

        lErro = Comando_ExecutarPos(lComando, sComando_SQL, 0, lNumIntDoc, dQuantOP, dQuant, dQuantCanc, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 95171
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95172
        
        If lErro <> AD_SQL_SEM_DADOS Then
        
            'Lock no registro
            lErro = Comando_LockExclusive(lComando)
            If lErro <> AD_SQL_SUCESSO Then gError 95174
        
            If objItemPV.iTipoOP = OP_TIPO_OP Then
            
                'Atualiza a quantOP
                sComando_SQL = "UPDATE ItensPedidoDeVendaBaixados SET QuantOP = QuantOP + ?"
            
                lErro = Comando_ExecutarPos(lComando1, sComando_SQL, lComando, dQuantidade)
                If lErro <> AD_SQL_SUCESSO Then gError 95175

            End If

            objItemPV.lNumIntDoc = lNumIntDoc
            objItemPV.iOrigemPedido = ORIGEM_PEDIDO_ITEM_PV

        Else

            'Atualiza QuantOp em ItensPVGrade
            lErro = ItensPVGrade_Atualiza_QuantOP(objItemPV, dQuantidade)
            If lErro <> SUCESSO And lErro <> 126530 And lErro <> 126531 Then gError 126529
            
            If lErro <> SUCESSO Then gError 41346
        
        End If
        
    End If

    OrdemDeProducao_Grava6 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava6:

    OrdemDeProducao_Grava6 = gErr

    Select Case gErr

        Case 41344, 41345, 95171, 95172
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 41346, 126529

        Case 41347, 95174
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDODEVENDA", gErr)

        Case 41348, 95175
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152596)

    End Select

    Exit Function

End Function

Public Function ItensPVGrade_Atualiza_QuantOP(ByVal objItemPV As ClassItemPedido, ByVal dQuantidade As Double) As Long
'Atualiza QuantOp em ItensPVGrade

Dim lErro As Long
Dim sComando_SQL As String
Dim lNumIntDoc As Long
Dim lNumIntDoc1 As Long
Dim dQuantOP As Double
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim sProdutoPai As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sTabela As String
Dim lComando As Long

On Error GoTo Erro_ItensPVGrade_Atualiza_QuantOP
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126522
    Next
    
'    sProdutoMascarado = String(STRING_PRODUTO, 0)
'
'    'coloca o Produto no formato que é exibida na tela
'    lErro = Mascara_RetornaProdutoTela(objItemPV.sProduto, sProdutoMascarado)
'    If lErro <> SUCESSO Then gError 126534

    sProdutoPai = String(STRING_PRODUTO, 0)

    'retorna a Produto "pai" da Produto em questão, se houver
    lErro = Mascara_RetornaProdutoPai(objItemPV.sProduto, sProdutoPai)
    If lErro <> SUCESSO Then gError 126535
    
    'Critica o formato do Produto
    lErro = CF("Produto_Formata", sProdutoPai, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 126536
    
    objItemPV.sProdutoPai = sProdutoFormatado
    
    sTabela = "ItensPedidoDeVenda"
    lComando = alComando(1)
    
    'Seleciona QuantOp e  NumIntDoc
    sComando_SQL = "SELECT NumIntDoc FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ? AND Produto = ?"
    
    lErro = Comando_ExecutarPos(alComando(1), sComando_SQL, 0, lNumIntDoc1, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProdutoPai)
    If lErro <> AD_SQL_SUCESSO Then gError 126523
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126524
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        sTabela = "ItensPedidoDeVendaBaixados"
        lComando = alComando(2)
    
        sComando_SQL = "SELECT NumIntDoc FROM ItensPedidoDeVendaBaixados WHERE FilialEmpresa = ? AND CodPedido = ? AND Produto = ?"

        lErro = Comando_ExecutarPos(alComando(2), sComando_SQL, 0, lNumIntDoc1, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProdutoPai)
        If lErro <> AD_SQL_SUCESSO Then gError 126525
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126526
        
        If lErro = AD_SQL_SEM_DADOS Then gError 126530
    
    End If
    
    'Seleciona QuantOp e  NumIntDoc
    sComando_SQL = "SELECT NumIntDoc, QuantOP FROM ItensPedidoDeVendaGrade WHERE NumIntItemPV = ? AND Produto = ?"

    lErro = Comando_ExecutarPos(alComando(3), sComando_SQL, 0, lNumIntDoc, dQuantOP, lNumIntDoc1, objItemPV.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 126527

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126528

    If lErro = AD_SQL_SEM_DADOS Then gError 126531

    'Lock no registro
    lErro = Comando_LockExclusive(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 126532

    If objItemPV.iTipoOP = OP_TIPO_OP Then

        'Atualiza a quantOP
        sComando_SQL = "UPDATE ItensPedidoDeVendaGrade SET QuantOP = QuantOP + ?"
    
        lErro = Comando_ExecutarPos(alComando(4), sComando_SQL, alComando(3), dQuantidade)
        If lErro <> AD_SQL_SUCESSO Then gError 126533
        
        'Atualiza a quantOP
        sComando_SQL = "UPDATE " & sTabela & " SET QuantOP = QuantOP + ?"
    
        lErro = Comando_ExecutarPos(alComando(5), sComando_SQL, lComando, dQuantidade)
        If lErro <> AD_SQL_SUCESSO Then gError 126533
        
    End If
        
    objItemPV.lNumIntDoc = lNumIntDoc
    objItemPV.iOrigemPedido = ORIGEM_PEDIDO_ITEM_PV_GRADE

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    ItensPVGrade_Atualiza_QuantOP = SUCESSO

    Exit Function

Erro_ItensPVGrade_Atualiza_QuantOP:

    ItensPVGrade_Atualiza_QuantOP = gErr

    Select Case gErr

        Case 126522
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126523, 126524
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 126525, 126526
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDABAIXADOS1", gErr)

        Case 126527, 126528
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVGRADE1", gErr, objItemPV.lCodPedido, objItemPV.sProdutoPai, objItemPV.sProduto)

        Case 126530, 126531, 126536

        Case 126532
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMPVGRADE", gErr, objItemPV.lCodPedido, objItemPV.sProdutoPai, objItemPV.sProduto)

        Case 126533
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITEMPVGRADE", gErr, objItemPV.lCodPedido, objItemPV.sProdutoPai, objItemPV.sProduto)

        Case 126534
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemPV.sProduto)

        Case 126535
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", gErr, sProdutoMascarado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152597)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function OrdemProducao_Grava_OPSubprodutos(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Gera uma Ordem de Produção com os insumos produzíveis dos produtos da OP passada

Dim lErro As Long
Dim iItem As Integer
Dim iIndice As Integer
Dim objKit As New ClassKit
Dim dSaldoAtual As Double
Dim iContAnterior As Integer
Dim objItemOP As ClassItemOP
Dim objProduto As New ClassProduto
Dim objOrdemProd As New ClassOrdemDeProducao

On Error GoTo Erro_OrdemProducao_Grava_OPSubprodutos

    'Para cada Item da Ordem de Produção passada
    For iIndice = 1 To objOrdemDeProducao.colItens.Count
        
        'Busca o produto no Banco de Dados
        objProduto.sCodigo = objOrdemDeProducao.colItens(iIndice).sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 62565
        
        If lErro <> SUCESSO Then gError 62566
        
        'Se o produto e produzivel
        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
            
            objKit.sProdutoRaiz = objProduto.sCodigo
            objKit.sVersao = objOrdemDeProducao.colItens(iIndice).sVersao
            
            'verifica se ele é um Kit
            lErro = CF("Kit_Le", objKit)
            If lErro <> SUCESSO And lErro <> 21826 Then gError 62567
            
            If lErro = SUCESSO Then
                
                'Inclui os subprodutos do Kit como itens em uma OP.
                lErro = ProdutoOP_Extrai_KitOP(objKit, objOrdemProd, objOrdemDeProducao.colItens(iIndice).dQuantidade, objOrdemDeProducao.colItens(iIndice).dtDataInicioProd)
                If lErro <> SUCESSO Then gError 62568
                
            End If
        End If
    Next
    
    'Para cada ItemOP na Ordem a ser criada
    For iIndice = objOrdemProd.colItens.Count To 1 Step -1
        
        Set objItemOP = objOrdemProd.colItens(iIndice)
        
        'Busca no BD o saldo atual no Estoque do produto
        lErro = CF("SaldoAtualProdProduzido", objItemOP.sProduto, dSaldoAtual)
        If lErro <> SUCESSO Then gError 62569
        
        If dSaldoAtual >= objItemOP.dQuantidade Then
            objOrdemProd.colItens.Remove iIndice
        Else
            objItemOP.dQuantidade = objItemOP.dQuantidade - dSaldoAtual
        End If

    Next
    'Atualiza os indices dos itens
    iContAnterior = 0
    
    For iIndice = 1 To objOrdemProd.colItens.Count
        
        Set objItemOP = objOrdemProd.colItens(iIndice)
        objItemOP.iItem = iIndice
        objItemOP.sCodigo = objOrdemDeProducao.sCodOPGerar
        If objItemOP.dtDataInicioProd < gdtDataHoje Then iContAnterior = iContAnterior + 1
    Next
    'Se há algum item na OP gerada
    If objOrdemProd.colItens.Count > 0 Then
        
        If iContAnterior > 0 Then
            Call Chama_Tela_Modal("ValidaGeracaoOP", objOrdemProd)
            If giRetornoTela = vbCancel Then gError 62570
        End If
                
        'Preenche os demais dados da OP
        objOrdemProd.dtDataEmissao = gdtDataHoje
        objOrdemProd.iFilialEmpresa = giFilialEmpresa
        objOrdemProd.iGeraOP = 0
        objOrdemProd.iGeraReqCompra = 0
        objOrdemProd.iNumItens = objOrdemProd.colItens.Count
        objOrdemProd.sCodigo = objOrdemDeProducao.sCodOPGerar
        objOrdemProd.sOPGeradora = objOrdemDeProducao.sCodigo
        
        'Grava a nova OP gerada
        lErro = CF("OrdemDeProducao_Grava_EmTrans", objOrdemProd)
        If lErro <> SUCESSO Then gError 62571
    End If

    OrdemProducao_Grava_OPSubprodutos = SUCESSO
    
    Exit Function
    
Erro_OrdemProducao_Grava_OPSubprodutos:

    OrdemProducao_Grava_OPSubprodutos = gErr
    
    Select Case gErr
    
        Case 62565, 62567, 62568, 62569, 62570, 62571
        
        Case 62566
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152598)
            
    End Select
    
    Exit Function

End Function

Function SaldoAtualProdComprado(ByVal sCodProduto As String, dSaldoAtual As Double, Optional dQuantReservada As Double, Optional dQuantEmReq As Double, Optional dQuantEmPed As Double, Optional dSaldoDisponivel As Double, Optional dQuantEmpenhada As Double) As Long
'Calcula o Saldo Atual para o produto comprado passado por parâmetro.

Dim lErro As Long
Dim lComando As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProduto As New ClassProduto

On Error GoTo Erro_SaldoAtualProdComprado
    
    'Lê a quantidade do produto em Requisição
    lErro = CF("Produto_Le_QuantRequisitada", sCodProduto, dQuantEmReq)
    If lErro <> SUCESSO Then gError 84584
    
    objEstoqueProduto.sProduto = sCodProduto
    'Lê o estoque do produto para a filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 52979 Then gError 84585
    
    'Calcula a quantidade do produto em Transferência
    lErro = CF("Produto_Le_QuantEmPedidoCompra", sCodProduto, dQuantEmPed)
    If lErro <> SUCESSO Then gError 84586
             
    'Calcula o Saldo Atual do produto
    With objEstoqueProduto
        dSaldoDisponivel = .dQuantDisponivel + .dQuantRecIndl
        dQuantEmpenhada = .dQuantEmpenhada
        dQuantReservada = .dQuantReservada
        dSaldoAtual = objEstoqueProduto.dQuantReservada + dSaldoDisponivel + dQuantEmPed + dQuantEmReq - .dQuantEmpenhada
    End With
       
    SaldoAtualProdComprado = SUCESSO
    
    Exit Function
    
Erro_SaldoAtualProdComprado:

    SaldoAtualProdComprado = gErr
    
    Select Case gErr
    
        Case 84584, 84585, 84586
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152599)
            
    End Select
    
    Exit Function
    
End Function


Function Produto_Le_QuantRequisitada(ByVal sProduto As String, dQuantEmRequisicao As Double) As Long
'Lê a quantidade do produto passado que estão em Requisição de Compras
'e que não foram pedidas ainda

Dim dQuantidade As Double
Dim lComando As Long
Dim lErro As Long
Dim dFator As Double
Dim sUM As String
Dim iClasseUM As Integer
Dim sUMEstoque As String

On Error GoTo Erro_Produto_Le_QuantRequisitada

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62575
    
    sUM = String(STRING_UM_SIGLA, 0)
    sUMEstoque = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT SUM(ItensReqCompra.Quantidade - ItensReqCompra.QuantCancelada - ItensReqCompra.QuantRecebida) , ItensReqCompra.UM, MIN(Produtos.ClasseUM), MIN(Produtos.SiglaUMEstoque)  FROM ItensReqCompra, RequisicaoCompra, Produtos WHERE ItensReqCompra.Produto = Produtos.Codigo AND RequisicaoCompra.NumIntDoc = ItensReqCompra.ReqCompra AND FilialEmpresa = ? AND ItensReqCompra.Produto = ? AND  RequisicaoCompra.NumIntDoc NOT IN (SELECT ItemRC FROM ItemRCItemPC) GROUP BY ItensReqCompra.Produto, ItensReqCompra.UM ", dQuantidade, sUM, iClasseUM, sUMEstoque, giFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 62576
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62577
    
    dQuantEmRequisicao = 0
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = CF("UM_Conversao_Trans", iClasseUM, sUM, sUMEstoque, dFator)
        If lErro <> SUCESSO Then Error 62578
        
        dQuantidade = dQuantidade * dFator
        
        dQuantEmRequisicao = dQuantEmRequisicao + dQuantidade

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 62579
        
    Loop
    
    Call Comando_Fechar(lComando)

    Produto_Le_QuantRequisitada = SUCESSO
    
    Exit Function

Erro_Produto_Le_QuantRequisitada:

    Produto_Le_QuantRequisitada = Err
    
    Select Case Err
    
        Case 62575
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62576, 62577, 62579
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err)
            
        Case 62578
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152600)
        
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_QuantEmPedidoCompra(ByVal sProduto As String, dQuantEmPedido As Double) As Long
'Lê a quantidade do produto passado que foi pedida

Dim dQuantidade As Double
Dim lComando As Long
Dim lErro As Long
Dim dFator As Double
Dim sUM As String
Dim iClasseUM As Integer
Dim sUMEstoque As String

On Error GoTo Erro_Produto_Le_QuantEmPedidoCompra

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62580
    
    sUM = String(STRING_UM_SIGLA, 0)
    sUMEstoque = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT SUM(ItensPedCompra.Quantidade - ItensPedCompra.QuantRecebida - ItensPedCompra.QuantRecebimento), ItensPedCompra.UM, MIN(Produtos.ClasseUM), MIN(Produtos.SiglaUMEstoque) FROM ItensPedCompra, Produtos, PedidoCompra WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.Produto = Produtos.Codigo AND PedidoCompra.FilialEmpresa = ? AND ItensPedCompra.Produto = ? AND (ItensPedCompra.Quantidade - ItensPedCompra.QuantRecebida - ItensPedCompra.QuantRecebimento) > 0 GROUP BY ItensPedCompra.Produto, ItensPedCompra.UM", dQuantidade, sUM, iClasseUM, sUMEstoque, giFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 62581
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62582
    
    dQuantEmPedido = 0
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = CF("UM_Conversao_Trans", iClasseUM, sUM, sUMEstoque, dFator)
        If lErro <> SUCESSO Then Error 62583
        
        dQuantidade = dQuantidade * dFator
        
        dQuantEmPedido = dQuantEmPedido + dQuantidade

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 62584
        
    Loop
    
    Call Comando_Fechar(lComando)

    Produto_Le_QuantEmPedidoCompra = SUCESSO
    
    Exit Function

Erro_Produto_Le_QuantEmPedidoCompra:

    Produto_Le_QuantEmPedidoCompra = Err
    
    Select Case Err
    
        Case 62580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62581, 62582, 62584
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err)
            
        Case 62583
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152601)
        
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SaldoAtualProdProduzido(ByVal sCodProduto As String, dSaldoAtual As Double)
'Calcula o saldo atual do produto produzido passado.

Dim lErro As Long
Dim lComando As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantOrdenada As Double
Dim dQuantEmTransf As Double

On Error GoTo Erro_SaldoAtualProdProduzido
    
''    'Lê a quantidade do produto em OPs
''    lErro = Produto_Le_QuantOP(sCodProduto, dQuantOrdenada)
''    If lErro <> SUCESSO Then Error 62585
''
''    'Lê a quantidade do produto em transferência
''    lErro = CF("Produto_Calcula_QuantTransferencia", sCodProduto, dQuantEmTransf)
''    If lErro <> SUCESSO Then Error 62586
       
    objEstoqueProduto.sProduto = sCodProduto
    'Lê as quantidades do produtos nos almoxs da filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, giFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 52979 Then Error 62587
    
    'Calcula o saldo atual
    With objEstoqueProduto
        dSaldoAtual = .dQuantDisponivel + .dQuantRecIndl + dQuantOrdenada + dQuantEmTransf - .dQuantEmpenhada
    End With

    dSaldoAtual = dSaldoAtual - dQuantOrdenada
    
    SaldoAtualProdProduzido = SUCESSO
    
    Exit Function
    
Erro_SaldoAtualProdProduzido:

    SaldoAtualProdProduzido = Err
    
    Select Case Err
    
        Case 62585, 62586, 62587
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152602)
            
    End Select
    
    Exit Function
    
End Function

Function ProdutoOP_Extrai_KitOP(ByVal objKit As ClassKit, ByVal objOrdemProd As ClassOrdemDeProducao, ByVal dQuantProduzir As Double, ByVal dtDataInicProd As Date) As Long
'Inclui os produtos e subprodutos do Kit passado em uma OP.

Dim objProdutoKit As ClassProdutoKit
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iPosicao As Integer
Dim bAchou As Boolean
Dim dQuantProd As Double
Dim dFator As Double
Dim dSaldoAtual As Double
Dim objItemOP As New ClassItemOP
Dim iAlmoxarifadoPadrao As Integer
Dim dtDataProd As Date
Dim iIntervalo As Integer
Dim objKitAux As ClassKit

On Error GoTo Erro_ProdutoOP_Extrai_KitOP

    'Lê os componentes do Kit passado
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 62588

    'Para cada componente encontrado
    For Each objProdutoKit In objKit.colComponentes

        If objProdutoKit.sProduto <> objKit.sProdutoRaiz Then
            objProduto.sCodigo = objProdutoKit.sProduto
            'Busca o Produto no BD
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 62589
            If lErro <> SUCESSO Then gError 62590
    
            'Se o produto for produzivel
            If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
    
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 62591
    
                'Verifica se o produto já é um item da OP
                bAchou = CF("BuscaProduto_ColItensOP", objProduto.sCodigo, objOrdemProd.colItens, iPosicao)
    
                'Calcula a quantidade de produção para esse produto
                dQuantProd = ((objProdutoKit.dQuantidade * dFator) * dQuantProduzir)
                
                lErro = CF("IntervaloProducao_Le", iIntervalo)
                If lErro <> SUCESSO Then gError 62592
                
                dtDataProd = dtDataInicProd - objProduto.iTempoProducao - iIntervalo
                
                'Se o produto já é um ItemOP
                If bAchou Then
                    'Acrescenta a quant a produzir
                    objOrdemProd.colItens(iPosicao).dQuantidade = objOrdemProd.colItens(iPosicao).dQuantidade + dQuantProd
                    If dtDataProd < objOrdemProd.colItens(iPosicao).dtDataInicioProd Then objOrdemProd.colItens(iPosicao).dtDataInicioProd = dtDataProd
                    
                'Senão
                Else
                    
                    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifadoPadrao)
                    If lErro <> SUCESSO And lErro <> 23796 Then gError 62593
                    If iAlmoxarifadoPadrao = 0 Then gError 62594
                    
                    'Cria um ItemOP novo com o produto
                    Set objItemOP = New ClassItemOP
    
                    objItemOP.dQuantidade = dQuantProd
                    objItemOP.dtDataFimProd = dtDataProd + objProduto.iTempoProducao
                    objItemOP.dtDataInicioProd = dtDataProd
                    objItemOP.iAlmoxarifado = iAlmoxarifadoPadrao
                    objItemOP.iDestinacao = ITEMOP_DESTINACAO_ESTOQUE
                    objItemOP.iFilialEmpresa = giFilialEmpresa
                    objItemOP.sProduto = objProduto.sCodigo
                    objItemOP.sSiglaUM = objProduto.sSiglaUMEstoque
                    objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
                    objItemOP.sVersao = objKit.sVersao
                    
                    objOrdemProd.colItens.Add objItemOP
    
                End If
    
                Set objKitAux = New ClassKit
                
                objKitAux.sProdutoRaiz = objProduto.sCodigo
                objKitAux.sVersao = objKit.sVersao
                
                'Verifica se o produto é um Kit
                lErro = CF("Kit_Le", objKitAux)
                If lErro <> AD_SQL_SUCESSO And lErro <> 21826 Then gError 62595
                
                'Se for
                If lErro = SUCESSO Then
                    
                    'Inclui os seus subprodutos na OP
                    lErro = ProdutoOP_Extrai_KitOP(objKitAux, objOrdemProd, dQuantProd, objItemOP.dtDataInicioProd)
                    If lErro <> SUCESSO Then gError 62596
                
                End If
                
            End If
            
        End If
        
    Next

    ProdutoOP_Extrai_KitOP = SUCESSO

    Exit Function

Erro_ProdutoOP_Extrai_KitOP:

    ProdutoOP_Extrai_KitOP = gErr

    Select Case gErr

        Case 62588, 62589, 62591, 62592, 62593, 62595, 62596

        Case 62590
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 62594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152603)

    End Select
    
End Function

Function BuscaProduto_ColItensOP(ByVal sProduto As String, ByVal colItensOP As Collection, iPosicao As Integer) As Boolean
'Busca o produto passado na coleção de Itens de Ordem de produção passada
'Se encontrar retorna True e a posição do produto na coleção, se nào encontrar retorna zero.

Dim bAchou As Boolean
Dim objItemOP As ClassItemOP

    iPosicao = 0
    bAchou = False
    
    'Para cada produto ca coleção
    For Each objItemOP In colItensOP
        
        'Atualiza a posição do produto
        iPosicao = iPosicao + 1
        
        'Se o produto procurado foi encontrado
        If objItemOP.sProduto = sProduto Then
            'Guarda que achou e sai do For
            bAchou = True
            Exit For
        End If
    Next
    
    'Retorna o resultado da busca para a rotina chamadora
    BuscaProduto_ColItensOP = bAchou
    
    Exit Function

End Function

Function Produto_Le_QuantOP(ByVal sProduto As String, dQuantOrdenada As Double) As Long
'Lê as quantidades das OPs que ainda não foram produzidas
'[Quantidade - QuantidadeProd]

Dim lComando As Long
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Produto_Le_QuantOP
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62597
    
    'Lê os Itens OP
    lErro = Comando_Executar(lComando, "SELECT SUM(Quantidade - QuantidadeProd)  FROM ItensOrdemProducao WHERE Produto = ? AND FilialEmpresa = ? AND Destinacao = ?", dQuantidade, sProduto, giFilialEmpresa, ITEMOP_DESTINACAO_ESTOQUE)
    If lErro <> AD_SQL_SUCESSO Then Error 62598
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 62599
    
    'para cada ItemOP encontrada
    Do While lErro = AD_SQL_SUCESSO
        'Acumula a quantidade
        dQuantOrdenada = dQuantOrdenada + dQuantidade
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62600
        
    Loop
    
    Call Comando_Fechar(lComando)

    Produto_Le_QuantOP = SUCESSO
    
    Exit Function

Erro_Produto_Le_QuantOP:

    Produto_Le_QuantOP = Err
    
    Select Case Err
    
        Case 62597
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62598, 62599, 62600
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152604)
            
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'??? será que falta implementar ??? Jones 16/01/2002
Function Produto_Calcula_QuantTransferencia(ByVal sCodProduto As String, dQuantTransf As Double) As Long
'Calcula a quantidade em Transferencia do produto passado

''''Dim lComando As Long
''''Dim lErro As Long
''''Dim dQuantidade As Double
''''Dim sUM As String
''''Dim sUMEst As String
''''Dim iClasseUM As Integer
''''
''''On Error GoTo Erro_Produto_Calcula_QuantTransferencia
''''
''''    'Abre comando
''''    lComando = Comando_Abrir()
''''    If lComando = 0 Then Error 62601
''''
''''    sUM = String(STRING_UM_SIGLA, 0)
''''    sUMEst = String(STRING_UM_SIGLA, 0)
''''
''''    'Le a quantidade dos itens
''''    lErro = Comando_Executar(lComando, "SELECT ItensNFiscal.UnidadeMed, ItensNFiscal.Quantidade, Produtos.SiglaUMEstoque, Produtos.ClasseUM FROM Produtos,ItensNFiscal, NFiscal, FiliaisClientes, FiliaisEmpresa WHERE ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND NFiscal.NaturezaOp IN (521, 522, 524, 592, 621, 622, 624, 692) AND FiliaisClientes.CodCliente = NFiscal.Cliente AND FiliaisClientes.CodFilial = NFiscal.FilialEntrega AND FiliaisClientes.CGC = FiliaisEmpresa.CGC AND ItensNFiscal.Produto = ? AND NFiscal.FilialEmpresa <> ? AND FiliaisEmpresa.FilialEmpresa = ?", sUM, dQuantidade, sUMEst, iClasseUM, sCodProduto, giFilialEmpresa, giFilialEmpresa)
''''    If lErro <> AD_SQL_SUCESSO Then Error 62602
''''
''''    lErro = Comando_BuscarPrimeiro(lComando)
''''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62603
''''
''''    'Enquanto encontrar
''''    Do While lErro <> AD_SQL_SEM_DADOS
''''
''''        '???? CONVERTER UM
''''
''''        dQuantTransf = dQuantTransf + dQuantidade
''''
''''        'Busca próximo
''''        lErro = Comando_BuscarProximo(lComando)
''''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62604
''''
''''    Loop
''''
''''    Call Comando_Fechar(lComando)
''''
''''    Produto_Calcula_QuantTransferencia = SUCESSO
''''
''''    Exit Function
''''
''''Erro_Produto_Calcula_QuantTransferencia:
''''
''''    Produto_Calcula_QuantTransferencia = Err
''''
''''    Select Case Err
''''
''''        Case 62601
''''            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''''
''''        Case 62602, 62603, 62604
''''            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152605)
''''
''''    End Select
''''
''''    Call Comando_Fechar(lComando)
''''
''''    Exit Function
''''
End Function

Function OrdemProducao_Grava_ReqSubprodutos(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim iItem As Integer
Dim iIndice As Integer
Dim objKit As New ClassKit
Dim dSaldoAtual As Double
Dim objProduto As New ClassProduto
Dim objItemRC As ClassItemReqCompras
Dim objReqCompras As New ClassRequisicaoCompras

On Error GoTo Erro_OrdemProducao_Grava_ReqSubprodutos

    'Para cada Item da Ordem de Produção passada
    For iIndice = 1 To objOrdemDeProducao.colItens.Count
        
        'Busca o produto no Banco de Dados
        objProduto.sCodigo = objOrdemDeProducao.colItens(iIndice).sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62605
        If lErro <> SUCESSO Then gError 62606
        
        'Se o produto e produzivel
        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then

            objKit.sProdutoRaiz = objProduto.sCodigo
            objKit.sVersao = objOrdemDeProducao.colItens(iIndice).sVersao

            'verifica se ele é um Kit
            lErro = CF("Kit_Le", objKit)
            If lErro <> SUCESSO And lErro <> 21826 Then gError 62607
            
            If lErro = SUCESSO Then
                
                'Inclui os subprodutos do Kit como itens em uma OP.
                lErro = ProdutoOP_Extrai_KitRC(objKit, objReqCompras, objOrdemDeProducao.colItens(iIndice).dQuantidade)
                If lErro <> SUCESSO Then gError 62608

            End If

        End If
        
    Next

    'Para cada ItemOP na Ordem a ser criada
    For iIndice = objReqCompras.colItens.Count To 1 Step -1

        Set objItemRC = objReqCompras.colItens(iIndice)

        'Busca no BD o saldo atual no Estoque do produto
        lErro = CF("SaldoAtualProdComprado", objItemRC.sProduto, dSaldoAtual)
        If lErro <> SUCESSO Then gError 62609

        If dSaldoAtual >= objItemRC.dQuantidade Then
            objReqCompras.colItens.Remove iIndice
        Else
            objItemRC.dQuantidade = objItemRC.dQuantidade - dSaldoAtual
        End If

    Next
    'Se há algum item na OP gerada
    If objReqCompras.colItens.Count > 0 Then

        'Chama Config_ObterAutomatico
        lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROXIMO_CODIGO_RC", "RequisicaoCompra", "Codigo", lCodigo)
        If lErro <> SUCESSO Then gError 62610
        
        'Preenche os demais dados da OP
        objReqCompras.dtData = gdtDataHoje
        objReqCompras.dtDataBaixa = DATA_NULA
        objReqCompras.dtDataEnvio = DATA_NULA
        objReqCompras.dtDataLimite = DATA_NULA
        objReqCompras.iFilialCompra = giFilialEmpresa
        objReqCompras.iFilialDestino = giFilialEmpresa
        objReqCompras.iFilialEmpresa = giFilialEmpresa
        objReqCompras.iTipoDestino = TIPO_DESTINO_EMPRESA
        objReqCompras.lCodigo = lCodigo
        
        objReqCompras.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO
        objReqCompras.sOPCodigo = objOrdemDeProducao.sCodigo
        
        'Grava a nova ReqCompra gerada
        lErro = CF("RequisicaoCompra_Grava_EmTrans", objReqCompras, True)
        If lErro <> SUCESSO Then gError 62611
        
    End If

    OrdemProducao_Grava_ReqSubprodutos = SUCESSO
    
    Exit Function
    
Erro_OrdemProducao_Grava_ReqSubprodutos:

    OrdemProducao_Grava_ReqSubprodutos = gErr
    
    Select Case gErr
    
        Case 62605, 62607, 62608, 62609, 62610, 62611
        
        Case 62606
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152606)
            
    End Select
    
    Exit Function

End Function

Function ProdutoOP_Extrai_KitRC(ByVal objKit As ClassKit, ByVal objReqCompras As ClassRequisicaoCompras, ByVal dQuantReqCompras As Double) As Long
'Inclui os produtos e subprodutos do Kit passado em uma OP.

Dim objProdutoKit As ClassProdutoKit
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iPosicao As Integer
Dim bAchou As Boolean
Dim dQuantReq As Double
Dim dFator As Double
Dim dSaldoAtual As Double
Dim objItemRC As New ClassItemReqCompras
Dim iAlmoxarifadoPadrao As Integer
Dim dQuantPedido As Double

On Error GoTo Erro_ProdutoOP_Extrai_KitRC

    'Lê os componentes do Kit passado
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 62612

    'Para cada componente encontrado
    For Each objProdutoKit In objKit.colComponentes
        
        If objKit.sProdutoRaiz <> objProdutoKit.sProduto Then

            objProduto.sCodigo = objProdutoKit.sProduto
            'Busca o Produto no BD
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 62613
            If lErro <> SUCESSO Then gError 62614
                
            lErro = CF("Produto_Le_QuantEmPedidoCompra", objProduto.sCodigo, dQuantPedido)
            If lErro <> SUCESSO Then gError 62615
    
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMCompra, dFator)
            If lErro <> SUCESSO Then gError 62616
    
            'Calcula a quantidade de requisicao para esse produto
            
            dQuantReq = ((objProdutoKit.dQuantidade * dFator) * IIf(objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL, dQuantReqCompras, 1))
    
            'Se o produto for compravel
            If objProduto.iCompras = PRODUTO_COMPRAVEL Then
    
                'Verifica se o produto já é um item da RC
                bAchou = CF("BuscaProduto_ColItensRC", objProduto.sCodigo, objReqCompras.colItens, iPosicao)
    
                'Se o produto já é um ItemReq
                If bAchou Then
                    'Se for um item do kit de composição variável
                    If objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                                        
                        'Acrescenta a quant a produzir
                        objReqCompras.colItens(iPosicao).dQuantidade = objReqCompras.colItens(iPosicao).dQuantidade + dQuantReq + dQuantPedido
                    End If
                'Senão
                Else
                               
                    'Cria um ItemRC nojvo com o produto
                    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifadoPadrao)
                    If lErro <> SUCESSO And lErro <> 23796 Then gError 62617
                    If iAlmoxarifadoPadrao = 0 Then gError 62618
                    
                    Set objItemRC = New ClassItemReqCompras
    
                    objItemRC.dQuantidade = dQuantReq
                    objItemRC.iAlmoxarifado = iAlmoxarifadoPadrao
                    objItemRC.sDescProduto = objProduto.sDescricao
                    objItemRC.sProduto = objProduto.sCodigo
                    objItemRC.sUM = objProduto.sSiglaUMCompra
    
                    objReqCompras.colItens.Add objItemRC
    
                End If
            
            Else
                
                objKit.sProdutoRaiz = objProduto.sCodigo
                
                'Verifica se o produto é um Kit
                lErro = CF("Kit_Le", objKit)
                If lErro <> AD_SQL_SUCESSO And lErro <> 21826 Then gError 62619
                
                'Se for
                If lErro = SUCESSO Then
                    'Inclui os seus subprodutos na RC
                    lErro = ProdutoOP_Extrai_KitRC(objKit, objReqCompras, dQuantReq)
                    If lErro <> SUCESSO Then gError 62620
                End If
            End If
        End If
    Next

    ProdutoOP_Extrai_KitRC = SUCESSO

    Exit Function

Erro_ProdutoOP_Extrai_KitRC:

    ProdutoOP_Extrai_KitRC = gErr

    Select Case gErr

        Case 62612, 62613, 62615, 62616, 62617, 62619, 62620

        Case 62614
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 62618
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152607)

    End Select

    Exit Function
    
End Function

Function BuscaProduto_ColItensRC(ByVal sProduto As String, ByVal colItensRC As Collection, iPosicao As Integer) As Boolean
'Busca o produto passado na coleção de Itens de ReqCompras passada
'Se encontrar retorna True e a posição do produto na coleção, se não encontrar retorna Falso.

Dim bAchou As Boolean
Dim objItemRC As ClassItemReqCompras

    iPosicao = 0
    bAchou = False
    
    'Para cada produto na coleção
    For Each objItemRC In colItensRC
        
        'Atualiza a posição do produto
        iPosicao = iPosicao + 1
        
        'Se o produto procurado foi encontrado
        If objItemRC.sProduto = sProduto Then
            'Guarda que achou e sai do For
            bAchou = True
            Exit For
        End If
    Next
    
    'Retorna o resultado da busca para a rotina chamadora
    BuscaProduto_ColItensRC = bAchou
    
    Exit Function

End Function


Function IntervaloProducao_Le(iIntervalo As Integer) As Long
'Busca em MATConfig o Intervalo entre o fim da produção dos insumos de um produto e
'o início na produção desse produto.

Dim lErro As Long
Dim lComando As Long
Dim sConteudo As String

On Error GoTo Erro_IntervaloProducao_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62621
    
    sConteudo = String(STRING_CONTEUDO, 0)
    
    'Busca o valor no BD
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ? ", sConteudo, "INTERVALO_INSUMO_INIC_PRODUCAO", giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 62622
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 62623
    
    'Armazena na var de retorno
    iIntervalo = CStr(sConteudo)
    
    'fecha o comando
    Call Comando_Fechar(lComando)

    IntervaloProducao_Le = SUCESSO

    Exit Function

Erro_IntervaloProducao_Le:

    IntervaloProducao_Le = Err
    
    Select Case Err
    
        Case 62621
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62622, 62623
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", Err, "INTERVALO_INSUMO_INIC_PRODUCAO")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152608)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RequisicaoCompra_Lock_Gravacao(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Faz Lock nas tabelas Associadas a Requisição de Compras

Dim lErro As Long
Dim lObservacao As Long

On Error GoTo Erro_RequisicaoCompra_Lock_Gravacao

    'Se a FilialCompra tiver sido preenchida
    If objRequisicaoCompra.iFilialCompra <> 0 Then
'??? diversos abaixo. Jones 16/01/2002
        'Le e locka FilialEmpresa
''''lErro = CF("FilialEmpresa_Lock",objRequisicaoCompra.iFilialCompra)
        If lErro <> SUCESSO And lErro <> 61725 Then gError (65508)
        If lErro = 61725 Then gError (65509)
        
    End If
    
    'Se o Ccl está preenchido
    If Len(Trim(objRequisicaoCompra.sCcl)) > 0 Then
        
        'Lê o locka Ccl
''''        lErro = CF("Ccl_Lock",objRequisicaoCompra.sCcl)
        If lErro <> SUCESSO And lErro <> 61731 Then gError (65510)
        If lErro = 61731 Then gError (65511)
    
    End If
    
    'Se o Requisitante estiver preenchido
    If objRequisicaoCompra.lRequisitante <> 0 Then
        
        'Lê e Faz Lock do Requisitante
''''        lErro = CF("Requisitante_Lock",objRequisicaoCompra.lRequisitante)
        If lErro <> SUCESSO And lErro <> 61737 Then gError (65512)
        If lErro = 61737 Then gError (65513)
    
    End If
    
    'Trata o Lock do Tipo Destino
    If objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_EMPRESA Then
        
''''        lErro = CF("FilialEmpresa_Lock",objRequisicaoCompra.iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 61725 Then gError (65514)
        
        'Se não encontrou FilialEmpresa, Erro
        If lErro = 61725 Then gError (65515)
        
    ElseIf objRequisicaoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then

        lErro = CF("FilialFornecedor_Lock", objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino)
        If lErro <> SUCESSO And lErro <> 42870 Then gError (65518)

        'Se nao encontrou FilialFornecedor, Erro
        If lErro = 42870 Then gError (65519)

    End If

    'Verifica se observacao esta preenchida
    If Len(Trim(objRequisicaoCompra.sObservacao)) > 0 Then

        'Faz Lock e insere nova Observação
''''        lErro = CF("Observacao_Lock",objRequisicaoCompra.sObservacao, lObservacao)
        If lErro <> SUCESSO Then gError (65520)
        
        objRequisicaoCompra.lObservacao = lObservacao
            
    End If
    
    RequisicaoCompra_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_RequisicaoCompra_Lock_Gravacao:

    RequisicaoCompra_Lock_Gravacao = gErr
    
    Select Case gErr
                
        Case 65508, 65510, 65512, 65514, 65518, 65520
        
        Case 65509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRequisicaoCompra.iFilialCompra)
        
        Case 65511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objRequisicaoCompra.sCcl)
        
        Case 65513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisicaoCompra.lRequisitante)
        
        Case 65515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objRequisicaoCompra.iFilialEmpresa)
                    
        Case 65519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152609)
    
    End Select

    Exit Function
    
End Function
'???? JA EXISTE EM REQCOMPRASOCX
Function ItensReqCompra_Lock_Gravacao(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Faz lock nas tabelas associados ao item da Requisição Compras

Dim lErro As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim objProduto As New ClassProduto
Dim lComando As Long
Dim lObservacao As Long

On Error GoTo Erro_ItensReqCompra_Lock_Gravacao
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError (65521)
    
    'Para cada item da Requisição de Compras
    For Each objItemReqCompra In objRequisicaoCompra.colItens
                
        'Lê e Faz Lock do Produto
        objProduto.sCodigo = objItemReqCompra.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then gError (65522)
        
        'Se o Almoxarifado estiver preenchido
        If objItemReqCompra.iAlmoxarifado <> 0 Then
        
            'Lê e Faz Lock do Almoxarifado
            lErro = CF("Almoxarifado_Lock", objItemReqCompra.iAlmoxarifado, lComando)
            If lErro <> SUCESSO Then gError (65523)
        
        End If
        
        'Se o Ccl estiver preenchido
        If Len(Trim(objItemReqCompra.sCcl)) > 0 Then
        
            'Lê e Faz lock do Ccl
''''            lErro = CF("Ccl_Lock",objItemReqCompra.sCcl)
            If lErro <> SUCESSO And lErro <> 61731 Then gError (65524)
            If lErro = 61731 Then gError (65525)
        
        End If
        
        'Se a Conta Contábil estiver preenchida
        If Len(Trim(objItemReqCompra.sContaContabil)) > 0 Then
            
            'Lê e Faz Lock da Conta Contábil
''''            lErro = CF("ContaContabil_Lock",objItemReqCompra.sContaContabil)
            If lErro <> SUCESSO And lErro <> 61762 Then gError (65526)
            If lErro = 61762 Then gError (65527)
            
        End If
        
        'Se Fornecedor e Filial estiverem Preenchidos
        If objItemReqCompra.lFornecedor <> 0 And objItemReqCompra.iFilial <> 0 Then
        
            'Lê e Faz Lock em FilialFornecedor
            lErro = CF("FilialFornecedor_Lock", objItemReqCompra.lFornecedor, objItemReqCompra.iFilial)
            If lErro <> SUCESSO And lErro <> 42870 Then gError (65528)
            If lErro = 42870 Then gError (65529)
        
        End If
        
        'Se a Observação estiver Preenchida
        If Len(Trim(objItemReqCompra.sObservacao)) > 0 Then
    
            'Lê e Faz Lock em Observação
''''            lErro = CF("Observacao_Lock",objItemReqCompra.sObservacao, lObservacao)
            If lErro <> SUCESSO Then gError (65530)
                        
            objItemReqCompra.lObservacao = lObservacao
        
        End If
        
    Next
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    ItensReqCompra_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_ItensReqCompra_Lock_Gravacao:

    ItensReqCompra_Lock_Gravacao = gErr

    Select Case gErr
    
        Case 65521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 65522, 65523, 65524, 65526, 65528, 65530
                 
        Case 65525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemReqCompra.sCcl)
        
        Case 65527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objItemReqCompra.sContaContabil)
        
        Case 65529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objItemReqCompra.lFornecedor, objItemReqCompra.iFilial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152610)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
'???? JA EXISTE EM REQCOMPRASOCX
Function ItensReqCompra_Exclui(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Exclui os Itens da Requisição de Compras

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntItem As Long

On Error GoTo Erro_ItensReqCompra_Exclui
    
    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66611
    Next
          
    'Lê o item da Requisição
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensReqCompra WHERE ReqCompra = ?", 0, lNumIntItem, objRequisicaoCompras.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66612
        
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66613
        
    'Para cada item de Requisição
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Exclui o Item do banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensReqCompra", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66615
            
        'Busca Próximo Item
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66614
        
    Loop
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensReqCompra_Exclui = SUCESSO
    
    Exit Function
    
Erro_ItensReqCompra_Exclui:
    
    ItensReqCompra_Exclui = gErr

    Select Case gErr
        
        Case 66611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 66612, 66613, 66614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 66615
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMREQCOMPRA", gErr, lNumIntItem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152611)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'???? JA EXISTE EM REQCOMPRASOCX
Function ItensReqCompra_Inclui(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Insere os Itens da Requisição Compras no BD

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim iIndice As Integer
Dim lNumIntItem As Long

On Error GoTo Erro_ItensReqCompra_Inclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError (65531)
    Next
    
    'Para cada item da Requisição de Compras
    For Each objItemReqCompra In objRequisicaoCompra.colItens
        
        'Gera numero interno automatico para o item
        lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_ITEM_RC", lNumIntItem)
        If lErro <> SUCESSO Then gError (65533)

        objItemReqCompra.lNumIntDoc = lNumIntItem
    
        'Grava o Item da Requisição de Compra no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO ItensReqCompraN (NumIntDoc, ReqCompra, Produto, DescProduto, Quantidade, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo, StatusBaixa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemReqCompra.lNumIntDoc, objRequisicaoCompra.lNumIntDoc, objItemReqCompra.sProduto, objItemReqCompra.sDescProduto, objItemReqCompra.dQuantidade, objItemReqCompra.sUM, objItemReqCompra.sCcl, objItemReqCompra.iAlmoxarifado, objItemReqCompra.sContaContabil, objItemReqCompra.iTipoTributacao, objItemReqCompra.lObservacao, objItemReqCompra.lFornecedor, objItemReqCompra.iFilial, objItemReqCompra.iExclusivo, ITENSREQCOMPRA_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError (65532)

    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensReqCompra_Inclui = SUCESSO
    
    Exit Function
    
Erro_ItensReqCompra_Inclui:

    ItensReqCompra_Inclui = gErr

    Select Case gErr
            
        Case 65531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 65532
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMREQCOMPRA", gErr, objItemReqCompra.lNumIntDoc)
            
        Case 65533
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152612)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function OrdemDeProducao_Exclui_EmTrans(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao

Dim lErro As Long
Dim iItem As Integer
Dim iIndice As Integer
Dim iSituacao As Integer, sProduto As String, iAlmoxarifado As Integer, dQtde As Double
Dim iNumItens As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocGrade As Long
Dim alComando(1 To 7) As Long
Dim objItemOP As New ClassItemOP
Dim lCodPedido As Long
Dim objOPGerada As New ClassOrdemDeProducao
Dim objReqCompra As New ClassRequisicaoCompras
Dim sCodigo As String
Dim lCodigo As Long
Dim objItemOPGrade As ClassItemOP

'########################
'INSERIDO POR WAGNER
Dim lNumIntOPGrade As Long
'########################

Dim iAux As Integer    'Inserido por Jorge Specian - 20/07/2005

On Error GoTo Erro_OrdemDeProducao_Exclui_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 21932
    Next

    'tenta ler se existe a OP
    lErro = Comando_ExecutarPos(alComando(3), "SELECT NumItens FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 21934

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21935

    If lErro = AD_SQL_SEM_DADOS Then gError 21936
    
    'Inserido por Jorge Specian - 20/07/2005
    '---------------------------------------
    'Verifica se esta OP se encontra no Plano Mestre de Producao
    lErro = Comando_Executar(alComando(7), "SELECT 1 FROM PlanoOperacional As PO Where PO.CodOPOrigem = ? And Exists (SELECT PMPItens.NumIntDoc " & _
                    "FROM PlanoMestreProducaoItens AS PMPItens, " & _
                    "PlanoMestreProducaoUltimo AS PMP " & _
                    "WHERE PMP.Codgeracao = PMPItens.CodGeracao AND " & _
                    "PMPItens.NumIntDoc = PO.NumIntDocPMP AND PMPItens.FilialEmpresa = ?)", iAux, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)

    lErro = Comando_BuscarPrimeiro(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137914
    
    'Se está ... Erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 137915
    '---------------------------------------

    lErro = Comando_LockExclusive(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 21937

    sCodigo = String(STRING_OPCODIGO, 0)
    
    'Verifica se a OP é uma geradora
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM OrdensDeProducao WHERE OPGeradora = ? AND FilialEmpresa = ?", sCodigo, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 62624
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62625
    
    'Se for
    '#####################################
    'ALTERADO POR WAGNER
    'Uma OP passa a poder ter mais de um filho
    'If lErro <> AD_SQL_SEM_DADOS Then
     Do While lErro <> AD_SQL_SEM_DADOS
     
        objOPGerada.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
        objOPGerada.sCodigo = sCodigo
        'Exclui a OP gerada por ela
        lErro = OrdemDeProducao_Exclui_EmTrans(objOPGerada)
        If lErro <> SUCESSO Then gError 62626
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129986
                
     Loop
    'End If
    '#####################################
    
    'Verifica se a OP tem Reqs associadas a ela
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM RequisicaoCompra WHERE OPCodigo = ? AND FilialEmpresa = ?", lCodigo, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 62627
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62628
    Do While lErro <> AD_SQL_SEM_DADOS
    
        objReqCompra.lNumIntDoc = lCodigo
        
        lErro = CF("RequisicaoCompra_Exclui_EmTrans", objReqCompra)
        If lErro <> SUCESSO Then gError 62629
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62630
    
    Loop
        
    'seleciona os itens da OP
    sProduto = String(STRING_PRODUTO, 0)
    
    '####################################
    'ALTERADO POR WAGNER
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Item, Situacao, Produto, Quantidade, Almoxarifado, CodPedido, NumIntItemOPGrade FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, iItem, iSituacao, sProduto, dQtde, iAlmoxarifado, lCodPedido, lNumIntOPGrade, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 21938
    '####################################

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21939

    'percorre os itens
    Do While lErro <> AD_SQL_SEM_DADOS

        If iSituacao <> ITEMOP_SITUACAO_NORMAL And iSituacao <> ITEMOP_SITUACAO_DESAB Then gError 21941

        objItemOP.lNumIntDoc = lNumIntDoc
        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
        objItemOP.sProduto = sProduto
        objItemOP.iAlmoxarifado = iAlmoxarifado
        objItemOP.dQuantidade = dQtde
        objItemOP.lCodPedido = lCodPedido
        objItemOP.iItem = iItem

        sProduto = String(STRING_PRODUTO, 0)

        'seleciona os itens da OP Grade se houverem
        '##############################################
        'ALTERADO POR WAGNER
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, Produto, Quantidade, Almoxarifado, CodPedido  FROM ItensOPGrade WHERE NumIntDoc = ? ", 0, lNumIntDocGrade, sProduto, dQtde, iAlmoxarifado, lCodPedido, lNumIntOPGrade)
        If lErro <> AD_SQL_SUCESSO Then gError 126539
        '##############################################
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126540
            
        '#############################################
        'ALTERADO POR WAGNER
        'Do While lErro = AD_SQL_SUCESSO
        If lErro <> AD_SQL_SEM_DADOS Then
    
            Set objItemOPGrade = New ClassItemOP
    
            objItemOPGrade.lNumIntDoc = lNumIntDocGrade
            objItemOPGrade.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
            objItemOPGrade.sProduto = sProduto
            objItemOPGrade.iAlmoxarifado = iAlmoxarifado
            objItemOPGrade.dQuantidade = dQtde
            objItemOPGrade.lCodPedido = lCodPedido
            objItemOPGrade.iItem = iItem
    
            'verifica se nao tem movimentoestoque associado
            'exclui os empenhos associados
            'diminui a quantidade em op em EstoqueProduto
            'diminui a quantidade em OP do Pedido de Venda
            lErro = OrdemDeProducao_Exclui_EmTrans1(objOrdemDeProducao, objItemOPGrade)
            If lErro <> SUCESSO Then gError 126537
    
            lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM ItensOPGrade", alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 126541
            
            lErro = ItemOPItemPV_Exclui(objItemOPGrade)
            If lErro <> SUCESSO Then gError 126543
            
            'lErro = Comando_BuscarProximo(alComando(2))
            'If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126542
            
         End If
        'Loop
        '#############################################
            
        'verifica se nao tem movimentoestoque associado
        'exclui os empenhos associados
        'diminui a quantidade em op em EstoqueProduto
        'diminui a quantidade em OP do Pedido de Venda
        lErro = OrdemDeProducao_Exclui_EmTrans1(objOrdemDeProducao, objItemOP)
        If lErro <> SUCESSO Then gError 126543

        'Incluido por Jorge Specian - 11/05/2005
        '---------------------------------------
        lErro = CF("OrdemDeProducao_Exclui_Operacoes_EmTrans", objItemOP)
        If lErro <> SUCESSO Then gError 137008
        '---------------------------------------

        lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM ItensOrdemProducao", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 126544
        
        lErro = ItemOPItemPV_Exclui(objItemOP)
        If lErro <> SUCESSO Then gError 126543
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21947

    Loop

    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM OrdensDeProducao", alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 21948
    
    lErro = CF("OrdemDeProducao_Exclui_Usu", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 181764

    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_OP, 0, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 181764
    '############################################

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui_EmTrans:

    OrdemDeProducao_Exclui_EmTrans = gErr

    Select Case gErr

        Case 21932
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 21934, 21935, 62624, 62625, 129986 'Inserido por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)

        Case 21936, 62626, 62629, 126537, 126543, 137008, 181764

        Case 21938, 21939, 21947
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)

        Case 21941
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_SITUACAO_NAO_EXCLUIVEL", gErr, iItem, objOrdemDeProducao.sCodigo)

        Case 21948
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORDENSDEPRODUCAO", gErr)

        Case 62627, 62628, 62630
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr)
        
        Case 126539, 126540, 126542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPGRADE", gErr, lNumIntDoc)
        
        Case 126541
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSOPGRADE", gErr)
        
        Case 126544
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSORDENSDEPRODUCAO", gErr)
            
        'Inserido por Jorge Specian - 20/07/2005
        '---------------------------------------
        Case 137913, 137914
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAOITENS", gErr)
            
        Case 137915
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_PARTICIPA_PLANOMESTREPRODUCAO", gErr)
        '---------------------------------------
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152613)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OrdemDeProducao_Exclui_EmTrans1(objOrdemDeProducao As ClassOrdemDeProducao, objItemOP As ClassItemOP) As Long
'verifica se nao tem movimentoestoque associado
'exclui os empenhos associados
'diminui a quantidade em op em EstoqueProduto
'diminui a quantidade em OP do Pedido de Venda

Dim lErro As Long
Dim iIndice As Integer
Dim lNumInt As Long
Dim alComando(1 To 3) As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objItemPV As New ClassItemPedido
Dim objProduto As New ClassProduto
Dim objItemOPItemPV As ClassItemOPItemPV

On Error GoTo Erro_OrdemDeProducao_Exclui_EmTrans1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126538
    Next

    objProduto.sCodigo = objItemOP.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 126571

    If lErro <> SUCESSO Then gError 126572

    'pesquisa movimento de estoque associado ao item da OP
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM MovimentoEstoque WHERE FilialEmpresa = ? AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", lNumInt, objOrdemDeProducao.iFilialEmpresa, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 21942

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21943

    If lErro = AD_SQL_SUCESSO Then gError 21944

    '#########################################
    'Inserido por Wagner - CROMATON 08/11/04
    objOrdemDeProducao.colItens.Add objItemOP
    '#########################################
    
    'exclui empenhos associados ao itemOP
    lErro = CF("Empenhos_Excluir_ItemOP", objItemOP)
    If lErro <> SUCESSO Then gError 21945

    'se nao for um produto pai de grade
    If Len(Trim(objProduto.sGrade)) = 0 And objItemOP.iTipo = OP_TIPO_OP Then

        'diminui qtde em op associado ao produto-almoxarifado
        objEstoqueProduto.sProduto = objItemOP.sProduto
        objEstoqueProduto.iAlmoxarifado = objItemOP.iAlmoxarifado
        
        lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, -objItemOP.dQuantidade)
        If lErro <> SUCESSO Then gError 41553
    
    End If
    
    'acerta a quantOP em Item de Pedido de Venda
    If objItemOP.lCodPedido <> 0 Then

        objItemPV.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
        objItemPV.lCodPedido = objItemOP.lCodPedido
        objItemPV.sProduto = objItemOP.sProduto
        objItemPV.iTipoOP = objItemOP.iTipo
    
        'diminui a quantidade em OP do Pedido de Venda
        lErro = OrdemDeProducao_Grava6(alComando(2), alComando(3), objItemPV, -objItemOP.dQuantidade)
        If lErro <> SUCESSO And lErro <> 41346 Then gError 55880
    
        If lErro = 41346 Then gError 55881
        
    Else
    
        lErro = CF("ItemOPItemPV_Le2", objItemOP)
        If lErro <> SUCESSO Then gError 41553
        
        For Each objItemOPItemPV In objItemOP.colItemPV
        
            Set objItemPV = New ClassItemPedido
            
            objItemPV.iFilialEmpresa = objItemOPItemPV.iFilialEmpresa
            objItemPV.lCodPedido = objItemOPItemPV.lCodigoPV
            objItemPV.sProduto = objItemOP.sProduto
            objItemPV.iTipoOP = objItemOP.iTipo
        
            'diminui a quantidade em OP do Pedido de Venda
            lErro = OrdemDeProducao_Grava6(alComando(2), alComando(3), objItemPV, -objItemOPItemPV.dQuantidade)
            If lErro <> SUCESSO And lErro <> 41346 Then gError 55880
        
            If lErro = 41346 Then gError 55881
        
        Next
        
    End If

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Exclui_EmTrans1 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui_EmTrans1:

    OrdemDeProducao_Exclui_EmTrans1 = gErr

    Select Case gErr

        Case 21942, 21943
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 21944
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_ASSOCIADO_MOVESTOQUE", gErr, objItemOP.iItem, objOrdemDeProducao.sCodigo)

        Case 21945, 41533, 55880

        Case 21946
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSORDENSDEPRODUCAO", gErr)

        Case 55881
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemPV.lCodPedido, objItemPV.sProduto)
        
        Case 126538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 126572
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152614)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function OrdemDeProducao_Exclui(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OrdemDeProducao_Exclui

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 21933

    lErro = CF("OrdemDeProducao_Exclui1", objOrdemDeProducao)
    If lErro <> SUCESSO Then Error 62563

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 21949

    OrdemDeProducao_Exclui = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui:

    OrdemDeProducao_Exclui = Err

    Select Case Err

        Case 21933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 62563
        
        Case 21949
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152615)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

'???catharine
Private Function OrdemDeProducaoBaixada_Reativar(ByVal objOrdemDeProducao As ClassOrdemDeProducao, alComando() As Long) As Long
'Transforma uma Ordem de Producao Baixada em Ativa novamente

' *** Função alterada por Luiz G.F.Nogueira em 10/08/01 ***
' *** Alteração feita para ler em itens baixados e gravar em itens
' o campo Beneficiamento ***

Dim lErro As Long
Dim iItensAtivados  As Integer
Dim lNumIntDoc As Long
Dim sCodigo As String
Dim sOPGeradora  As String
Dim tOrdemProducao As typeOrdemProducao
Dim objItemOP As ClassItemOP
Dim objItemOPGrade As ClassItemOP
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dValorAtual As Double
Dim dValorInicial As Double
Dim dValorFinal As Double
Dim dValorApropriado As Double
Dim dValorApropInicial As Double
Dim dValorApropFinal As Double, lEmpenho As Long, objProduto As New ClassProduto
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_OrdemDeProducaoBaixada_Reativar

    lErro = OrdemDeProducao_Grava31(alComando(13), lEmpenho)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iItensAtivados = 0
    
    For Each objItemOP In objOrdemDeProducao.colItens
    
        If objItemOP.iPossuiGrade = DESMARCADO Then
    
            objProduto.sCodigo = objItemOP.sProduto
                
            lErro = CF("Produto_le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            objItemOP.iClasseUM = objProduto.iClasseUM
            objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
            
            lErro = ItemOP_Reativar(objItemOP, iItensAtivados)
            If lErro <> SUCESSO Then gError 126709
        
            If objItemOP.iTipo = OP_TIPO_OP And objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
        
                'cria os empenhos
                lErro = CF("OrdemDeProducao_Grava2", objItemOP, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
                If lErro <> SUCESSO Then gError 126709
        
                lErro = OrdemDeProducao_Grava7(objItemOP.lNumIntDoc)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If
        
        Else
        
            For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade

                Set objItemOPGrade = New ClassItemOP
    
                objItemOPGrade.dQuantidade = objItemRomaneioGrade.dQuantidade
                objItemOPGrade.iAlmoxarifado = objItemRomaneioGrade.iAlmoxarifado
                objItemOPGrade.sProdutoPai = objItemOP.sProduto
                objItemOPGrade.sProduto = objItemRomaneioGrade.sProduto
                objItemOPGrade.sVersao = objItemRomaneioGrade.sVersao
                objItemOPGrade.lNumIntDoc = objItemRomaneioGrade.lNumIntDoc
                objItemOPGrade.lNumItemOP = objItemOP.lNumIntDoc
                objItemOPGrade.iFilialPedido = objItemOP.iFilialPedido
                objItemOPGrade.lCodPedido = objItemOP.lCodPedido
                objItemOPGrade.lNumIntOrigem = objItemOP.lNumIntOrigem
                objItemOPGrade.iOrigemPedido = objItemOP.iOrigemPedido
                objItemOPGrade.iClasseUM = objItemOP.iClasseUM
                objItemOPGrade.sSiglaUMEstoque = objItemOP.sSiglaUMEstoque
                objItemOPGrade.sSiglaUM = objItemOP.sSiglaUM
                objItemOPGrade.iFilialEmpresa = objItemOP.iFilialEmpresa
                objItemOPGrade.sCodigo = objItemOP.sCodigo
                objItemOPGrade.iItem = objItemOP.iItem
                objItemOPGrade.sCcl = objItemOP.sCcl
                objItemOPGrade.dtDataInicioProd = objItemOP.dtDataInicioProd
                objItemOPGrade.dtDataFimProd = objItemOP.dtDataFimProd
                objItemOPGrade.iPrioridade = objItemOP.iPrioridade
                objItemOPGrade.iSituacao = objItemOP.iSituacao
                objItemOPGrade.iDestinacao = objItemOP.iDestinacao
                objItemOPGrade.iBeneficiamento = objItemOP.iBeneficiamento
                objItemOPGrade.lNumIntEquipamento = objItemOP.lNumIntEquipamento
                objItemOPGrade.lNumIntEquipamento2 = objItemOP.lNumIntEquipamento2
                objItemOPGrade.iPossuiGrade = DESMARCADO
                objItemOPGrade.iTipo = objItemOP.iTipo
            
                lErro = ItemOP_Reativar(objItemOPGrade, iItensAtivados)
                If lErro <> SUCESSO Then gError 126710
            
            Next
            
        End If
    
    Next
    
    'Exclui os Itens da Ordem de Producao Baixada da tabela de ItensOrdemProducaoBaixadas
    lErro = ItensOpBaixada_Exclui(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 82819
    
    'Transfere a Ordem de Producao Baixada para a tabela de Ordem de Producao
    lErro = OrdemDeProducaoBaixada_Desfaz_Baixa(objOrdemDeProducao, iItensAtivados)
    If lErro <> SUCESSO Then gError 82821
    
    lErro = OrdemDeProducao_Grava51(alComando(13), lEmpenho)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    OrdemDeProducaoBaixada_Reativar = SUCESSO
    
    Exit Function
    
Erro_OrdemDeProducaoBaixada_Reativar:

    OrdemDeProducaoBaixada_Reativar = gErr
    
    Select Case gErr
    
        Case 82819, 82821, 126709, 126710, ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152616)
            
    End Select
    
    Exit Function
    
End Function

Private Function ItemOP_Reativar(objItemOP As ClassItemOP, iItensAtivados As Integer) As Long

Dim tItemOP As typeItemOP
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 7) As Long
Dim dValorAtual As Double
Dim dValorInicial As Double
Dim dValorFinal As Double
Dim dValorApropriado As Double
Dim dValorApropInicial As Double
Dim dValorApropFinal As Double
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_ItemOP_Reativar

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 82805
    Next

        tItemOP.sCodigo = String(STRING_OPCODIGO, 0)
        tItemOP.sCcl = String(STRING_CCL, 0)
        tItemOP.sProduto = String(STRING_PRODUTO, 0)
        tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
        tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
        
        'Verifica se é um item baixado
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntItemOPGrade, NumIntDoc, FilialEmpresa,Codigo,Item,Produto,FilialPedido,CodPedido,NumIntOrigem,Almoxarifado,Ccl, SiglaUM,Quantidade,DataInicioProd,DataFimProd,Prioridade,QuantidadeProd,Situacao,Destinacao,Beneficiamento,ValorAtual,ValorInicial,ValorFinal, " _
            & "ValorApropriado,ValorApropInicial,ValorApropFinal, Versao, NumIntEquipamento, NumIntEquipamento2, Tipo, ProduzLogo FROM ItensOrdemProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo=? AND Produto=?", 0, _
            tItemOP.lNumIntItemOP, tItemOP.lNumIntDoc, tItemOP.iFilialEmpresa, tItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.lNumIntOrigem, tItemOP.iAlmoxarifado, _
            tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.iBeneficiamento, dValorAtual, dValorInicial, dValorFinal, dValorApropriado, dValorApropInicial, dValorApropFinal, tItemOP.sVersao, tItemOP.lNumIntEquipamento, tItemOP.lNumIntEquipamento2, tItemOP.iTipo, tItemOP.iProduzLogo, objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 82806
        
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82807
        
        If lErro = AD_SQL_SUCESSO Then
            
            objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
            
            'Verifica se a situacao do Item (na tela) é normal
            If objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
            
                If objItemOP.dQuantidade > tItemOP.dQuantidadeProd And tItemOP.iTipo = OP_TIPO_OP Then
                
                    objEstoqueProduto.sProduto = objItemOP.sProduto
                    objEstoqueProduto.iAlmoxarifado = objItemOP.iAlmoxarifado
                    
                    'Atualiza a tabela de EstoqueProduto
                    lErro = CF("EstoqueProduto_AlterarQtdOP", objEstoqueProduto, (objItemOP.dQuantidade - tItemOP.dQuantidadeProd))
                    If lErro <> SUCESSO Then gError 82818
                
                End If
                                    
                'Atualiza o número de itens ativados
                iItensAtivados = iItensAtivados + 1
            End If
            
            'Grava o Item em ItemOrdemDeProducao
            lErro = Comando_Executar(alComando(7), "INSERT INTO ItensOrdemProducao (NumIntItemOPGrade, NumIntDoc, FilialEmpresa, Codigo, Item, Produto, FilialPedido," & _
                "CodPedido, NumIntOrigem, Almoxarifado, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd," & _
                "Situacao, Destinacao, Beneficiamento, Versao, NumIntEquipamento, NumIntEquipamento2, Tipo, ProduzLogo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                tItemOP.lNumIntItemOP, tItemOP.lNumIntDoc, tItemOP.iFilialEmpresa, objItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.lNumIntOrigem, _
                tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, objItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, _
                tItemOP.dQuantidadeProd, objItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.iBeneficiamento, tItemOP.sVersao, tItemOP.lNumIntEquipamento, tItemOP.lNumIntEquipamento2, tItemOP.iTipo, tItemOP.iProduzLogo)
            If lErro <> SUCESSO Then gError 82816
    
        End If
        

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
           
    ItemOP_Reativar = SUCESSO
    
    Exit Function
    
Erro_ItemOP_Reativar:

    ItemOP_Reativar = gErr
    
    Select Case gErr
    
        Case 82805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 82806, 82807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAOBAIXADAS", gErr)
                        
        Case 82816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMORDEMPRODUCAO", gErr)
            
        Case 82818
            
        Case 82819, 82821
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152617)
            
    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function OrdemDeProducaoBaixada_Desfaz_Baixa(ByVal objOrdemDeProducao As ClassOrdemDeProducao, ByVal iItensAtivados As Integer) As Long
'Transfere a Ordem de Producao Baixada para a tabela de Ordem de Producao

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim sOPGeradora  As String, lCodPrestador As Long
Dim tOrdemProducao As typeOrdemProducao
Dim iIndice As Integer
Dim sUsuario As String

On Error GoTo Erro_OrdemDeProducaoBaixada_Desfaz_Baixa

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 82822
    Next
    
    tOrdemProducao.sCodigo = String(STRING_OPCODIGO, 0)
    sOPGeradora = String(STRING_OPCODIGO, 0)
    sUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    'Busca no BD a Ordem de Producao Baixada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa,Codigo,DataEmissao,NumItens,NumItensBaixados,OPGeradora,CodPrestador,Tipo, NumIntDocOper, CodigoNumerico,Usuario, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo=?", 0, tOrdemProducao.iFilialEmpresa, tOrdemProducao.sCodigo, tOrdemProducao.dtDataEmissao, tOrdemProducao.iNumItens, tOrdemProducao.iNumItensBaixados, sOPGeradora, lCodPrestador, tOrdemProducao.iTipo, tOrdemProducao.lNumIntDocOper, tOrdemProducao.lCodigoNumerico, sUsuario, tOrdemProducao.iTipoTerc, tOrdemProducao.lCodTerc, tOrdemProducao.iFilialTerc, tOrdemProducao.iEscaninhoTerc, tOrdemProducao.iIgnoraEst, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82823
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82813
    
    'Insere a Ordem de Producao na tabela de OrdensdeProducao
    lErro = Comando_Executar(alComando(1), "INSERT INTO OrdensDeProducao (FilialEmpresa,Codigo,DataEmissao,NumItens,NumItensBaixados,OPGeradora,CodPrestador,Tipo, NumIntDocOper, CodigoNumerico,Usuario, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tOrdemProducao.iFilialEmpresa, tOrdemProducao.sCodigo, tOrdemProducao.dtDataEmissao, tOrdemProducao.iNumItens, tOrdemProducao.iNumItensBaixados - iItensAtivados, sOPGeradora, lCodPrestador, tOrdemProducao.iTipo, tOrdemProducao.lNumIntDocOper, tOrdemProducao.lCodigoNumerico, sUsuario, tOrdemProducao.iTipoTerc, tOrdemProducao.lCodTerc, tOrdemProducao.iFilialTerc, tOrdemProducao.iEscaninhoTerc, tOrdemProducao.iIgnoraEst)
    If lErro <> AD_SQL_SUCESSO Then gError 82814
    
    'Exclui a OP da tabela de OrdensdeProducaoBaixadas
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM OrdensDeProducaoBaixadas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 82815

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
           
    OrdemDeProducaoBaixada_Desfaz_Baixa = SUCESSO
    
    Exit Function
    
Erro_OrdemDeProducaoBaixada_Desfaz_Baixa:

    OrdemDeProducaoBaixada_Desfaz_Baixa = gErr
    
    Select Case gErr
    
        Case 82813, 82823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOBAIXADA", gErr)
        
        Case 82814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDEPRODUCAO", gErr)
            
        Case 82815
            '???lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORDENSPRODUCAOBAIXADAS", gErr)
    
        Case 82822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152618)
            
    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function ItensOpBaixada_Exclui(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Exclui os Itens da Ordem de Producao Baixada da tabela de ItensOrdemProducaoBaixadas

Dim lErro As Long
Dim tItemOP As typeItemOP
Dim alComando(0 To 2) As Long
Dim sCodigo As String
Dim iIndice As Integer

On Error GoTo Erro_ItensOpBaixada_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 82820
    Next

    sCodigo = String(STRING_OPCODIGO, 0)
    
    'Lê no BD os Itens da Ordem de Producao Baixada em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensOrdemProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo=?", 0, tItemOP.lNumIntDoc, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> SUCESSO Then gError 82809
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82810
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui o item da tabela ItensOrdemProducaoBaixadas
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensOrdemProducaoBaixadas", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 82811
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82812
        
    Loop
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
           
    ItensOpBaixada_Exclui = SUCESSO
    
    Exit Function
    
Erro_ItensOpBaixada_Exclui:

    ItensOpBaixada_Exclui = gErr
    
    Select Case gErr
    
        Case 82809, 82810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAOBAIXADAS", gErr)
        
        Case 82811
           ''' lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOPBAIXADA", gErr)
        
        Case 82812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOBAIXADA", gErr)
        
        Case 82820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152619)
            
    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function


'--------------------------------------------------------------------
'??? Fernando fazer chama_funções, quando for passar
'para ClassPVSelect

'William
'Data 05/04/01
'Função: PedidoVenda_Busca_ReqCompra
Public Function PedidoVenda_Busca_ReqCompra(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Verifica se existe requisição de compra, baixada ou não antes de excluir um pedido de venda

Dim lErro As Long
Dim lComando As Long
Dim lReqCompra As Long

On Error GoTo Erro_PedidoVenda_Busca_ReqCompra

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87530
    
    'Faz leitura na pedido de venda
    lErro = Comando_Executar(lComando, "SELECT ReqCompra FROM PedidosDeVenda WHERE FilialEmpresa=? AND Codigo=?", lReqCompra, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 87526
    
    'Posiciona o ponteiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87527
    
    'Se o campo ReqCompra é <> 0
        '==> Existe Requisicao de Compras para este pedido - Erro
    If lReqCompra <> 0 Then gError 87528
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    If lErro <> SUCESSO Then gError 87529
    
    PedidoVenda_Busca_ReqCompra = SUCESSO

    Exit Function
    
Erro_PedidoVenda_Busca_ReqCompra:

    PedidoVenda_Busca_ReqCompra = gErr
    
    Select Case gErr
    
        Case 87530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 87526, 87527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)
           
        Case 87528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQCOMPRA_EXISTENTE", gErr, lReqCompra)
 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152620)
    
    End Select
    
    'Fecha Comandos - Saida por erro
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function PedidoVenda_EmbalagensGrava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava as embalagens associadas ao pedido de venda

Dim lErro As Long, objItemPV As ClassItemPedido
Dim objEmbalagensItem As ClassEmbalagensItem, colEmbalagensItem As New Collection

On Error GoTo Erro_PedidoVenda_EmbalagensGrava

    'transfere as colecoes de embalagens a nivel de item p/uma colecao unica
    'acertando o numintdoc dos itens
    For Each objItemPV In objPedidoVenda.colItensPedido
        
        For Each objEmbalagensItem In objItemPV.colEmbalagens
        
            objEmbalagensItem.lNumIntDocItem = objItemPV.lNumIntDoc
            colEmbalagensItem.Add objEmbalagensItem
            
        Next
        
    Next
        
    If colEmbalagensItem.Count <> 0 Then
    
        'Insere as embalagens do pedido de venda no BD
        lErro = CF("EmbalagensItem_Grava", colEmbalagensItem)
        If lErro <> SUCESSO Then gError 100010
    
    End If
    
    PedidoVenda_EmbalagensGrava = SUCESSO

    Exit Function
    
Erro_PedidoVenda_EmbalagensGrava:

    PedidoVenda_EmbalagensGrava = gErr
    
    Select Case gErr
    
        Case 100010
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152621)
    
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_EmbalagensExclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui as embalagens associadas ao pedido de venda

Dim lErro As Long, colEmbalagensItem As New Collection
Dim objEmbalagensItem As ClassEmbalagensItem
Dim lComando As Long, lNumIntDocItem As Long, iEmbalagem As Integer, sProduto As String

On Error GoTo Erro_PedidoVenda_EmbalagensExclui

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 100012
    
    sProduto = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDocItem, Embalagem, ItensPedidoDeVenda.Produto FROM PedidosDeVenda, ItensPedidoDeVenda, EmbalagensItem WHERE PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND ItensPedidoDeVenda.NumIntDoc = EmbalagensItem.NumIntDocItem AND TipoDocOrigem = ? AND PedidosDeVenda.FilialEmpresa = ? AND PedidosDeVenda.Codigo = ?", lNumIntDocItem, iEmbalagem, sProduto, DOCORIGEM_PV, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 100013
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100014
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objEmbalagensItem = New ClassEmbalagensItem
            
        With objEmbalagensItem
        
            .iTipoDocOrigem = DOCORIGEM_PV
            .lNumIntDocItem = lNumIntDocItem
            .sProduto = sProduto
            .iEmbalagem = iEmbalagem
        
        End With
        
        Call colEmbalagensItem.Add(objEmbalagensItem)
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100015
    
    Loop
    
    If colEmbalagensItem.Count <> 0 Then
    
        lErro = CF("EmbalagensItem_Exclui", colEmbalagensItem)
        If lErro <> SUCESSO Then gError 1000016
    
    End If
    
    Call Comando_Fechar(lComando)
    
    PedidoVenda_EmbalagensExclui = SUCESSO

    Exit Function
    
Erro_PedidoVenda_EmbalagensExclui:

    PedidoVenda_EmbalagensExclui = gErr
    
    Select Case gErr
    
        Case 100012
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 100013, 100014, 100015
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGENSPV", gErr)
        
        Case 100016
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152622)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Por leo em 25/02/02
Public Function BloqueioDiasAtraso_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de Dias de Atraso, se o cliente não tiver parcelapag em aberto com vcto anterior ou igual a
'gdtDataHoje - gobjCRFAT.iDiasAtraso

Dim lErro As Long
Dim lComando As Long
Dim dtDataLimite As Date
Dim lNumIntDoc As Long

On Error GoTo Erro_BloqueioDiasAtraso_Adiciona

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94359
    
    'Verifica a data limite p/ atraso de uma parcela antes do bloqueio por atraso. Retorna DATA_NULA caso a empresa não use o Bloqueio por atraso.
    lErro = CF("Verifica_DataLimite_Bloqueio_Atraso", dtDataLimite)
    If lErro <> SUCESSO Then gError 94390
    
    If dtDataLimite <> DATA_NULA Then
    
        If objPedidoVenda.iFilialEmpresa > 50 Then
    
            'Seleciona parcelasrec do Cliente recebido em objPedidoVenda que esteja em atraso de acordo com os dias máximos p/ atraso definido p/ Bloqueios por atraso de pagamento.
            lErro = Comando_Executar(lComando, "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND ParcelasRec.DataVencimentoReal <= ? AND TitulosRec.Cliente = ? AND TitulosRec.FilialEmpresa > 50", lNumIntDoc, STATUS_EXCLUIDO, STATUS_BAIXADO, dtDataLimite, objPedidoVenda.lCliente)
            If lErro <> SUCESSO Then gError 94366
        
        Else
        
            'Seleciona parcelasrec do Cliente recebido em objPedidoVenda que esteja em atraso de acordo com os dias máximos p/ atraso definido p/ Bloqueios por atraso de pagamento.
            lErro = Comando_Executar(lComando, "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND ParcelasRec.DataVencimentoReal <= ? AND TitulosRec.Cliente = ?", lNumIntDoc, STATUS_EXCLUIDO, STATUS_BAIXADO, dtDataLimite, objPedidoVenda.lCliente)
            If lErro <> SUCESSO Then gError 94366
        
        
        End If
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94367
        
        'se existe pelo menos uma parcela em atraso
        If lErro = SUCESSO Then
    
            'Adiciona o bloqueio ao Pedido
            objPedidoVenda.colBloqueiosPV.Add giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_DIAS_ATRASO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""
    
        End If
    
    End If
    
    Call Comando_Fechar(lComando)
    
    BloqueioDiasAtraso_Adiciona = SUCESSO
     
    Exit Function
    
Erro_BloqueioDiasAtraso_Adiciona:

    BloqueioDiasAtraso_Adiciona = gErr
     
    Select Case gErr
          
        Case 94390
        
        Case 94366, 94367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
        
        Case 94359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152623)
     
    End Select
     
    Call Comando_Fechar(lComando)
     
    Exit Function
    
End Function

'Incluida por Leo em 25/02/02
Public Function BloqueiosPV_Exclui_Tipo(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal iTipoBloqueio As Integer) As Long
'Exclui os Bloqueios do tipo recebido como parâmetro não liberados ligados ao Pedido de Venda passado
'Função Genérica

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iTipoBloqAux As Integer

On Error GoTo Erro_BloqueiosPV_Exclui_Tipo

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94373

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 94374

    'Busca os Bloqueios de Crédito ligados ao Pedido
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND TipoDeBloqueio = ? AND DataLib = ?", 0, iTipoBloqAux, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iTipoBloqueio, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 94375

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94376

    'Para cada Bloqueio encontrado
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui o bloqueio da tabela de Bloqueios de Pedido de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BloqueiosPV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 94377

        'Busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94378

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    BloqueiosPV_Exclui_Tipo = SUCESSO

    Exit Function

Erro_BloqueiosPV_Exclui_Tipo:

    BloqueiosPV_Exclui_Tipo = gErr

    Select Case gErr

        Case 94373, 94374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 94375, 94376, 94378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", gErr, objPedidoVenda.lCodigo)

        Case 94377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOSPV", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152624)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Public Function Verifica_DataLimite_Bloqueio_Atraso(dtDataLimite As Date) As Long
'Verifica a Data Limite p/ não haver bloqueio por atraso de parcelas
'gdtDataHoje - gobjCRFAT.iDiasAtraso
'Retorna DATA_NULA caso a empresa não use o Bloqueio por atraso de pagamento.
Dim iDias As Integer
Dim lErro As Integer

    'Se existe Nºde dias máximos p/ bloqueio por atraso de pagamento cadastrado pela empresa.
    If gobjCRFAT.iDiasBloqueioAtraso <> BLOQUEIO_ATRASO_NAO_BLOQUEAR Then

         'Alterado por Cyntia em 21/05/02
        lErro = CF("Dias_Uteis_Antes", gdtDataHoje, gobjCRFAT.iDiasBloqueioAtraso, dtDataLimite)
        If lErro <> SUCESSO Then gError 99306
                
    Else
        
        'DATA_NULA caso a empresa não use o Bloqueio por atraso de pagamento.
        dtDataLimite = DATA_NULA
            
    End If
    
    Verifica_DataLimite_Bloqueio_Atraso = SUCESSO
     
    Exit Function
    
Verifica_DataLimite_Bloqueio_Atraso:

    Verifica_DataLimite_Bloqueio_Atraso = gErr
     
    Select Case gErr
                      
        Case 99306
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152625)
     
    End Select
     
    Exit Function
    
End Function

'Função criada por Leo em 29/04/02. Verificar se a função está em seu local de destino correto.
Function PedidoVenda_Desvincula_Orcamento(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Desvincula o Orçamento de Venda do Pedido de Venda a ser excluido.

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_PedidoVenda_Desvincula_Orcamento
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94479
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 94480
    
    'Tenta buscar um Orçamento Vinculado ao Pedido de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM OrcamentoVenda WHERE NumIntPedVenda = ?", 0, lNumIntDoc, objPedVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 94481

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94482
    
    If lErro = SUCESSO Then
    
        'Desvincula o Pedido da Tabela de Orçamentos.
        lErro = Comando_ExecutarPos(lComando1, "UPDATE OrcamentoVenda SET NumIntPedVenda = ?", lComando, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 94483

    End If

    PedidoVenda_Desvincula_Orcamento = SUCESSO

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

Erro_PedidoVenda_Desvincula_Orcamento:

    PedidoVenda_Desvincula_Orcamento = gErr

    Select Case gErr

        Case 94479, 94480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 94481, 94482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)

        Case 94483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOVENDA", gErr, objPedVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152626)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

'Rotinas inseridas por Daniel em 15/05/02
Public Function ItemOV_ExcluiTributacao(ByVal objItemOV As ClassItemOV, ByVal lComando As Long, ByVal lComando1 As Long) As Long

Dim lErro As Long
Dim iNaturezaOpManual As Integer

On Error GoTo Erro_ItemOV_ExcluiTributacao

'    lErro = Comando_ExecutarPos(lComando, "SELECT NaturezaOpManual FROM TributacaoItemOV WHERE NumIntDoc = ? ", 0, iNaturezaOpManual, objItemOV.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 101123
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101124
'    If lErro <> AD_SQL_SUCESSO Then gError 101125
'
'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoItemOV", lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 101126

    lErro = CF("TributacaoDocItem_Exclui", TIPODOC_TRIB_OV, objItemOV.lNumIntDoc)
    If lErro <> SUCESSO Then gError 101126

    ItemOV_ExcluiTributacao = SUCESSO

    Exit Function

Erro_ItemOV_ExcluiTributacao:

    ItemOV_ExcluiTributacao = gErr

    Select Case gErr

'        Case 101123, 101124
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMORCAMENTO", gErr)
'
'        Case 101125
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBITEMORCAMENTO_NAO_CADASTRADO", gErr)

        Case 101126
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMORCAMENTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152627)

    End Select

    Exit Function

End Function

Public Function ItemOV_GravaTributacao(ByVal objItemOV As ClassItemOV, ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ItemOV_GravaTributacao

'    With objItemOV.objTributacaoItemOV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoItemOV (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito  ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objItemOV.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 101127

    objItemOV.objTributacao.lNumIntDoc = objItemOV.lNumIntDocOV
    objItemOV.objTributacao.lNumIntDocItem = objItemOV.lNumIntDoc
    objItemOV.objTributacao.iItem = objItemOV.iItem
    objItemOV.objTributacao.dDescontoGrid = objItemOV.dValorDesconto
    objItemOV.objTributacao.dPrecoUnitario = objItemOV.dPrecoUnitario
    'objItemOV.objTributacao.dPrecoUnitarioMoeda = objItemOV.dPrecoUnitario
    objItemOV.objTributacao.dQuantidade = objItemOV.dQuantidade
    objItemOV.objTributacao.sProduto = objItemOV.sProduto
    objItemOV.objTributacao.sProdutoDescricao = objItemOV.sDescricao
    objItemOV.objTributacao.sUnidadeMed = objItemOV.sUnidadeMed
    
    If objItemOV.iHistorico <> MARCADO Then
        objItemOV.objTributacao.iTipoDoc = TIPODOC_TRIB_OV
    Else
        objItemOV.objTributacao.iTipoDoc = TIPODOC_TRIB_OVHIST
    End If

    lErro = CF("TributacaoDocItem_Grava", objItemOV.objTributacao)
    If lErro <> SUCESSO Then Error 41540

    ItemOV_GravaTributacao = SUCESSO

    Exit Function

Erro_ItemOV_GravaTributacao:

    ItemOV_GravaTributacao = gErr

    Select Case gErr

        Case 101127
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMOV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152628)

    End Select

    Exit Function

End Function

Public Function ItensOV_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui os itens presentes na coleção de itens do Pedido de vanda passado

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objItemOV As ClassItemOV
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_ItensOV_Exclui

    lComando = Comando_Abrir
    If lComando = 0 Then gError 101128
    
    lComando1 = Comando_Abrir
    If lComando1 = 0 Then gError 101129

    'Para cada item na coleção
    For Each objItemOV In objOrcamentoVenda.colItens

        'Busca o ítem no BD
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ItensOrcamentoVenda WHERE NumIntDoc = ?", 0, lNumIntDoc, objItemOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 101130

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101131
        If lErro = AD_SQL_SEM_DADOS Then gError 101132

        'Exclui o item da tabela de itens de Orcamento de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ItensOrcamentoVenda", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 101133
        
        lErro = ItensOVGrade_Exclui(objItemOV)
        If lErro <> SUCESSO Then gError 86363
        
        lErro = ItensOV_Exclui_FormPreco(objItemOV)
        If lErro <> SUCESSO Then gError 86363
        
        objItemOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOV
        objItemOV.objInfoAdicDocItem.lNumIntDocItem = objItemOV.lNumIntDoc

        lErro = CF("InfoAdicDocItem_Exclui", objItemOV.objInfoAdicDocItem)
        If lErro <> SUCESSO Then gError 101134
        
        lErro = CF("ItemOV_ExcluiTributacao", objItemOV, lComando, lComando1)
        If lErro <> SUCESSO Then gError 101134

    Next

    Call Comando_Fechar(lComando)
    
    Call Comando_Fechar(lComando1)

    ItensOV_Exclui = SUCESSO

    Exit Function

Erro_ItensOV_Exclui:

    ItensOV_Exclui = gErr

    Select Case gErr

        Case 86363

        Case 101130, 101131
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr)

        Case 101132
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOV_NAO_ENCONTRADO", gErr, objOrcamentoVenda.lNumIntDoc)

        Case 101133
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case 101134
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case 101128, 101129
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152629)

    End Select

    Call Comando_Fechar(lComando)
    
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function ItensOV_Grava(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Insere os itens do orcamento de Venda passado no BD

Dim alComando(0 To 5) As Long
Dim lErro As Long
Dim objItemOV As ClassItemOV
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOP As New ClassItemOP
Dim sProduto As String
Dim iStatus As Integer

On Error GoTo Erro_ItensOV_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 101135
    Next

    'Para cada item da coleção
    For Each objItemOV In objOrcamentoVenda.colItens

'        sProduto = String(STRING_PRODUTO, 0)
'
'        'Localiza o item no Banco de Dados
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, NumIntDoc FROM ItensOrcamentoVenda WHERE CodOrcamento = ? AND FilialEmpresa = ? AND Produto = ? ", 0, sProduto, lNumIntDoc, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa, objItemOV.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 101136
'
'        lErro = Comando_BuscarPrimeiro(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101137
'        If lErro <> AD_SQL_SUCESSO Then

            'Obtém o número interno automático do item
            If objOrcamentoVenda.iHistorico <> MARCADO Then
                lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ITEM_ORCAMENTOVENDA", lNumIntDoc)
            Else
                lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ITEM_ORCAMENTOVENDAHIST", lNumIntDoc)
            End If
            If lErro <> SUCESSO Then gError 101138

            objItemOV.iFilialEmpresa = giFilialEmpresa
            objItemOV.lCodOrcamento = objOrcamentoVenda.lCodigo
            objItemOV.lNumIntDoc = lNumIntDoc

            If objOrcamentoVenda.iHistorico <> MARCADO Then
                lErro = Comando_Executar(alComando(0), "INSERT INTO ItensOrcamentoVenda (NumIntDoc,FilialEmpresa,CodOrcamento,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Status, MotivoPerda, Observacao, PCSituacao, PCPrecoUnitCalc, PrazoEntrega, Marcado) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItemOV.iFilialEmpresa, objItemOV.lCodOrcamento, objItemOV.sProduto, objItemOV.dQuantidade, objItemOV.dPrecoUnitario, objItemOV.dPrecoTotal, objItemOV.iClasseUM, objItemOV.sUnidadeMed, objItemOV.dValorDesconto, objItemOV.dtDataEntrega, objItemOV.sDescricao, objItemOV.sVersaoKit, objItemOV.sVersaoKitBase, objItemOV.lStatus, objItemOV.lMotivoPerda, objItemOV.sObservacao, objItemOV.iPCSituacao, objItemOV.dPCPrecoUnitCalc, objItemOV.iPrazoEntrega, objItemOV.iMarcado)
                objItemOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOV
            Else
                lErro = Comando_Executar(alComando(0), "INSERT INTO ItensOrcamentoVendaHist (NumIntDoc,FilialEmpresa,CodOrcamento,Versao,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Status, MotivoPerda, Observacao, PCSituacao, PCPrecoUnitCalc, PrazoEntrega, Marcado) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItemOV.iFilialEmpresa, objItemOV.lCodOrcamento, objOrcamentoVenda.iVersao, objItemOV.sProduto, objItemOV.dQuantidade, objItemOV.dPrecoUnitario, objItemOV.dPrecoTotal, objItemOV.iClasseUM, objItemOV.sUnidadeMed, objItemOV.dValorDesconto, objItemOV.dtDataEntrega, objItemOV.sDescricao, objItemOV.sVersaoKit, objItemOV.sVersaoKitBase, objItemOV.lStatus, objItemOV.lMotivoPerda, objItemOV.sObservacao, objItemOV.iPCSituacao, objItemOV.dPCPrecoUnitCalc, objItemOV.iPrazoEntrega, objItemOV.iMarcado)
                objItemOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOVHIST
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 101139
            
            objItemOV.objInfoAdicDocItem.lNumIntDocItem = objItemOV.lNumIntDoc
            
            lErro = CF("InfoAdicDocItem_Insere", objItemOV.objInfoAdicDocItem, alComando(5))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        
'        'Se for uma alteração
'        Else
'
'            objItemOV.lNumIntDoc = lNumIntDoc
'
'            'Faz a atualização no Banco de Dados
'            '#####################################
'            'ALTERADO POR WAGNER
'            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensOrcamentoVenda SET Produto =?, Quantidade = ?, PrecoUnitario = ?, PrecoTotal = ?, ClasseUM = ?, UnidadeMed = ?, ValorDesconto =?, DataEntrega = ?, Descricao = ?, VersaoKit = ? , VersaoKitBase = ? ", alComando(0), objItemOV.sProduto, objItemOV.dQuantidade, objItemOV.dPrecoUnitario, objItemOV.dPrecoTotal, objItemOV.iClasseUM, objItemOV.sUnidadeMed, objItemOV.dValorDesconto, objItemOV.dtDataEntrega, objItemOV.sDescricao, objItemOV.sVersaoKit, objItemOV.sVersaoKitBase)
'            If lErro <> AD_SQL_SUCESSO Then gError 101140
'            '#####################################
'
'            lErro = CF("ItemOV_ExcluiTributacao", objItemOV, alComando(2), alComando(3))
'            If lErro <> SUCESSO Then gError 101141
'
'            lErro = ItensOVGrade_Exclui(objItemOV)
'            If lErro <> SUCESSO Then gError 86361
'
'        End If


        objItemOV.iHistorico = objOrcamentoVenda.iHistorico
        objItemOV.lNumIntDocOV = objOrcamentoVenda.lNumIntDoc

        lErro = CF("ItemOV_GravaTributacao", objItemOV, alComando(4))
        If lErro <> SUCESSO Then gError 101142
        
        lErro = ItensOVGrade_Inclui(objItemOV)
        If lErro <> SUCESSO Then gError 86362
        
        If objOrcamentoVenda.iHistorico <> MARCADO Then
            
            lErro = ItensOV_Grava_Kit(objItemOV)
            If lErro <> SUCESSO Then gError 129970

            lErro = ItensOV_Grava_FormPreco(objOrcamentoVenda, objItemOV)
            If lErro <> SUCESSO Then gError 129970
        End If
        
        '#####################################
        'Inserido por Wagner 09/12/05
        lErro = CF("ItensOV_Grava_Customizado", objItemOV)
        If lErro <> SUCESSO Then gError 141369
        '#####################################
        
    Next

'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("OVItens_Grava_Customizado", objOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 177046
'    '############################################

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensOV_Grava = SUCESSO

    Exit Function

Erro_ItensOV_Grava:

    ItensOV_Grava = gErr

    Select Case gErr

        Case 86361, 86362, 101138, 129970, 141369, 177046

        Case 101135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101139, 101106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSOV", gErr)

        Case 101136, 101137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr)

        Case 101140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSOV", gErr, objItemOV.lNumIntDoc)
            
        Case 101141
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITENSOV", gErr, objOrcamentoVenda.lCodigo)
            
        Case 101142
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152630)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrcamentoVenda_Altera_BD(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objOrcamentoVenda_BD As ClassOrcamentoVenda, ByVal lComando1 As Long) As Long
'Faz as alterações no Orcamento de Venda
'Passa os dados de ítens de OrcamentoVenda no BD para objOrcamentoVenda_BD

Dim lErro As Long
Dim objItemOV As New ClassItemOV
Dim objBloqueioPV_BD As New ClassBloqueioPV
Dim objBloqueioPV1 As ClassBloqueioPV
Dim bAchouBloqueioBD As Boolean
Dim vbMsgRes As VbMsgBoxResult
Dim objBloqueioPV As New ClassBloqueioPV
Dim iIndice As Integer
Dim iBloqueioEstoqueAdicionado As Integer
Dim dPedidoFaltaReservar As Double
Dim bPossuiBloqueioTotal As Boolean
Dim objMovEstoque As New ClassMovEstoque
Dim dQuantNaoCancelada As Double
Dim iCreditoAprovado As Integer
Dim objItemOV_BD As ClassItemOV
Dim bAdicionaBloqueio As Boolean
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando As Long

On Error GoTo Erro_OrcamentoVenda_Altera_BD

    lComando = Comando_Abrir
    If lComando = 0 Then gError 101167
    
    lComando2 = Comando_Abrir
    If lComando2 = 0 Then gError 101168
    
    lComando3 = Comando_Abrir
    If lComando3 = 0 Then gError 101169
    
    'Inicialização de flags
    bPossuiBloqueioTotal = False

    objOrcamentoVenda_BD.lCodigo = objOrcamentoVenda.lCodigo
    objOrcamentoVenda_BD.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa

    'Lê e faz o "lock" no orçamento
    lErro = CF("OrcamentoVenda_Lock", objOrcamentoVenda_BD, lComando1)
    If lErro <> SUCESSO And lErro <> 101244 Then gError 101170
    If lErro = 101244 Then gError 101171

    objOrcamentoVenda.lNumIntDoc = objOrcamentoVenda_BD.lNumIntDoc

    'Lê e locka os itens do orcamento
    lErro = CF("ItensOV_Le_Lock", objOrcamentoVenda_BD)
    If lErro <> SUCESSO And lErro <> 101151 Then gError 101172
    If lErro = 101151 Then gError 101270

    lErro = ItensOV_Exclui_FPAlteracao(objOrcamentoVenda, objOrcamentoVenda_BD)
    If lErro <> SUCESSO Then gError 101173
    
    lErro = CF("ItensOV_Exclui", objOrcamentoVenda_BD)
    If lErro <> SUCESSO Then gError 101173
    
    lErro = CF("OrcamentoVenda_ExcluiTributacao", objOrcamentoVenda_BD, lComando, lComando1, lComando2, lComando3)
    If lErro <> SUCESSO Then gError 101178
    
    lErro = CF("ItensOV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101174
    
    'Apaga os itens da tabela de parcelas
    lErro = CF("ParcelasOV_Exclui", objOrcamentoVenda_BD)
    If lErro <> SUCESSO Then gError 101175
    
    'grava os novos itens de parcelas no BD
    lErro = CF("ParcelasOV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101176
    
    lErro = Comando_ExecutarPos(lComando2, "UPDATE OrcamentoVenda SET NaturezaOp = ?, NomeFilialCli = ?, NomeCli = ?, Vendedor=?, Vendedor2=?, FilialEmpresa=?, Codigo=?, Cliente=?, Filial=?, CondicaoPagto=?, PercAcrescFinanceiro=?, DataEmissao=?, DataReferencia=?, ValorTotal=?, ValorProdutos=?, ValorFrete=?, ValorDesconto=?, ValorSeguro=?, ValorOutrasDespesas=?, TabelaPreco=?, CobrancaAutomatica=?, NumIntDoc=?, PrazoValidade=?, Usuario = ?, CodigoBase = ?, Status=?, MotivoPerda=?, StatusComercial = ?, Versao = ?, DataUltAlt = ?, HoraUltAlt = ?, ValorDescontoTit = ?,ValorDescontoItens = ?,ValorItens = ?, Contato = ?, Email = ?, " & _
        "FilialEntrega = ?,CodTransportadora = ?,MensagemNota = ?,PedidoCliente = ?,CanalVenda = ?,Placa = ?,PlacaUF = ?,VolumeQuant = ?,VolumeEspecie = ?,VolumeMarca = ?,VolumeNumero = ?,PesoLiq = ?,PesoBruto = ?,FreteRespons = ?,CodTranspRedesp = ?,DetPagFrete = ?,DataEntrega = ?,VolumeTotal = ?,PedidoRepresentante = ?,PrazoEntrega = ?,DataEnvio = ?,DataEnt = ?, Moeda = ?, DataPerda = ?, PrazoTexto=?, DataPrevReceb = ?, DataProxCobr = ?, Idioma = ?, PercParticVend2=?", lComando1, _
        objOrcamentoVenda.sNaturezaOp, objOrcamentoVenda.sNomeFilialCli, objOrcamentoVenda.sNomeCli, objOrcamentoVenda.iVendedor, objOrcamentoVenda.iVendedor2, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial, objOrcamentoVenda.iCondicaoPagto, objOrcamentoVenda.dPercAcrescFinanceiro, objOrcamentoVenda.dtDataEmissao, objOrcamentoVenda.dtDataReferencia, objOrcamentoVenda.dValorTotal, objOrcamentoVenda.dValorProdutos, objOrcamentoVenda.dValorFrete, objOrcamentoVenda.dValorDesconto, objOrcamentoVenda.dValorSeguro, objOrcamentoVenda.dValorOutrasDespesas, objOrcamentoVenda.iTabelaPreco, objOrcamentoVenda.iCobrancaAutomatica, objOrcamentoVenda.lNumIntDoc, objOrcamentoVenda.iPrazoValidade, objOrcamentoVenda.sUsuario, objOrcamentoVenda.lCodigoBase, objOrcamentoVenda.lStatus, objOrcamentoVenda.lMotivoPerda, objOrcamentoVenda.iStatusComercial, objOrcamentoVenda.iVersao, objOrcamentoVenda.dtDataUltAlt, objOrcamentoVenda.dHoraUltAlt, _
        objOrcamentoVenda.dValorDescontoTit, objOrcamentoVenda.dValorDescontoItens, objOrcamentoVenda.dValorItens, objOrcamentoVenda.sContato, objOrcamentoVenda.sEmail, _
        objOrcamentoVenda.iFilialEntrega, objOrcamentoVenda.iCodTransportadora, objOrcamentoVenda.sMensagemNota, objOrcamentoVenda.sPedidoCliente, objOrcamentoVenda.iCanalVenda, objOrcamentoVenda.sPlaca, objOrcamentoVenda.sPlacaUF, objOrcamentoVenda.lVolumeQuant, objOrcamentoVenda.lVolumeEspecie, objOrcamentoVenda.lVolumeMarca, objOrcamentoVenda.sVolumeNumero, objOrcamentoVenda.dPesoLiq, objOrcamentoVenda.dPesoBruto, objOrcamentoVenda.iFreteRespons, objOrcamentoVenda.iCodTranspRedesp, objOrcamentoVenda.iDetPagFrete, objOrcamentoVenda.dtDataEntrega, objOrcamentoVenda.dVolumeTotal, objOrcamentoVenda.sPedidoRepresentante, objOrcamentoVenda.iPrazoEntrega, objOrcamentoVenda.dtDataEnvio, objOrcamentoVenda.iDataEnt, objOrcamentoVenda.iMoeda, objOrcamentoVenda.dtDataPerda, objOrcamentoVenda.sPrazoTexto, objOrcamentoVenda.dtDataPrevReceb, objOrcamentoVenda.dtDataProxCobr, objOrcamentoVenda.iIdioma, objOrcamentoVenda.dPercParticVend2)
    If lErro <> AD_SQL_SUCESSO Then gError 101177
    
'    lErro = CF("OrcamentoVenda_ExcluiTributacao", objOrcamentoVenda_BD, lComando, lComando1, lComando2, lComando3)
'    If lErro <> SUCESSO Then gError 101178
    
    lErro = CF("OrcamentoVenda_GravaTributacao", objOrcamentoVenda, lComando1, lComando2)
    If lErro <> SUCESSO Then gError 101179
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando)
    
    OrcamentoVenda_Altera_BD = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Altera_BD:

    OrcamentoVenda_Altera_BD = gErr

    Select Case gErr

        Case 101170, 101172, 84354 To 84356, 101175, 101176, 101173, 101174

        Case 101177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOVENDA", gErr)
            
        Case 101169, 101168, 101167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 101171, 101270
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)
            
        Case 101178
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBOV", gErr, objOrcamentoVenda.lCodigo)
        
        Case 101179
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBOV", gErr, objOrcamentoVenda.lCodigo)
            
        Case 101180
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITENSOV", gErr, objOrcamentoVenda.lCodigo)
            
        Case 101181
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152631)

    End Select

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcamentoVenda_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui o Orcamento de Venda passado

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objItemOV As ClassItemOV

On Error GoTo Erro_OrcamentoVenda_Exclui

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 101181

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 101182
    Next

    'Faz lock no orcamento
    lErro = CF("OrcamentoVenda_Lock", objOrcamentoVenda, alComando(0))
    If lErro <> SUCESSO And lErro <> 101244 Then gError 101183
    If lErro = 101244 Then gError 101184

    'Lê e faz lock nos itens do Pedido
    lErro = CF("ItensOV_Le_Lock", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101151 Then gError 101185
    If lErro = 101151 Then gError 101271
    
    'Exclui as Parcelas do Orcamento
    lErro = CF("ParcelasOV_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101186

    'Exclui os itens do Pedido
    lErro = CF("ItensOV_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101187

    'Exclui o Orcamento de Venda do BD
    lErro = CF("OrcamentoVenda_Exclui_BD", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101188
    
    'Exclui a Tributacao e seus complementos
    lErro = CF("OrcamentoVenda_ExcluiTributacao", objOrcamentoVenda, alComando(0), alComando(1), alComando(2), alComando(3))
    If lErro <> SUCESSO Then gError 101189

    lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_OV, objOrcamentoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181764
    
    lErro = CF("InfoAdicionais_Exclui", TIPODOC_INFOADIC_OV, objOrcamentoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181764
    
    lErro = OrcamentoVendaHist_Exclui(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101186
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 101190

    lErro = CF("WorkFlow_Trata_Transacao", "FAT", objOrcamentoVenda.objTela, gsUsuario)
    If lErro <> SUCESSO Then gError 178239

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrcamentoVenda_Exclui = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Exclui:

    OrcamentoVenda_Exclui = gErr

    Select Case gErr

        Case 101181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 101182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101183, 101185, 101188, 101186, 101187, 101189, 178239, 181764

        Case 101184, 101271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case 101190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152632)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Public Function OrcamentoVenda_Exclui_BD(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui o Orcamento de Venda da tabela de Orcamento de Vendas

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_OrcamentoVenda_Exclui_BD

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101191

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 101192
    
    'Busca o Orcamento no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM OrcamentoVenda WHERE FilialEmpresa = ? AND Codigo = ? ", 0, lNumIntDoc, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101193

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101194
    If lErro = AD_SQL_SEM_DADOS Then gError 101195

    'Exclui o Orcamento da Tabela de Orcamentos
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM OrcamentoVenda", lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 101196

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    OrcamentoVenda_Exclui_BD = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Exclui_BD:

    OrcamentoVenda_Exclui_BD = gErr

    Select Case gErr
    
        Case 101193, 101194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)

        Case 101195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case 101196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 101191, 101192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152633)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function OrcamentoVenda_ExcluiTributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long

Dim lErro As Long
Dim iItem As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_OrcamentoVenda_ExcluiTributacao

'    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM TributacaoComplOV WHERE FilialEmpresa = ? AND Codorcamento = ? ", 0, iItem, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 101197
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101198
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoComplOV", lComando)
'        If lErro <> AD_SQL_SUCESSO Then gError 101199
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101200
'
'    Loop
'
'    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa FROM TributacaoOV WHERE FilialEmpresa = ? AND Codorcamento = ? ", 0, iFilialEmpresa, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 101201
'
'    lErro = Comando_BuscarPrimeiro(lComando2)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101202
'    If lErro <> AD_SQL_SUCESSO Then gError 101203
'
'    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM TributacaoOV", lComando2)
'    If lErro <> AD_SQL_SUCESSO Then gError 101204

    lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_OV, objOrcamentoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 101204

    OrcamentoVenda_ExcluiTributacao = SUCESSO

    Exit Function

Erro_OrcamentoVenda_ExcluiTributacao:

    OrcamentoVenda_ExcluiTributacao = gErr

    Select Case gErr

'        Case 101197, 101198, 101200
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLORCAMENTO", gErr)
'
'        Case 101199
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLORCAMENTO", gErr, objOrcamentoVenda.lCodigo)
'
'        Case 101201, 101202
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBORCAMENTO", gErr)
'
'        Case 101203
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBORCAMENTO_NAO_ENCONTRADA", gErr, objOrcamentoVenda.lCodigo)
'
        Case 101204
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBORCAMENTO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152634)

    End Select

    Exit Function

End Function

Public Function OrcamentoVenda_Grava(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Grava o Orcamento de Venda

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OrcamentoVenda_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 101205

    'Chama a rotina de gravação em transação
    lErro = CF("OrcamentoVenda_Grava_EmTrans", objOrcamentoVenda) 'Alterado por Jorge Specian
    If lErro <> SUCESSO Then gError 137937
    
    'Faz o commit na transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 101214

    lErro = CF("WorkFlow_Trata_Transacao", "FAT", objOrcamentoVenda.objTela, gsUsuario)
    If lErro <> SUCESSO Then gError 178239
    
    OrcamentoVenda_Grava = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Grava:

    OrcamentoVenda_Grava = gErr

    Select Case gErr

        Case 137937, 178239
            'erro tratado na rotina chamada
        
        Case 101205
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 101214
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152635)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Function OrcamentoVenda_GravaTributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_OrcamentoVenda_GravaTributacao

'    With objOrcamentoVenda.objTributacaoOV
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoOV (FilialEmpresa, CodOrcamento, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'           .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 101215
'
'    'inserir os complementos
'    lErro = CF("OrcamentoVenda_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then gError 101216
'
'    lErro = CF("OrcamentoVenda_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then gError 101217
'
'    lErro = CF("OrcamentoVenda_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then gError 101218
'
'    lErro = CF("OrcamentoVenda_GravaTributacaoCompl", objOrcamentoVenda, objOrcamentoVenda.objTributacaoOV.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then gError 101219

    objOrcamentoVenda.objTributacao.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        objOrcamentoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_OV
    Else
        objOrcamentoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_OVHIST
    End If

    lErro = CF("TributacaoDoc_Grava", objOrcamentoVenda.objTributacao)
    If lErro <> SUCESSO Then gError 101219

    OrcamentoVenda_GravaTributacao = SUCESSO

    Exit Function

Erro_OrcamentoVenda_GravaTributacao:

    OrcamentoVenda_GravaTributacao = gErr

    Select Case gErr

'        Case 101215
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBNFISCAL", gErr)
'
        Case 101216 To 101219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152636)

    End Select

    Exit Function

End Function

'Public Function OrcamentoVenda_GravaTributacaoCompl(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objTribComplOV As ClassTributacaoItemPV, ByVal lComando As Long) As Long
'
'Dim lErro As Long, iItem As Integer
'
'On Error GoTo Erro_OrcamentoVenda_GravaTributacaoCompl
'
'    iItem = objTribComplOV.lNumIntDoc
'
'    With objTribComplOV
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplOV (FilialEmpresa, CodOrcamento, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual , PISCRedito, COFINSCredito, IPICredito, ICMSCredito) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 101220
'
'    OrcamentoVenda_GravaTributacaoCompl = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoVenda_GravaTributacaoCompl:
'
'    OrcamentoVenda_GravaTributacaoCompl = gErr
'
'    Select Case gErr
'
'        Case 101220
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLOV", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152637)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function OrcamentoVenda_Insere_BD(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Insere o novo Orcamento de Venda no Banco de Dados

Dim lErro As Long
Dim dQuantNaoCancelada As Double
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long

On Error GoTo Erro_OrcamentoVenda_Insere_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101221
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 101222
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 101223
        
    'Gera o número interno automático do Orcamento de Venda
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ORCAMENTO_VENDA", lNumIntDoc)
    Else
        lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ORCAMENTO_VENDA_HIST", lNumIntDoc)
    End If
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objOrcamentoVenda.lNumIntDoc = lNumIntDoc

    'Insere o novo Pedido no BD
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        lErro = Comando_Executar(lComando, "INSERT INTO OrcamentoVenda (NaturezaOp, NomeCli, NomeFilialCli, Vendedor, Vendedor2, FilialEmpresa, Codigo, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao, DataReferencia, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, TabelaPreco, CobrancaAutomatica, NumIntDoc, PrazoValidade, Usuario, CodigoBase, Status, MotivoPerda, StatusComercial, Versao, DataUltAlt, HoraUltAlt,ValorDescontoTit,ValorDescontoItens,ValorItens, Contato, Email, " & _
            "FilialEntrega,CodTransportadora,MensagemNota,PedidoCliente,CanalVenda,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,VolumeNumero,PesoLiq,PesoBruto,FreteRespons,CodTranspRedesp,DetPagFrete,DataEntrega,VolumeTotal,PedidoRepresentante,PrazoEntrega,DataEnvio,DataEnt,Moeda,DataPerda,PrazoTexto,DataPrevReceb, DataProxCobr, Idioma,PercParticVend2) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objOrcamentoVenda.sNaturezaOp, objOrcamentoVenda.sNomeCli, objOrcamentoVenda.sNomeFilialCli, objOrcamentoVenda.iVendedor, objOrcamentoVenda.iVendedor2, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial, objOrcamentoVenda.iCondicaoPagto, objOrcamentoVenda.dPercAcrescFinanceiro, objOrcamentoVenda.dtDataEmissao, objOrcamentoVenda.dtDataReferencia, objOrcamentoVenda.dValorTotal, objOrcamentoVenda.dValorProdutos, objOrcamentoVenda.dValorFrete, _
            objOrcamentoVenda.dValorDesconto, objOrcamentoVenda.dValorSeguro, objOrcamentoVenda.dValorOutrasDespesas, objOrcamentoVenda.iTabelaPreco, objOrcamentoVenda.iCobrancaAutomatica, objOrcamentoVenda.lNumIntDoc, objOrcamentoVenda.iPrazoValidade, objOrcamentoVenda.sUsuario, objOrcamentoVenda.lCodigoBase, objOrcamentoVenda.lStatus, objOrcamentoVenda.lMotivoPerda, objOrcamentoVenda.iStatusComercial, objOrcamentoVenda.iVersao, objOrcamentoVenda.dtDataUltAlt, objOrcamentoVenda.dHoraUltAlt, objOrcamentoVenda.dValorDescontoTit, objOrcamentoVenda.dValorDescontoItens, objOrcamentoVenda.dValorItens, objOrcamentoVenda.sContato, objOrcamentoVenda.sEmail, _
            objOrcamentoVenda.iFilialEntrega, objOrcamentoVenda.iCodTransportadora, objOrcamentoVenda.sMensagemNota, objOrcamentoVenda.sPedidoCliente, objOrcamentoVenda.iCanalVenda, objOrcamentoVenda.sPlaca, objOrcamentoVenda.sPlacaUF, objOrcamentoVenda.lVolumeQuant, objOrcamentoVenda.lVolumeEspecie, objOrcamentoVenda.lVolumeMarca, objOrcamentoVenda.sVolumeNumero, objOrcamentoVenda.dPesoLiq, objOrcamentoVenda.dPesoBruto, objOrcamentoVenda.iFreteRespons, objOrcamentoVenda.iCodTranspRedesp, objOrcamentoVenda.iDetPagFrete, objOrcamentoVenda.dtDataEntrega, objOrcamentoVenda.dVolumeTotal, objOrcamentoVenda.sPedidoRepresentante, objOrcamentoVenda.iPrazoEntrega, objOrcamentoVenda.dtDataEnvio, objOrcamentoVenda.iDataEnt, objOrcamentoVenda.iMoeda, objOrcamentoVenda.dtDataPerda, objOrcamentoVenda.sPrazoTexto, objOrcamentoVenda.dtDataPrevReceb, objOrcamentoVenda.dtDataProxCobr, objOrcamentoVenda.iIdioma, objOrcamentoVenda.dPercParticVend2)
    Else
        lErro = Comando_Executar(lComando, "INSERT INTO OrcamentoVendaHist (NaturezaOp, NomeCli, NomeFilialCli, Vendedor, Vendedor2, FilialEmpresa, Codigo, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao, DataReferencia, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, TabelaPreco, CobrancaAutomatica, NumIntDoc, PrazoValidade, Usuario, CodigoBase, Status, MotivoPerda, StatusComercial, Versao, DataUltAlt, HoraUltAlt,ValorDescontoTit,ValorDescontoItens,ValorItens, Contato, Email, " & _
            "FilialEntrega,CodTransportadora,MensagemNota,PedidoCliente,CanalVenda,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,VolumeNumero,PesoLiq,PesoBruto,FreteRespons,CodTranspRedesp,DetPagFrete,DataEntrega,VolumeTotal,PedidoRepresentante,PrazoEntrega,DataEnvio,DataEnt,Moeda,DataPerda,PrazoTexto,DataPrevReceb, DataProxCobr, Idioma,PercParticVend2) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objOrcamentoVenda.sNaturezaOp, objOrcamentoVenda.sNomeCli, objOrcamentoVenda.sNomeFilialCli, objOrcamentoVenda.iVendedor, objOrcamentoVenda.iVendedor2, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial, objOrcamentoVenda.iCondicaoPagto, objOrcamentoVenda.dPercAcrescFinanceiro, objOrcamentoVenda.dtDataEmissao, objOrcamentoVenda.dtDataReferencia, objOrcamentoVenda.dValorTotal, objOrcamentoVenda.dValorProdutos, objOrcamentoVenda.dValorFrete, _
            objOrcamentoVenda.dValorDesconto, objOrcamentoVenda.dValorSeguro, objOrcamentoVenda.dValorOutrasDespesas, objOrcamentoVenda.iTabelaPreco, objOrcamentoVenda.iCobrancaAutomatica, objOrcamentoVenda.lNumIntDoc, objOrcamentoVenda.iPrazoValidade, objOrcamentoVenda.sUsuario, objOrcamentoVenda.lCodigoBase, objOrcamentoVenda.lStatus, objOrcamentoVenda.lMotivoPerda, objOrcamentoVenda.iStatusComercial, objOrcamentoVenda.iVersao, objOrcamentoVenda.dtDataUltAlt, objOrcamentoVenda.dHoraUltAlt, objOrcamentoVenda.dValorDescontoTit, objOrcamentoVenda.dValorDescontoItens, objOrcamentoVenda.dValorItens, objOrcamentoVenda.sContato, objOrcamentoVenda.sEmail, _
            objOrcamentoVenda.iFilialEntrega, objOrcamentoVenda.iCodTransportadora, objOrcamentoVenda.sMensagemNota, objOrcamentoVenda.sPedidoCliente, objOrcamentoVenda.iCanalVenda, objOrcamentoVenda.sPlaca, objOrcamentoVenda.sPlacaUF, objOrcamentoVenda.lVolumeQuant, objOrcamentoVenda.lVolumeEspecie, objOrcamentoVenda.lVolumeMarca, objOrcamentoVenda.sVolumeNumero, objOrcamentoVenda.dPesoLiq, objOrcamentoVenda.dPesoBruto, objOrcamentoVenda.iFreteRespons, objOrcamentoVenda.iCodTranspRedesp, objOrcamentoVenda.iDetPagFrete, objOrcamentoVenda.dtDataEntrega, objOrcamentoVenda.dVolumeTotal, objOrcamentoVenda.sPedidoRepresentante, objOrcamentoVenda.iPrazoEntrega, objOrcamentoVenda.dtDataEnvio, objOrcamentoVenda.iDataEnt, objOrcamentoVenda.iMoeda, objOrcamentoVenda.dtDataPerda, objOrcamentoVenda.sPrazoTexto, objOrcamentoVenda.dtDataPrevReceb, objOrcamentoVenda.dtDataProxCobr, objOrcamentoVenda.iIdioma, objOrcamentoVenda.dPercParticVend2)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 101225

    'grava a tributacao da nota como um todo
    lErro = CF("OrcamentoVenda_GravaTributacao", objOrcamentoVenda, lComando2, lComando3)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Insere as Parcelas do Pedido no BD
    lErro = CF("ParcelasOV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Insere os itens do Pedido no BD
    lErro = CF("ItensOV_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    OrcamentoVenda_Insere_BD = SUCESSO

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

Erro_OrcamentoVenda_Insere_BD:

    OrcamentoVenda_Insere_BD = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 101221, 101222, 101223
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101225
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152638)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Public Function ParcelasOV_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui do BD as Parcelas do Orcamento

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iNumParcela As Integer

On Error GoTo Erro_ParcelasOV_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101255

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 101256

    'Busca as Parcelas do Orcamento de Venda
    lErro = Comando_ExecutarPos(lComando, "SELECT NumParcela FROM ParcelasOV WHERE FilialEmpresa = ? AND OrcamentoVenda = ?", 0, iNumParcela, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101257

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101258

    'Para cada parcela encontrada
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui a parcela da tabela de parcelas de Orcamento de Venda
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasOV", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 101259

        'Busca a Próxima
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101260

    Loop

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasOV_Exclui = SUCESSO

    Exit Function

Erro_ParcelasOV_Exclui:

    ParcelasOV_Exclui = gErr

    Select Case gErr

        Case 101255, 101256
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101257, 101258, 101260
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASOV", gErr, objOrcamentoVenda.lCodigo)

        Case 101259
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASOV", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152639)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function ParcelasOV_Grava(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Insere as no BD Parcelas do orcamento passado

Dim lErro As Long
Dim lComando As Long
Dim objParcelaOV As ClassParcelaOV

On Error GoTo Erro_ParcelasOV_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101261
    
    'Para cada parcela da coleção
    For Each objParcelaOV In objOrcamentoVenda.colParcela

        'Insere em ParcelasOV
        If objOrcamentoVenda.iHistorico <> MARCADO Then
            lErro = Comando_Executar(lComando, "INSERT INTO ParcelasOV (FilialEmpresa,OrcamentoVenda,NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objParcelaOV.iNumParcela, objParcelaOV.dtDataVencimento, objParcelaOV.dValor, objParcelaOV.iDesconto1Codigo, objParcelaOV.dtDesconto1Ate, objParcelaOV.dDesconto1Valor, _
            objParcelaOV.iDesconto2Codigo, objParcelaOV.dtDesconto2Ate, objParcelaOV.dDesconto2Valor, objParcelaOV.iDesconto3Codigo, objParcelaOV.dtDesconto3Ate, objParcelaOV.dDesconto3Valor)
        Else
            lErro = Comando_Executar(lComando, "INSERT INTO ParcelasOVHist (FilialEmpresa,OrcamentoVenda,Versao,NumParcela,DataVencimento,Valor,Desconto1Codigo,Desconto1Ate,Desconto1Valor,Desconto2Codigo,Desconto2Ate,Desconto2Valor,Desconto3Codigo,Desconto3Ate,Desconto3Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iVersao, objParcelaOV.iNumParcela, objParcelaOV.dtDataVencimento, objParcelaOV.dValor, objParcelaOV.iDesconto1Codigo, objParcelaOV.dtDesconto1Ate, objParcelaOV.dDesconto1Valor, _
            objParcelaOV.iDesconto2Codigo, objParcelaOV.dtDesconto2Ate, objParcelaOV.dDesconto2Valor, objParcelaOV.iDesconto3Codigo, objParcelaOV.dtDesconto3Ate, objParcelaOV.dDesconto3Valor)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 101262

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ParcelasOV_Grava = SUCESSO

    Exit Function

Erro_ParcelasOV_Grava:

    ParcelasOV_Grava = gErr

    Select Case gErr

        Case 101261
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 101262
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELASOV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152640)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function ItensOpRel_Prepara1(ByVal objProdutoKitProdutos As ClassProdutoKitProdutos, ByVal colItensOP As Collection, sSiglaEmb As String) As Long
'Le a partir do Produto/Versao todos os componentes do Kit e os retorna na colecao (PERMITINDO duplicidade na colecao)
'A Embalagem do ProdutoRaiz é adicionada como último elemento
'sSiglaEmb retorna a sigla da embalagem padrao do item sendo produzido

Dim lErro As Long
Dim dFator As Double, dFatorRaiz As Double, dFatorItemOP As Double
Dim lComando As Long
Dim dQtdEmb As Double
Dim dQuantidade As Double, dQuantidadeArr As Double
Dim objProduto As New ClassProduto
Dim objProdutoKit As New ClassProdutoKit
Dim objEmbalagem As New ClassEmbalagem, iSeq As Integer
Dim tProdutoKitProdutos As typeProdutoKitProdutos
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim objProdutoKitProdutosTEMP As New ClassProdutoKitProdutos

On Error GoTo Erro_ItensOpRel_Prepara1

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111806
    
    'Faz a leitura do percentual de perda do produto raiz para
    'futuro empenho da embalagem do produto raiz
    objProdutoKit.sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
    objProdutoKit.sVersao = objProdutoKitProdutos.sVersao
    
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 111807
    
    'Se encontrou o kit
    If lErro = SUCESSO Then
    
        'descobre o fator de conversao da UM p/UM de estoque do produto raiz
        lErro = CF("UM_Conversao_Trans", objProdutoKitProdutos.iClasseUM, objProdutoKit.sUnidadeMed, objProdutoKitProdutos.sSiglaUMEstoque, dFatorRaiz)
        If lErro <> SUCESSO Then gError 106640

        'descobre o fator de conversao da UM p/UM de estoque do produto ao qual o item da OP se refere
        lErro = CF("UM_Conversao_Trans", objProdutoKitProdutos.iClasseUM, objProdutoKitProdutos.sUnidadeMed, objProdutoKitProdutos.sSiglaUMEstoque, dFatorItemOP)
        If lErro <> SUCESSO Then gError 106641
        
        'Inicializa as strings do Type
        With tProdutoKitProdutos
        
            .sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
            .sUnidadeMed = String(STRING_UM_SIGLA, 0)
            .sProduto = String(STRING_PRODUTO, 0)
        
            'para obter apenas os nós do nivel 1
            lErro = Comando_Executar(lComando, "SELECT ProdutoKit.Seq, ProdutoKit.Produto, Quantidade, ProdutoKit.UnidadeMed, Composicao, ControleEstoque, SiglaUMEstoque, ClasseUM FROM ProdutoKit, Produtos WHERE ProdutoRaiz = ? AND Versao = ? AND ProdutoKit.Produto = Produtos.Codigo AND ProdutoKit.Nivel = 1 ORDER BY Seq", iSeq, .sProduto, .dQuantidade, .sUnidadeMed, .iComposicao, .iControleEstoque, .sSiglaUMEstoque, .iClasseUM, objProdutoKitProdutos.sProdutoRaiz, objProdutoKitProdutos.sVersao)
            
        End With
        
        If lErro <> AD_SQL_SUCESSO Then gError 111809
    
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111813
        
        'Se nao Encontrou => Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 111814
        
        'Enquanto tiver dados
        Do While lErro <> AD_SQL_SEM_DADOS
            
            'Transfere os Dados para o obj
            With objProdutoKitProdutosTEMP
            
                .dQuantidade = tProdutoKitProdutos.dQuantidade
                .iClasseUM = tProdutoKitProdutos.iClasseUM
                .iComposicao = tProdutoKitProdutos.iComposicao
                .iControleEstoque = tProdutoKitProdutos.iControleEstoque
                .sProduto = tProdutoKitProdutos.sProduto
                .sSiglaUMEstoque = tProdutoKitProdutos.sSiglaUMEstoque
                .sUnidadeMed = tProdutoKitProdutos.sUnidadeMed
                .iSeq = iSeq
                .sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
                .sVersao = objProdutoKitProdutos.sVersao
                                    
            End With
            
            'Faz a Conversao de unidades de medida
            lErro = CF("UM_Conversao_Trans", objProdutoKitProdutosTEMP.iClasseUM, objProdutoKitProdutosTEMP.sUnidadeMed, objProdutoKitProdutosTEMP.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 111810
                    
            objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade * dFator
            
            If objProdutoKitProdutosTEMP.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
    
                objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade * ((objProdutoKitProdutos.dQuantidade * dFatorItemOP) / (objProdutoKit.dQuantidade * dFatorRaiz))
    
            End If
    
            Call CF("Arredonda_Insumos", objProdutoKitProdutosTEMP.sSiglaUMEstoque, objProdutoKitProdutosTEMP.dQuantidade, dQuantidadeArr)
            objProdutoKitProdutosTEMP.dQuantidade = dQuantidadeArr
            
            'Adiciona na colecao
            colItensOP.Add objProdutoKitProdutosTEMP
                
            'Instancia outro obj
            Set objProdutoKitProdutosTEMP = New ClassProdutoKitProdutos
                
            'Busca o Próximo Registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111815
        
        Loop
    
    End If
    
    'Le a Embalagem do Produto Raiz
    objProdutoEmbalagem.sProduto = objProdutoKitProdutos.sProdutoRaiz
        
    lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
    If lErro <> SUCESSO And lErro <> 100000 Then gError 111811
        
    'Se encontrou => Le o Produto Associado a Embalagem Padrao
    If lErro <> 100000 Then
    
        'Le os dados da Embalagem Padrao obtida
        objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
            
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 111812
    
        'Se encontrou os Dados => Adiciona à Colecao
        If lErro <> 82763 Then
    
            objProduto.sCodigo = objEmbalagem.sProduto
                
            'Tenta ler o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 111817
        
            'Ausência de Produto no BD
            If lErro <> 28030 Then
    
                'Instancia um novo obj
                Set objProdutoKitProdutosTEMP = New ClassProdutoKitProdutos
                    
                'Faz a Conversao de unidades de medida
                lErro = CF("UM_Conversao_Trans", objProdutoKitProdutos.iClasseUM, objProdutoKitProdutos.sUnidadeMed, objProdutoEmbalagem.sUMEmbalagem, dFator)
                If lErro <> SUCESSO Then gError 111816
                    
                'Transfere para o objProdutoKitProdutosTEMP os dados da embalagem
                dQuantidade = ((objProdutoKitProdutos.dQuantidade * (1 - objProdutoKit.dPercentualPerda)) * dFator) / objProdutoEmbalagem.dCapacidade
                    
                'Arredonda para cima a divisao de Produto/Embalagem
                objProdutoKitProdutosTEMP.dQuantidade = Arredonda_ParaCima(dQuantidade)
                    
                'Transfere para o objProdutoKitProdutosTEMP os dados da embalagem
                With objProdutoKitProdutosTEMP
                    
                    .iClasseUM = objProduto.iClasseUM
                    .iComposicao = PRODUTOKIT_COMPOSICAO_FIXA
                    .iControleEstoque = objProduto.iControleEstoque
                    .sProduto = objProduto.sCodigo
                    .sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
                    .sSiglaUMEstoque = objProduto.sSiglaUMEstoque
                    
                    'retornando a sigla da embalagem do item em questao...
                    sSiglaEmb = objEmbalagem.sSigla
                    
                    .sUnidadeMed = objProduto.sSiglaUMEstoque
                    .sVersao = objProdutoKitProdutos.sVersao
                    .iSeq = colItensOP.Count + 1
                        
                End With
                    
                'Adiciona na colecao
                colItensOP.Add objProdutoKitProdutosTEMP
                    
            End If
                
        End If
            
    End If
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

    ItensOpRel_Prepara1 = SUCESSO
    
    Exit Function
    
Erro_ItensOpRel_Prepara1:

    ItensOpRel_Prepara1 = gErr
    
    Select Case gErr
    
        Case 111806
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 111809, 111813, 111815
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT_PRODUTOS", gErr)
            
        Case 111814 'Sem Dados
        
        Case 111807, 111810, 111811, 111812, 111816, 111817, 106640, 106641
            
        Case 111808
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRODUTO_RAIZ", gErr, objProdutoKit.sProdutoRaiz)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152641)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ItensOPRel_Prepara(ByVal objOP As ClassOrdemDeProducao, lNumIntRel As Long) As Long
'objOP eh parametro de input que traz a ordem de producao e seus itens...

Dim lTransacao As Long, lErro As Long
Dim objItemOP As ClassItemOP
Dim objProdKitProd As New ClassProdutoKitProdutos
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto

On Error GoTo Erro_ItensOPRel_Prepara

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106637

    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "MATConfig", "NUMINTREL_ITENSOPREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 111822
    
    'percorre a colecao de itens de op
    For Each objItemOP In objOP.colItens
    
        If objItemOP.iPossuiGrade = DESMARCADO Then
    
            'passa a chave para obter os dados dos insumos do item...
            'no caso seria a ordenacao correta( o sequencial em produtokit ) e a
            'sua embalagem
            objProdKitProd.sProdutoRaiz = objItemOP.sProduto
            objProdKitProd.sVersao = objItemOP.sVersao
            objProdKitProd.sUnidadeMed = objItemOP.sSiglaUM
            objProdKitProd.dQuantidade = objItemOP.dQuantidade
                    
            objProduto.sCodigo = objItemOP.sProduto
                    
            lErro = ItensOPRel_Prepara3(objProdKitProd, objProduto, lNumIntRel, objItemOP.lNumIntDoc)
            If lErro <> SUCESSO Then gError 126569
                    
                    
        Else
        
            For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
            
                objProdKitProd.sProdutoRaiz = objItemRomaneioGrade.sProduto
                objProdKitProd.sVersao = objItemRomaneioGrade.sVersao
                objProdKitProd.sUnidadeMed = objItemRomaneioGrade.sUMEstoque
                objProdKitProd.dQuantidade = objItemRomaneioGrade.dQuantidade
                
                objProduto.sCodigo = objItemRomaneioGrade.sProduto
                
                lErro = ItensOPRel_Prepara3(objProdKitProd, objProduto, lNumIntRel, objItemRomaneioGrade.lNumIntDoc)
                If lErro <> SUCESSO Then gError 126570
            
            Next
    
        End If
    
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 106638
    
    ItensOPRel_Prepara = SUCESSO
    
    Exit Function
    
Erro_ItensOPRel_Prepara:

    ItensOPRel_Prepara = gErr
    
    Select Case gErr
        
        Case 111822, 126569, 126570
        
        Case 106637
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 106638
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152642)
            
    End Select
    
    Call Transacao_Rollback
    
End Function

Private Function ItensOPRel_Prepara3(objProdKitProd As ClassProdutoKitProdutos, objProduto As ClassProduto, lNumIntRel As Long, ByVal lNumIntDoc As Long) As Long

Dim colItensProd As New Collection
Dim lComando As Long
Dim sSiglaEmb As String
Dim lErro As Long

On Error GoTo Erro_ItensOPRel_Prepara3

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111820

    'Tenta ler o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 106639
    
    objProdKitProd.iClasseUM = objProduto.iClasseUM
    objProdKitProd.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
    If Trim(objProdKitProd.sUnidadeMed) = "" Then objProdKitProd.sUnidadeMed = objProduto.sSiglaUMEstoque
    
    'preenche colItensProd com os insumos e sSiglaEmb com a sigla da embalagem
    Set colItensProd = New Collection
    lErro = ItensOpRel_Prepara1(objProdKitProd, colItensProd, sSiglaEmb)
    If lErro <> SUCESSO Then gError 111825

    'Gravar os itens de colItensProd em ItensOpRel
    For Each objProdKitProd In colItensProd
    
        lErro = Comando_Executar(lComando, "INSERT INTO ItensOpRel (NumIntRel, NumIntDocItemOP, Seq, Produto, Quantidade, SiglaEmb) VALUES (?,?,?,?,?,?)", lNumIntRel, lNumIntDoc, objProdKitProd.iSeq, objProdKitProd.sProduto, objProdKitProd.dQuantidade, IIf(objProdKitProd.iSeq = colItensProd.Count, sSiglaEmb, ""))
        If lErro <> SUCESSO Then gError 111821
    
    Next

    Call Comando_Fechar(lComando)
    
    ItensOPRel_Prepara3 = SUCESSO
    
    Exit Function
    
Erro_ItensOPRel_Prepara3:

    ItensOPRel_Prepara3 = gErr
    
    Select Case gErr
        
        Case 106639, 111825
        
        Case 111820
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 111821
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSOPREL", gErr, lNumIntRel, lNumIntDoc, objProdKitProd.sProduto, objProdKitProd.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152643)
            
    End Select
    
    Call Comando_Fechar(lComando)

End Function

Public Sub Arredonda_Insumos(ByVal sUnidadeMed As String, ByVal dQuantidadeOrig As Double, dQuantidadeArr As Double)
'Arredonda a quantidade de Insumos. A implementacao padrao é nao arredondar. Esta funcao foi criada apenas p/serem criadas versoes customizadas a serem usadas por CF()

    dQuantidadeArr = dQuantidadeOrig
    
End Sub

Function OrdemDeProducao_Grava0(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ou atualiza uma ordem de producao

Dim lErro As Long

On Error GoTo Erro_OrdemDeProducao_Grava0
    
    lErro = CF("OrdemDeProducao_Grava_EmTrans", objOrdemDeProducao)
    If lErro <> AD_SQL_SUCESSO Then gError 105138
    
    '#####################
    'ALTERADO POR WAGNER
    If objOrdemDeProducao.iGeraOPsArvore <> 0 Then
        lErro = CF("OrdemProducao_Grava_Nodes", objOrdemDeProducao)
        If lErro <> SUCESSO Then gError 129980
    End If
    '#####################

    OrdemDeProducao_Grava0 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava0:

    OrdemDeProducao_Grava0 = gErr

    Select Case gErr

        Case 105138, 129980

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152644)

    End Select

    Exit Function

End Function

Function OrdemDeProducao_Exclui1(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao

Dim lErro As Long

On Error GoTo Erro_OrdemDeProducao_Exclui1

    lErro = CF("OrdemDeProducao_Exclui_EmTrans", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 105141

    OrdemDeProducao_Exclui1 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui1:

    OrdemDeProducao_Exclui1 = gErr

    Select Case gErr

        Case 105141
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152645)

    End Select

    Exit Function

End Function

Public Function ItensOPRel_Prepara2(ByVal iFilialEmpresa As Integer, lNumIntRel As Long, ByVal sProd_I As String, ByVal sProd_F As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal sOP_I As String, ByVal sOP_F As String) As Long
'gera registros auxiliares à execucao do relatorio de distribuicao de materias primas por maquina/reator.
'os registros gerados irao fazer parte da view do relatorio (OrdemProdMaquinas)

Dim lTransacao As Long, lErro As Long, lComando As Long
Dim objItemOP As ClassItemOP, lComando2 As Long
Dim sSiglaEmb As String
Dim objProdKitProd As New ClassProdutoKitProdutos
Dim colItensProd As New Collection
Dim objProduto As New ClassProduto
Dim tItemOP As typeItemOP, iClasseUM As Integer, sSiglaUMEstoque As String

On Error GoTo Erro_ItensOPRel_Prepara2

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111820

    'Abre o comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 111820

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106637

    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "MATConfig", "NUMINTREL_ITENSOPREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 111822
    
    With tItemOP
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sVersao = String(STRING_KIT_VERSAO, 0)
    End With
    
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    
    If sOP_I <> "" And sOP_F <> "" Then
    
        lErro = Comando_Executar(lComando, "SELECT ItensOrdemProducao.NumIntDoc, ItensOrdemProducao.Produto , ItensOrdemProducao.Versao, ItensOrdemProducao.SiglaUM, ItensOrdemProducao.Quantidade, Produtos.ClasseUM, Produtos.SiglaUMEstoque FROM ItensOrdemProducao, Produtos WHERE ItensOrdemProducao.Produto = Produtos.Codigo AND ItensOrdemProducao.FilialEmpresa = ? AND ItensOrdemProducao.Codigo >= ? AND ItensOrdemProducao.Codigo <= ?", _
            tItemOP.lNumIntDoc, tItemOP.sProduto, tItemOP.sVersao, tItemOP.sSiglaUM, tItemOP.dQuantidade, iClasseUM, sSiglaUMEstoque, iFilialEmpresa, sOP_I, sOP_F)
        If lErro <> AD_SQL_SUCESSO Then gError 106748
        
    Else
    
        If dtDataInicial <> DATA_NULA And dtDataInicial <> DATA_NULA Then
            
            lErro = Comando_Executar(lComando, "SELECT ItensOrdemProducao.NumIntDoc, ItensOrdemProducao.Produto , ItensOrdemProducao.Versao, ItensOrdemProducao.SiglaUM, ItensOrdemProducao.Quantidade, Produtos.ClasseUM, Produtos.SiglaUMEstoque FROM OrdensDeProducao, ItensOrdemProducao, Produtos WHERE OrdensDeProducao.FilialEmpresa = ItensOrdemProducao.FilialEmpresa AND OrdensDeProducao.Codigo = ItensOrdemProducao.Codigo AND ItensOrdemProducao.Produto = Produtos.Codigo AND ItensOrdemProducao.FilialEmpresa = ? AND OrdensDeProducao.DataEmissao >= ? AND OrdensDeProducao.DataEmissao <= ?", _
                tItemOP.lNumIntDoc, tItemOP.sProduto, tItemOP.sVersao, tItemOP.sSiglaUM, tItemOP.dQuantidade, iClasseUM, sSiglaUMEstoque, iFilialEmpresa, dtDataInicial, dtDataFinal)
            If lErro <> AD_SQL_SUCESSO Then gError 106749
                
        Else
        
            lErro = Comando_Executar(lComando, "SELECT ItensOrdemProducao.NumIntDoc, ItensOrdemProducao.Produto , ItensOrdemProducao.Versao, ItensOrdemProducao.SiglaUM, ItensOrdemProducao.Quantidade, Produtos.ClasseUM, Produtos.SiglaUMEstoque FROM ItensOrdemProducao, Produtos WHERE ItensOrdemProducao.Produto = Produtos.Codigo AND ItensOrdemProducao.FilialEmpresa = ? AND ItensOrdemProducao.Produto >= ? AND ItensOrdemProducao.Produto <= ?", _
                tItemOP.lNumIntDoc, tItemOP.sProduto, tItemOP.sVersao, tItemOP.sSiglaUM, tItemOP.dQuantidade, iClasseUM, sSiglaUMEstoque, iFilialEmpresa, sProd_I, sProd_F)
            If lErro <> AD_SQL_SUCESSO Then gError 106750
                
        End If
    
    End If
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106751
    
    Do While lErro = AD_SQL_SUCESSO
    
        'passa a chave para obter os dados dos insumos do item...
        'no caso seria a ordenacao correta( o sequencial em produtokit ) e a
        'sua embalagem
        With objProdKitProd
            .sProdutoRaiz = tItemOP.sProduto
            .sVersao = tItemOP.sVersao
            .sUnidadeMed = tItemOP.sSiglaUM
            .dQuantidade = tItemOP.dQuantidade
            .iClasseUM = iClasseUM
            .sSiglaUMEstoque = sSiglaUMEstoque
        End With
        
        'preenche colItensProd com os insumos e sSiglaEmb com a sigla da embalagem
        Set colItensProd = New Collection
        lErro = ItensOpRel_Prepara1(objProdKitProd, colItensProd, sSiglaEmb)
        If lErro <> SUCESSO Then gError 111825
    
        'Gravar os itens de colItensProd em ItensOpRel
        For Each objProdKitProd In colItensProd
        
            lErro = Comando_Executar(lComando2, "INSERT INTO ItensOpRel (NumIntRel, NumIntDocItemOP, Seq, Produto, Quantidade, SiglaEmb) VALUES (?,?,?,?,?,?)", lNumIntRel, tItemOP.lNumIntDoc, objProdKitProd.iSeq, objProdKitProd.sProduto, objProdKitProd.dQuantidade, IIf(objProdKitProd.iSeq = colItensProd.Count, sSiglaEmb, ""))
            If lErro <> SUCESSO Then gError 111821
        
        Next
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106752
    
    Loop

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 106638
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    ItensOPRel_Prepara2 = SUCESSO
    
    Exit Function
    
Erro_ItensOPRel_Prepara2:

    ItensOPRel_Prepara2 = gErr
    
    Select Case gErr
        
        Case 106748 To 106752
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPREL_PREPARA2", Err)
        
        Case 111825, 111822, 106639
        
        Case 106637
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 106638
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 111820
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 111821
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSOPREL", gErr, lNumIntRel, objItemOP.lNumIntDoc, objProdKitProd.sProduto, objProdKitProd.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152646)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

End Function

Function ItensPVGrade_Exclui(objItemPV As ClassItemPedido, Optional ByVal bValidaExclusao As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim sProduto As String, bAchou As Boolean, sSerie As String, lNumNF As Long, dtDataEmi As Date
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ItensPVGrade_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 86333
    Next
    
    sProduto = String(STRING_PRODUTO, 0)

    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Produto FROM ItensPedidoDeVendaGrade WHERE NumIntItemPV = ? ", 0, lNumIntDoc, sProduto, objItemPV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 86334
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86335
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        bAchou = False
        For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
            If objItemRomaneio.sProduto = sProduto Then
                objItemRomaneio.lNumIntDoc = lNumIntDoc
                bAchou = True
            End If
        Next
        'Se não vai regravar o item
        If Not bAchou And bValidaExclusao Then
        
            sSerie = String(STRING_SERIE, 0)
        
            'Verifica se já tem tem uma NF apontando para ele
            lErro = Comando_Executar(alComando(2), "SELECT N.Serie, N.NumNotaFiscal, N.DataEmissao FROM ItensNFiscalGrade AS G, ItensNFiscal AS I, NFiscal AS N WHERE G.NumIntItemNF = I.NumIntDoc AND N.NumIntDoc = I.NumIntNF AND G.NumIntItemPVGrade = ?", sSerie, lNumNF, dtDataEmi, lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209800
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209801
            
            If lErro = AD_SQL_SUCESSO Then gError 209802 'Não pode excluir se estiver faturado
            
        End If
        
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensPedidoDeVendaGrade", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 86336
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86337
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensPVGrade_Exclui = SUCESSO
    
    Exit Function
    
Erro_ItensPVGrade_Exclui:

    ItensPVGrade_Exclui = gErr
    
    Select Case gErr
    
        Case 86333
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86334, 86335, 86337
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVGRADE", gErr, objItemPV.sProduto)
        
        Case 86336
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPVGRADE", gErr)
            
        Case 209800, 209801
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFGRADE", gErr, objItemPV.sProduto)
        
        Case 209802
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPVGRADE_VINCULADO_ITEMNFGRADE", gErr, sProduto, sSerie, lNumNF, dtDataEmi)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152647)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ItensPVGrade_Inclui(objItemPV As ClassItemPedido) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim lNumIntDoc As Long

On Error GoTo Erro_ItensPVGrade_Inclui

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86340
       
    For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
        
        If objItemRomaneio.lNumIntDoc = 0 Then
        
            'descobre o proximo NumIntDoc de ItemPVGrade disponivel
            lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_ITEMPVGRADE", lNumIntDoc)
            If lErro <> SUCESSO Then gError 86341
            
            objItemRomaneio.lNumIntDoc = lNumIntDoc
            
        End If
        
        lErro = Comando_Executar(lComando, "INSERT INTO ItensPedidoDeVendaGrade (NumIntDoc,NumIntItemPV,Produto,Quantidade,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC) VALUES (?,?,?,?,?,?,?,?,?)", objItemRomaneio.lNumIntDoc, objItemPV.lNumIntDoc, objItemRomaneio.sProduto, objItemRomaneio.dQuantidade, objItemRomaneio.dQuantCancelada, objItemRomaneio.dQuantReservada, objItemRomaneio.dQuantFaturada, objItemRomaneio.dQuantOP, objItemRomaneio.dQuantSC)
        If lErro <> AD_SQL_SUCESSO Then gError 86342
        
    Next
    
    Call Comando_Fechar(lComando)

    ItensPVGrade_Inclui = SUCESSO
    
    Exit Function

Erro_ItensPVGrade_Inclui:

    ItensPVGrade_Inclui = gErr
    
    Select Case gErr
    
        Case 86340
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86341
        
        Case 86342
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEMPVGRADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152648)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
        
End Function

Function ItensOVGrade_Inclui(objItemOV As ClassItemOV) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim lNumIntDoc As Long

On Error GoTo Erro_ItensOVGrade_Inclui

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86340
    
    For Each objItemRomaneio In objItemOV.colItensRomaneioGrade
        
        'descobre o proximo NumIntDoc de ItemPVGrade disponivel
        If objItemOV.iHistorico <> MARCADO Then
            lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_ITEMOVGRADE", lNumIntDoc)
        Else
            lErro = CF("Config_ObterNumInt", "MATConfig", "NUM_PROX_INT_ITEMOVGRADEHIST", lNumIntDoc)
        End If
        If lErro <> SUCESSO Then gError 86341
        
        objItemRomaneio.lNumIntDoc = lNumIntDoc
        
        If objItemOV.iHistorico <> MARCADO Then
            lErro = Comando_Executar(lComando, "INSERT INTO ItensOrcamentoVendaGrade (NumIntDoc,NumIntItemOV,Produto,Quantidade) VALUES (?,?,?,?)", lNumIntDoc, objItemOV.lNumIntDoc, objItemRomaneio.sProduto, objItemRomaneio.dQuantidade)
        Else
            lErro = Comando_Executar(lComando, "INSERT INTO ItensOrcamentoVendaGradeHist (NumIntDoc,NumIntItemOV,Produto,Quantidade) VALUES (?,?,?,?)", lNumIntDoc, objItemOV.lNumIntDoc, objItemRomaneio.sProduto, objItemRomaneio.dQuantidade)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 86342
        
    Next
    
    Call Comando_Fechar(lComando)

    ItensOVGrade_Inclui = SUCESSO
    
    Exit Function

Erro_ItensOVGrade_Inclui:

    ItensOVGrade_Inclui = gErr
    
    Select Case gErr
    
        Case 86340
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 86341
        
        Case 86342
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEMOVGRADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152649)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
        
End Function

Function ItensOVGrade_Exclui(objItemOV As ClassItemOV) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_ItensOVGrade_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 86333
    Next

    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensOrcamentoVendaGrade WHERE NumIntItemOV = ? ", 0, lNumIntDoc, objItemOV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 86334
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86335
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensOrcamentoVendaGrade", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 86336
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86337
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensOVGrade_Exclui = SUCESSO
    
    Exit Function
    
Erro_ItensOVGrade_Exclui:

    ItensOVGrade_Exclui = gErr
    
    Select Case gErr
    
        Case 86333
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86334, 86335, 86337
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVGRADE", gErr, objItemOV.sProduto)
        
        Case 86336
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPVGRADE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152650)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function PedidoVenda_AtualizaObsBloq(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long, lTransacao As Long

On Error GoTo Erro_PedidoVenda_AtualizaObsBloq

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 125283

    'Realiza a Gravação do BloqueioPV
    lErro = PedidoVenda_AtualizaObsBloq_Trans(objPedidoVenda)
    If lErro <> SUCESSO Then gError 125278
    
    'Faz o Commit da Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125284

    PedidoVenda_AtualizaObsBloq = SUCESSO
     
    Exit Function
    
Erro_PedidoVenda_AtualizaObsBloq:

    PedidoVenda_AtualizaObsBloq = gErr
     
    Select Case gErr
          
        Case 125278
        
        Case 125283
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 125284
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152651)
     
    End Select
     
    Call Transacao_Rollback
    
    Exit Function

End Function

Function PedidoVenda_AtualizaObsBloq_Trans(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'atualiza observacoes de bloqueios que já estavam no BD

Dim lErro As Long, sObservacao As String, iIndice As Integer
Dim objBloqueioPV As New ClassBloqueioPV, alComando(1 To 2) As Long

On Error GoTo Erro_PedidoVenda_AtualizaObsBloq_Trans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 124132
    Next

    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
    
        If objBloqueioPV.iSequencial = 0 Then Exit For
    
        sObservacao = String(STRING_BLOQUEIOSPV_OBSERVACAO, 0)
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Observacao FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND Sequencial = ?", 0, sObservacao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial)
        If lErro <> AD_SQL_SUCESSO Then gError 124130
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124133
        If lErro <> AD_SQL_SUCESSO Then gError 124134
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE BloqueiosPV SET Observacao = ?", alComando(1), objBloqueioPV.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 124131
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoVenda_AtualizaObsBloq_Trans = SUCESSO
     
    Exit Function
    
Erro_PedidoVenda_AtualizaObsBloq_Trans:

    PedidoVenda_AtualizaObsBloq_Trans = gErr
     
    Select Case gErr
          
        Case 124134
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPV_NAO_ENCONTRADO", gErr)
            
        Case 124132
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 124130, 124133
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOPV_OBS", gErr)
        
        Case 124131
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIOPV_OBS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152652)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function PedidoVenda_BloqAutoAtualiza(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal colBloqBD As Collection) As Long
'atualiza o campo "observacao" de bloqueios do pv recriados, que já estavam no bd

Dim lErro As Long, objBloqueioPV As ClassBloqueioPV, objBloqueioPVBD As ClassBloqueioPV

On Error GoTo Erro_PedidoVenda_BloqAutoAtualiza

    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
    
        If objBloqueioPV.dtDataLib = DATA_NULA Then
        
            For Each objBloqueioPVBD In colBloqBD
            
                If objBloqueioPVBD.dtDataLib = DATA_NULA And objBloqueioPVBD.iTipoDeBloqueio = objBloqueioPV.iTipoDeBloqueio And objBloqueioPVBD.iSequencial = objBloqueioPV.iSequencial Then
                
                    objBloqueioPV.sObservacao = objBloqueioPVBD.sObservacao
                    Exit For
                
                End If
            
            Next
    
        End If
        
    Next
    
    PedidoVenda_BloqAutoAtualiza = SUCESSO
     
    Exit Function
    
Erro_PedidoVenda_BloqAutoAtualiza:

    PedidoVenda_BloqAutoAtualiza = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152653)
     
    End Select
     
    Exit Function

End Function


Public Function BloqueioPagtoAnt_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de pagtos antecipados quando a condicao de pagamento for pagto antecipado

    BloqueioPagtoAnt_Adiciona = SUCESSO
     
End Function

Public Function BloqueioPagtoAnt_Altera(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_BD As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de pagtos antecipados quando a condicao de pagamento for pagto antecipado

    BloqueioPagtoAnt_Altera = SUCESSO
     
End Function


Public Function Inicializa_Campos_Obrigatorios(objUserControl As Object) As Long

    Inicializa_Campos_Obrigatorios = SUCESSO
    
End Function

Public Function Orcamento_Form_Load(objUserControl As Object, objContainer As Object) As Long

    Orcamento_Form_Load = SUCESSO
    
End Function

Public Sub Orcamento_Inicializa_Grid_Itens1(ByVal objGridInt As AdmGrid)

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens2(ByVal objGridInt As AdmGrid, ByVal objUserControl As Object)

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens3(iIncremento As Integer)

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens4(ByVal objGridInt As AdmGrid)

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens5(ByVal objGridInt As AdmGrid, ByVal objUserControl As Object)

End Sub

Public Sub Orcamento_Inicializa_Grid_Itens6(iIncremento As Integer)

End Sub

Public Function Carrega_Grid_Itens_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal sPrecoDesc As String, ByVal sProduto As String, ByVal iGrid_DescProduto_Col As Integer) As Long

    Carrega_Grid_Itens_PrecoDesc = SUCESSO

End Function

Public Sub PrecoTotal_Calcula_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal sPrecoDesc As String)

End Sub

Public Function PedidoVenda_Altera_Customizacoes(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objPedidoVenda_BD As ClassPedidoDeVenda) As Long

    PedidoVenda_Altera_Customizacoes = SUCESSO

End Function

Function OrdemProducao_Grava_OPSub_ReqCom(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'gera ops de subprodutos e requisicoes de compras na criacao de uma OP

Dim lErro As Long, objCalcNecesProd As New ClassCalcNecesProd, objNecesProdInfo As ClassNecesProdInfo
Dim objOrdemProd As ClassOrdemDeProducao, objItemOP As ClassItemOP
Dim iAlmoxarifadoPadrao As Integer, iItemOP As Integer, iItemRC As Integer, dFatorConversao As Double
Dim objReqCompras As New ClassRequisicaoCompras, objItemRC As ClassItemReqCompras
Dim lCodigo As Long, objProduto As New ClassProduto, dQuantNecesEst As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objNecesProdInfoAux As ClassNecesProdInfo 'Inserido por Wagner
Dim i1 As Integer 'Inserido por Wagner
Dim i2 As Integer 'Inserido por Wagner
Dim colProdutos As New Collection 'Inserido por Wagner
Dim dQuantidade As Double 'Inserido por Wagner
Dim bAchou As Boolean 'Inserido por Wagner

On Error GoTo Erro_OrdemProducao_Grava_OPSub_ReqCom

    For Each objItemOP In objOrdemDeProducao.colItens
    
        If objItemOP.iPossuiGrade = DESMARCADO Then
    
            Set objNecesProdInfo = New ClassNecesProdInfo
    
            With objNecesProdInfo
            
                .sProduto = objItemOP.sProduto
                .sUMNecesInfo = objItemOP.sSiglaUM
                .dQuantNecesInfo = objItemOP.dQuantidade
                .dtData1 = objItemOP.dtDataInicioProd
                .dtData2 = objItemOP.dtDataFimProd
                .lNumIntItemOP = objItemOP.lNumIntDoc
                .sVersao = objItemOP.sVersao
                Set .objItemOP = objItemOP
                
                objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
                
            End With
            
        Else
        
            For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade

                Set objNecesProdInfo = New ClassNecesProdInfo
        
                With objNecesProdInfo
                
                    .sProduto = objItemRomaneioGrade.sProduto
                    .sUMNecesInfo = objItemRomaneioGrade.sUMEstoque
                    .dQuantNecesInfo = objItemRomaneioGrade.dQuantidade
                    .dtData1 = objItemOP.dtDataInicioProd
                    .dtData2 = objItemOP.dtDataFimProd
                    .lNumIntItemOP = objItemOP.lNumIntDoc
                    .sVersao = objItemOP.sVersao
                    objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
                    
                End With

            Next
            
        End If
            
    Next
    
    objCalcNecesProd.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
    
    objCalcNecesProd.iIgnoraEst = objOrdemDeProducao.iIgnoraEst
    
    lErro = CF("Producao_Calcula_Necessidades", objCalcNecesProd)
    If lErro <> SUCESSO Then gError 124224
    
    '#####################################################
    'Inserido por Wagner
    i1 = 0
    
    
    For Each objNecesProdInfo In objCalcNecesProd.colNecesCalcProd
        
        i1 = i1 + 1
        dQuantidade = 0
        
        bAchou = False
        For Each objNecesProdInfoAux In colProdutos
            If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto Then bAchou = True
        Next
        
        'Se o produto já foi adiconado, ignora
        If Not bAchou Then
        
            i2 = 0
        
            For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd
                        
                i2 = i2 + 1
                
                'Se tiver repetição de produto, então
                If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto And i1 <> i2 Then
                    
                    'Se eles tiverem versões diferentes => Erro
                    If objNecesProdInfoAux.sVersao <> objNecesProdInfo.sVersao Then gError 131331
                                       
                    'Acumula total do produto
                    'Ele volta na Um de Estoque, por isso não precisa converter
                    dQuantidade = objNecesProdInfoAux.dQuantNecesCalcEst + dQuantidade
            
                End If
            Next
            
            objNecesProdInfo.dQuantNecesCalcEst = objNecesProdInfo.dQuantNecesCalcEst + dQuantidade
        
            colProdutos.Add objNecesProdInfo
        End If
    Next
    '#############################################################
    
    For Each objNecesProdInfo In colProdutos 'Alterado por Wagner
           
        '#################################
        ' Inserido por Wagner
        'Se existir um mesmo produto com versão diferente => Erro
'        For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd
'            If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto And objNecesProdInfoAux.sVersao <> objNecesProdInfo.sVersao Then gError 131331
'        Next
        '#################################
           
        If objNecesProdInfo.dQuantNecesCalcEst > objNecesProdInfo.dQuantSaldoLivreInicial Then
        
            dQuantNecesEst = CDbl(Format(objNecesProdInfo.dQuantNecesCalcEst - objNecesProdInfo.dQuantSaldoLivreInicial, FORMATO_ESTOQUE))

            'Busca o produto no Banco de Dados
            objProduto.sCodigo = objNecesProdInfo.sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124227
            If lErro <> SUCESSO Then gError 124228
            
            If objOrdemDeProducao.iGeraOP <> 0 And objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL And objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

                If objOrdemProd Is Nothing Then
            
                    Set objOrdemProd = New ClassOrdemDeProducao
                    
                    'Preenche os demais dados da OP
                    objOrdemProd.dtDataEmissao = gdtDataAtual
                    objOrdemProd.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
                    objOrdemProd.iGeraOP = 0
                    objOrdemProd.iGeraReqCompra = 0
                    objOrdemProd.sCodigo = objOrdemDeProducao.sCodOPGerar
                    objOrdemProd.sOPGeradora = objOrdemDeProducao.sCodigo
            
                End If
                
                lErro = CF("AlmoxarifadoPadrao_Le", objOrdemDeProducao.iFilialEmpresa, objNecesProdInfo.sProduto, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 124225
                If iAlmoxarifadoPadrao = 0 Then gError 124226
                
                'Cria um ItemOP novo com o produto
                Set objItemOP = New ClassItemOP
    
                With objItemOP
                
                    iItemOP = iItemOP + 1
                    .iItem = iItemOP
                    .dQuantidade = dQuantNecesEst
                    .dtDataFimProd = objNecesProdInfo.dtData2
                    .dtDataInicioProd = objNecesProdInfo.dtData1
                    .iAlmoxarifado = iAlmoxarifadoPadrao
                    .iDestinacao = ITEMOP_DESTINACAO_ESTOQUE
                    .iFilialEmpresa = objOrdemProd.iFilialEmpresa
                    .sProduto = objNecesProdInfo.sProduto
                    .sSiglaUM = objNecesProdInfo.sUMEstoque
                    .sSiglaUMEstoque = objNecesProdInfo.sUMEstoque
                    .sVersao = objNecesProdInfo.sVersao
                    .sCodigo = objOrdemProd.sCodigo
                    
                    '###############################################
                    'INSERIDO POR WAGNER
                    'TEM QUE GUARDAR O NUMINTDOC DO PAI
                    'SÓ FUNCIONA SE A OP PAI TIVER APENAS UM ITEM
                    'SENÃO ========> NÃO DÁ PARA SABER QUEM É O ITEM PAI
                    If objOrdemDeProducao.colItens.Count = 1 Then
                        .lNumIntDocPai = objOrdemDeProducao.colItens.Item(1).lNumIntDoc
                    End If
                    '###############################################
                    
                End With

                
                objOrdemProd.colItens.Add objItemOP
        
            End If
            
            If objOrdemDeProducao.iGeraReqCompra <> 0 And objNecesProdInfo.iCompras = PRODUTO_COMPRAVEL Then
                
                'Cria um ItemRC novo com o produto
                lErro = CF("AlmoxarifadoPadrao_Le", objOrdemDeProducao.iFilialEmpresa, objNecesProdInfo.sProduto, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 124229
                'If iAlmoxarifadoPadrao = 0 Then gError 124230
                
                If objReqCompras Is Nothing Then Set objReqCompras = New ClassRequisicaoCompras
                
                Set objItemRC = New ClassItemReqCompras

                iItemRC = iItemRC + 1
                objItemRC.iItem = iItemRC
                
                'Descobre o fator de conversao da UM de estoque para a de Compra
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNecesProdInfo.sUMEstoque, objNecesProdInfo.sUMCompra, dFatorConversao)
                If lErro <> SUCESSO Then gError 124231
                
                objItemRC.dQuantidade = CDbl(Format(dQuantNecesEst * dFatorConversao, FORMATO_ESTOQUE))
                
                objItemRC.iAlmoxarifado = iAlmoxarifadoPadrao
                
                objItemRC.sDescProduto = objProduto.sDescricao
                
                objItemRC.sProduto = objNecesProdInfo.sProduto
                objItemRC.sUM = objNecesProdInfo.sUMCompra

                objReqCompras.colItens.Add objItemRC
            
            End If
        
        End If
        
    Next
    
    If Not (objOrdemProd Is Nothing) Then
    
        objOrdemProd.iNumItens = objOrdemProd.colItens.Count
    
        'Grava a nova OP gerada
        lErro = CF("OrdemDeProducao_Grava_EmTrans", objOrdemProd)
        If lErro <> SUCESSO Then gError 124232
                
    End If
    
    If Not (objReqCompras Is Nothing) Then
    
        If objReqCompras.colItens.Count <> 0 Then
        
            'Chama Config_ObterAutomatico
            lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROXIMO_CODIGO_RC", "RequisicaoCompra", "Codigo", lCodigo)
            If lErro <> SUCESSO Then gError 124233
            
            'Preenche os demais dados
            objReqCompras.dtData = gdtDataAtual
            objReqCompras.dtDataBaixa = DATA_NULA
            objReqCompras.dtDataEnvio = DATA_NULA
            objReqCompras.dtDataLimite = DATA_NULA
            objReqCompras.iFilialCompra = objOrdemDeProducao.iFilialEmpresa
            objReqCompras.iFilialDestino = objOrdemDeProducao.iFilialEmpresa
            objReqCompras.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
            objReqCompras.iTipoDestino = TIPO_DESTINO_EMPRESA
            objReqCompras.lCodigo = lCodigo
            
            objReqCompras.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO
            objReqCompras.sOPCodigo = objOrdemDeProducao.sCodigo
            
            '###############################################
            'INSERIDO POR WAGNER
            'TEM QUE GUARDAR O NUMINTDOC DO PAI
            'SÓ FUNCIONA SE A OP PAI TIVER APENAS UM ITEM
            'SENÃO ========> NÃO DÁ PARA SABER QUEM É O ITEM PAI
            If objOrdemDeProducao.colItens.Count = 1 Then
                objReqCompras.lNumIntDocItemOP = objOrdemDeProducao.colItens.Item(1).lNumIntDoc
            End If
            '###############################################
            
            'Grava a nova ReqCompra gerada
            lErro = CF("RequisicaoCompra_Grava_EmTrans", objReqCompras, True)
            If lErro <> SUCESSO Then gError 124234
            
        End If
        
    End If
    
    OrdemProducao_Grava_OPSub_ReqCom = SUCESSO
     
    Exit Function
    
Erro_OrdemProducao_Grava_OPSub_ReqCom:

    OrdemProducao_Grava_OPSub_ReqCom = gErr
     
    Select Case gErr
          
        Case 124224, 124225, 124227, 124229, 124231, 124232, 124233, 124234

        
        Case 124226, 124230
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)
        
        Case 124228
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 131331
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMOP_PRODUTO_REPETIDO", gErr, objOrdemDeProducao.sCodOPGerar, objNecesProdInfoAux.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152654)
     
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Insere(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long) As Long
'insere o pv na tabela

Dim lErro As Long

On Error GoTo Erro_PedidoDeVenda_Insere
            
    With objPedidoVenda
        lErro = Comando_Executar(lComando, "INSERT INTO PedidosDeVenda (FilialEmpresa,Codigo,FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,ValorProdutos,VolumeNumero,FreteRespons, CodTranspRedesp, DetPagFrete, DataEntrega, PedidoRepresentante, DataRefFluxo,VolumeTotal,Moeda,TaxaMoeda,CodExterno,ReqCompra, OBS, Andamento,ValorDescontoTit,ValorDescontoItens,ValorItens, CodigoBase, Parc, Emitente, UsuarioUltAlteracao, DataInclusao, DataAlteracao, HoraInclusao, HoraAlteracao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCliente, objPedidoVenda.iFilial, _
            .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataEmissao, .dtDataReferencia, .sMensagemNota, .sNaturezaOp, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iProxSeqBloqueio, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .lNumIntDoc, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .dValorProdutos, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .dtDataEntrega, .sPedidoRepresentante, .dtDataRefFluxo, .dVolumeTotal, .iMoeda, .dTaxaMoeda, 0, 0, .sOBS, .iAndamento, .dValorDescontoTit, .dValorDescontoItens, .dValorItens, .lCodigoBase, .iParc, IIf(.sEmitente = "", gsUsuario, .sEmitente), .sUsuarioUltAlteracao, IIf(.dtDataInclusao = DATA_NULA, Date, .dtDataInclusao), .dtDataAlteracao, IIf(.dHoraInclusao = 0, CDbl(Time), .dHoraInclusao), .dHoraAlteracao)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 62220
       
    lErro = CF("PV_Grava_Customizado", objPedidoVenda)
    If lErro <> SUCESSO Then gError 180413
      
    PedidoDeVenda_Insere = SUCESSO
        
    Exit Function
   
Erro_PedidoDeVenda_Insere:
    
    PedidoDeVenda_Insere = gErr
    
    Select Case gErr
    
        Case 62220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDVENDA", gErr, objPedidoVenda.lCodigo)
        
        Case 180413
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152655)
            
    End Select

    Exit Function

End Function


'###########################################################
'###########################################################
'INSERIDO POR WAGNER
'ALTERAÇÃO PDL
Function ItensOV_Grava_Kit(objItemOV As ClassItemOV) As Long

Dim lErro As Long
Dim objKit As New ClassKit
Dim colKit As New Collection
Dim colKitAux As New Collection

On Error GoTo Erro_ItensOV_Grava_Kit
    
    If Len(Trim(objItemOV.sVersaoKit)) <> 0 Then
           
        objKit.sProdutoRaiz = objItemOV.sProduto
        objKit.sVersao = objItemOV.sVersaoKitBase
        
        'Le o KitBase
        lErro = CF("Kit_Le", objKit)
        If lErro <> SUCESSO And lErro <> 21826 Then gError 129972
    
        'Se o kitBase não existe, erro
        If lErro = 21826 Then gError 129973
        
        lErro = CF("Kit_Le_Componentes", objKit)
        If lErro <> SUCESSO And 21831 Then gError 131323
        
        'Le todos os Kits descendentes do Kit Base com seus componentes
        lErro = Kit_Le_Filhos(objKit, colKit, colKitAux)
        If lErro <> SUCESSO Then gError 129974
       
        'Recebe uma coleção de Kits e cria uma nova versão para eles
        lErro = Kit_MudaVersoes(colKit, objItemOV.sVersaoKit)
        If lErro <> SUCESSO Then gError 129975
        
        For Each objKit In colKit
            
'            'Le o KitBase
'            lErro = CF("Kit_Le", objKit)
'            If lErro <> SUCESSO And lErro <> 21826 Then gError 131327
'
'            If lErro = SUCESSO Then gError 131328
            
            'Grava o Kit
            objKit.iSituacao = KIT_SITUACAO_ATIVO
            
            lErro = CF("Kit_Grava_EmTrans", objKit)
            If lErro <> SUCESSO Then gError 129976
        Next
        
    End If

    ItensOV_Grava_Kit = SUCESSO
    
    Exit Function

Erro_ItensOV_Grava_Kit:

    ItensOV_Grava_Kit = gErr
    
    Select Case gErr
    
        Case 129972, 129974, 129975, 129976, 131323, 131327
        
        Case 129973
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_INEXISTENTE", gErr, objKit.sVersao, objKit.sProdutoRaiz)
        
        Case 131328
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_OV_REPETIDO", gErr, objKit.sProdutoRaiz, objKit.sVersao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152656)
            
    End Select

    Exit Function
        
End Function

Function Kit_MudaVersoes(ByVal colKit As Collection, sVersaoNova As String) As Long

Dim lErro As Long
Dim objKitBase1 As ClassKit
Dim objKitBase2 As ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim sVersaoKit As String
Dim objKit As ClassKit

On Error GoTo Erro_Kit_MudaVersoes
    
    'Para cada Kit
    For Each objKitBase1 In colKit

        'Verifica se já existe algum kit com produto igual e retorna a versão que vai ser usada
        lErro = Kit_Retorna_VersaoValidada(objKitBase1, colKit, sVersaoNova, sVersaoKit)
        If lErro <> SUCESSO Then gError 131330
        
        'Percorre os outros kits alterando as versão dos componentes
        For Each objKitBase2 In colKit
'            If objKitBase1.sVersao <> objKitBase2.sVersao Or objKitBase1.sProdutoRaiz <> objKitBase2.sProdutoRaiz Then
                For Each objProdutoKit In objKitBase2.colComponentes
                    
                    '###################################################
                    'ACERTO -- Se a Versão está em branco Então considera a versão padrão
                    If Len(Trim(objProdutoKit.sVersaoKitComp)) = 0 Then
                        Set objKit = New ClassKit
                        objKit.sProdutoRaiz = objProdutoKit.sProduto
        
                        lErro = CF("Kit_Le_Padrao", objKit)
                        If lErro <> SUCESSO And lErro <> 106304 Then gError 131457
                    
                        objProdutoKit.sVersaoKitComp = objKit.sVersao
                    End If
                    '###################################################
                    
                    'Se existe algum componente se referenciando a esse kit => Altera Versao
                    If objKitBase1.sVersao = objProdutoKit.sVersaoKitComp And objKitBase1.sProdutoRaiz = objProdutoKit.sProduto Then
                        objProdutoKit.sVersaoKitComp = sVersaoKit
                    End If
                Next
'            End If
        Next
        
        'Percorre os componentes alterado as versões do Kit
        For Each objProdutoKit In objKitBase1.colComponentes
            objProdutoKit.sVersao = sVersaoKit
        Next
        
        'Altera a versão do próprio Kit
        objKitBase1.sVersao = sVersaoKit
    
    Next

    Kit_MudaVersoes = SUCESSO
    
    Exit Function

Erro_Kit_MudaVersoes:

    Kit_MudaVersoes = gErr
    
    Select Case gErr
    
        Case 131330, 131457
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152657)
            
    End Select

    Exit Function
        
End Function

Function Kit_Le_Filhos(ByVal objKit As ClassKit, ByVal colKit As Collection, ByVal colKitAux As Collection) As Long

Dim lErro As Long
Dim objProdutoKit As ClassProdutoKit
Dim objKitFilho As ClassKit
Dim objKitAux As ClassKit
Dim bJaTem As Boolean

On Error GoTo Erro_Kit_Le_Filhos
    
    'Adiciona do Pai para o Filho
    colKitAux.Add objKit

    For Each objProdutoKit In objKit.colComponentes
       
        Set objKitFilho = New ClassKit
    
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
            
            objKitFilho.sVersao = objProdutoKit.sVersaoKitComp
            objKitFilho.sProdutoRaiz = objProdutoKit.sProduto
            
            lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKitFilho)
            If lErro <> SUCESSO And lErro <> 103203 Then gError 131321
                
            bJaTem = False
            
            For Each objKitAux In colKitAux
                If objKitFilho.sProdutoRaiz = objKitAux.sProdutoRaiz And objKitFilho.sVersao = objKitAux.sVersao Then
                    bJaTem = True
                    Exit For
                End If
            Next
                
            If Not bJaTem Then
                
                'Se o componente é um Kit => Desce na recursividade
                If lErro = SUCESSO Then
                    lErro = Kit_Le_Filhos(objKitFilho, colKit, colKitAux)
                    If lErro <> SUCESSO Then gError 131322
                End If
                
            End If
                
        End If
    
    Next
    
    'Adiciona do Filho para o Pai
    colKit.Add objKit

    Kit_Le_Filhos = SUCESSO
    
    Exit Function

Erro_Kit_Le_Filhos:

    Kit_Le_Filhos = gErr
    
    Select Case gErr
    
        Case 131321, 131322
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152658)
            
    End Select

    Exit Function
        
End Function


Function Kit_Retorna_VersaoValidada(ByVal objKit As ClassKit, ByVal colKit As Collection, ByVal sVersaoBase As String, sVersaoNova As String) As Long

Dim lErro As Long
Dim objKitBase As New ClassKit
Dim iMax As Integer
Dim iPos As Integer
Dim iPosAnt As Integer
Dim bAchouSeq As Boolean
Dim bAchouProduto As Boolean
Dim sParteVersao As String
Dim sParteSeq As String

On Error GoTo Erro_Kit_Retorna_VersaoValidada

    bAchouSeq = False
    bAchouProduto = False

    For Each objKitBase In colKit
    
        'Se existe um Kit de um mesmo produto e com versão diferente => Calcula nova versão
        If objKitBase.sProdutoRaiz = objKit.sProdutoRaiz And objKitBase.sVersao <> objKit.sVersao Then
        
            'Achou o produto repetido
            bAchouProduto = True
        
            'Tenta localizar o último seqüencial da versão
            iPosAnt = 0
            iPos = InStr(1, objKitBase.sVersao, SEPARADOR)
            Do While iPos <> 0
                iPosAnt = iPos
                iPos = InStr(iPos + 1, objKitBase.sVersao, SEPARADOR)
            Loop
            
            If iPosAnt <> 0 Then
                sParteVersao = left(objKitBase.sVersao, iPosAnt - 1)
                sParteSeq = Mid(objKitBase.sVersao, iPosAnt + 1, Len(objKitBase.sVersao) - iPosAnt)
            End If
            
            If iPosAnt <> 0 And sVersaoBase = sParteVersao Then
                bAchouSeq = True
                If iMax < StrParaInt(sParteSeq) Then iMax = StrParaInt(sParteSeq)
            End If
        
        End If
    
    Next
    
    If bAchouProduto Then
        If bAchouSeq Then
            sVersaoNova = sVersaoBase & SEPARADOR & CStr(iMax + 1)
        Else
            sVersaoNova = sVersaoBase & SEPARADOR & "1"
        End If
    Else
        sVersaoNova = sVersaoBase
    End If

    Kit_Retorna_VersaoValidada = SUCESSO
    
    Exit Function

Erro_Kit_Retorna_VersaoValidada:

    Kit_Retorna_VersaoValidada = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152659)
            
    End Select

    Exit Function
        
End Function





'######################################################
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'######################################################
'INSERIDO POR WAGNER

Function OrdemProducao_Grava_Nodes(ByVal objOPRaiz As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objKit As New ClassKit
Dim objNecesProdInfo As ClassNecesProdInfo
Dim objCalcNecesProd As ClassCalcNecesProd
Dim lTransacao As Long
Dim iProxima As Integer
Dim iItemOP As Integer

On Error GoTo Erro_OrdemProducao_Grava_Nodes

    iProxima = 1
    iItemOP = 0

    For Each objItemOP In objOPRaiz.colItens
    
        iItemOP = iItemOP + 1
    
        Set objNecesProdInfo = New ClassNecesProdInfo
        Set objCalcNecesProd = New ClassCalcNecesProd
    
        With objNecesProdInfo
        
            .sProduto = objItemOP.sProduto
            .sUMNecesInfo = objItemOP.sSiglaUM
            .dQuantNecesInfo = objItemOP.dQuantidade
            .dtData1 = objItemOP.dtDataInicioProd
            .dtData2 = objItemOP.dtDataFimProd
            .sVersao = objItemOP.sVersao
            
            Set .objItemOP = objItemOP
                    
        End With
        
        objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
                    
        objCalcNecesProd.iFilialEmpresa = objOPRaiz.iFilialEmpresa
        
        'IMPLEMENTADA PARA CALCULAR TODA NECESSIDADE DA ORDEM DE PRODUCÃO,
        'RETORNANDO TODAS NECESSIDADES MAS COM OS CAMPOS iSeq e iNivel PARA
        'PODER IDENTIFICAR O PRODUTO KIT QUE ORIGINOU A NECESSIDADE
        'TAMBÉM NÃO AGRUPA AS NECESSIDADES POR PRODUTO
        lErro = Producao_Calcula_Necessidades_Nivel(objCalcNecesProd, objOPRaiz, iItemOP)
        If lErro <> SUCESSO Then gError 129984
        
        'GRAVAÇÃO RECURSIVA DOS FILHOS
        lErro = OrdemProducao_Grava_Nodes1(objCalcNecesProd, objOPRaiz, objItemOP, iProxima)
        If lErro <> SUCESSO Then gError 129985
    Next
    
    OrdemProducao_Grava_Nodes = SUCESSO
     
    Exit Function
    
Erro_OrdemProducao_Grava_Nodes:

    OrdemProducao_Grava_Nodes = gErr
     
    Select Case gErr
    
        Case 129984 To 129985
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152660)
     
    End Select
    
    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function OrdemProducao_Grava_Nodes1(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objOrdemDeProducaoPai As ClassOrdemDeProducao, ByVal objItemOPPai As ClassItemOP, iProxima As Integer) As Long
'Gera OPs de subprodutos e requisicoes de compras na criacao de uma OP
'A RECURSIVIDADE ACABA QUANDO NÃO ENCONTRAR MAIS FILHOS PARA OP PASSADA

Dim lErro As Long
Dim objItemOPFilha As ClassItemOP
Dim objOrdemDeProducaoFilha As ClassOrdemDeProducao
Dim objNecesProdInfo As New ClassNecesProdInfo
Dim objKit As New ClassKit
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_OrdemProducao_Grava_Nodes1

    If objItemOPPai Is Nothing Then Exit Function
    If Len(Trim(objItemOPPai.sProduto)) = 0 Then Exit Function

    objKit.sProdutoRaiz = objItemOPPai.sProduto
    objKit.sVersao = objItemOPPai.sVersao

    If objItemOPPai.sVersao = "" Then
        'Lê o Kit
        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 129981
        If lErro = 106304 Then Exit Function
    Else
        objKit.sVersao = objItemOPPai.sVersao
        'Lê o Kit
        lErro = CF("Kit_Le", objKit)
        If lErro <> SUCESSO And lErro <> 21826 Then gError 129982
        If lErro = 21826 Then Exit Function
    End If
        
    lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
    If lErro <> SUCESSO And lErro <> 103203 Then gError 129983
    
    For Each objProdutoKit In objKit.colComponentes
    
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
        
            Set objItemOPFilha = New ClassItemOP
            Set objOrdemDeProducaoFilha = New ClassOrdemDeProducao
            
            objItemOPFilha.sProduto = objProdutoKit.sProduto
            objItemOPFilha.sVersao = objProdutoKit.sVersaoKitComp
                        
            'CALCULA AS NECESSIDADES DO PRODUTOKIT
            lErro = OrdemProducao_CalculaNecessidades(objProdutoKit, objCalcNecesProd, objNecesProdInfo)
            If lErro <> SUCESSO Then gError 129987
            
            'PREENCHE A OP FILHO e RC FILHO E GRAVA
            lErro = OrdemProducao_Grava_OP_RC(objOrdemDeProducaoPai, objItemOPPai, objOrdemDeProducaoFilha, objItemOPFilha, objNecesProdInfo, iProxima)
            If lErro Then gError 129988
                
            'RECURSIVIDADE PARA BUSCAR OS DESCENDENTES INDIRETOS
            lErro = OrdemProducao_Grava_Nodes1(objCalcNecesProd, objOrdemDeProducaoFilha, objItemOPFilha, iProxima)
            If lErro Then gError 129989
        
        End If

    Next
    
    OrdemProducao_Grava_Nodes1 = SUCESSO
     
    Exit Function
    
Erro_OrdemProducao_Grava_Nodes1:

    OrdemProducao_Grava_Nodes1 = gErr
     
    Select Case gErr
    
        Case 129981 To 129983
    
        Case 129987 To 129989
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152661)
     
    End Select

    Exit Function

End Function

Function OrdemProducao_Grava_OP_RC(ByVal objOrdemDeProducaoPai As ClassOrdemDeProducao, ByVal objItemOPPai As ClassItemOP, ByVal objOrdemDeProducaoFilha As ClassOrdemDeProducao, ByVal objItemOPFilha As ClassItemOP, ByVal objNecesProdInfo As ClassNecesProdInfo, iProxima As Integer) As Long
'Essa função preenche as informações da OPFilha considerando estar recebendo
'os dados da OPPai, do ItemOPPai e a Necessidade do Produto

Dim lErro As Long
Dim iAlmoxarifadoPadrao As Integer
Dim dQuantNecesEst As Double
Dim objProduto As New ClassProduto
Dim objReqCompras As New ClassRequisicaoCompras
Dim objItemRC As New ClassItemReqCompras
Dim lCodigo As Long
Dim dFatorConversao As Double
Dim sCodigoOPGerada As String, bProduzir As Boolean

On Error GoTo Erro_OrdemProducao_Grava_OP_RC

    objOrdemDeProducaoFilha.sCodigo = objOrdemDeProducaoPai.sCodigo
    objOrdemDeProducaoFilha.iFilialEmpresa = objOrdemDeProducaoPai.iFilialEmpresa
    objOrdemDeProducaoFilha.iGeraReqCompra = objOrdemDeProducaoPai.iGeraReqCompra
    objOrdemDeProducaoFilha.iGeraOP = objOrdemDeProducaoPai.iGeraOP
    objOrdemDeProducaoFilha.iGeraOPsArvore = objOrdemDeProducaoPai.iGeraOPsArvore
    
    objItemOPFilha.lNumIntDoc = objItemOPPai.lNumIntDoc

    If objNecesProdInfo.dQuantNecesCalcEst > objNecesProdInfo.dQuantSaldoLivreInicial Then
    
        dQuantNecesEst = CDbl(Format(objNecesProdInfo.dQuantNecesCalcEst - objNecesProdInfo.dQuantSaldoLivreInicial, FORMATO_ESTOQUE))

        'Busca o produto no Banco de Dados
        objProduto.sCodigo = objNecesProdInfo.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129990
        If lErro <> SUCESSO Then gError 129991
        
        If objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then

            bProduzir = True
            
            If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then
            
                lErro = CF("Produto_ProducaoLocal", objNecesProdInfo.sProduto, objOrdemDeProducaoFilha.iFilialEmpresa)
                If lErro <> 103277 And lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If lErro <> SUCESSO Then bProduzir = False
                
            End If
            
            If bProduzir Then

                'Preenche os demais dados da OP
                sCodigoOPGerada = objOrdemDeProducaoPai.sCodOPGerar & iProxima
                If Len(sCodigoOPGerada) > STRING_OPCODIGO Then
                
                    sCodigoOPGerada = left(objOrdemDeProducaoPai.sCodOPGerar, STRING_OPCODIGO - Len(CStr(iProxima))) & iProxima
                
                End If
                iProxima = iProxima + 1
                
                objOrdemDeProducaoFilha.dtDataEmissao = gdtDataAtual
                objOrdemDeProducaoFilha.iFilialEmpresa = objOrdemDeProducaoPai.iFilialEmpresa
                objOrdemDeProducaoFilha.iGeraReqCompra = objOrdemDeProducaoPai.iGeraReqCompra
                objOrdemDeProducaoFilha.iGeraOP = objOrdemDeProducaoPai.iGeraOP
                objOrdemDeProducaoFilha.iGeraOPsArvore = objOrdemDeProducaoPai.iGeraOPsArvore
                objOrdemDeProducaoFilha.sCodigo = sCodigoOPGerada
                objOrdemDeProducaoFilha.sOPGeradora = objOrdemDeProducaoPai.sCodigo
                objOrdemDeProducaoFilha.sCodOPGerar = objOrdemDeProducaoPai.sCodOPGerar
                
                lErro = CF("OrdemProducao_Le", objOrdemDeProducaoFilha)
                If lErro <> SUCESSO And lErro <> 30368 Then gError 129957
                If lErro = SUCESSO Then gError 129969
                
                lErro = CF("AlmoxarifadoPadrao_Le", objOrdemDeProducaoPai.iFilialEmpresa, objNecesProdInfo.sProduto, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 129992
                If iAlmoxarifadoPadrao = 0 Then gError 129993
            
                objItemOPFilha.iItem = 1
                objItemOPFilha.lNumIntDoc = 0
                objItemOPFilha.dQuantidade = dQuantNecesEst
                objItemOPFilha.dtDataFimProd = objNecesProdInfo.dtData2
                objItemOPFilha.dtDataInicioProd = objNecesProdInfo.dtData1
                objItemOPFilha.iAlmoxarifado = iAlmoxarifadoPadrao
                objItemOPFilha.iDestinacao = ITEMOP_DESTINACAO_ESTOQUE
                objItemOPFilha.iFilialEmpresa = objOrdemDeProducaoFilha.iFilialEmpresa
                objItemOPFilha.sProduto = objNecesProdInfo.sProduto
                objItemOPFilha.sSiglaUM = objNecesProdInfo.sUMEstoque
                objItemOPFilha.sSiglaUMEstoque = objNecesProdInfo.sUMEstoque
                objItemOPFilha.sVersao = objNecesProdInfo.sVersao
                objItemOPFilha.sCodigo = objOrdemDeProducaoFilha.sCodigo
                objItemOPFilha.lNumIntDocPai = objItemOPPai.lNumIntDoc
                objItemOPFilha.sProdutoPai = objItemOPPai.sProdutoPai
                
                objOrdemDeProducaoFilha.colItens.Add objItemOPFilha
            
                objOrdemDeProducaoFilha.iNumItens = objOrdemDeProducaoFilha.colItens.Count
            
                'Grava a nova OP gerada
                lErro = CF("OrdemDeProducao_Grava_EmTrans", objOrdemDeProducaoFilha)
                If lErro <> SUCESSO Then gError 129994
            
            End If
            
        End If
            
        If objNecesProdInfo.iCompras = PRODUTO_COMPRAVEL And objOrdemDeProducaoPai.iGeraReqCompra <> 0 Then
                        
            'Cria um ItemRC novo com o produto
            lErro = CF("AlmoxarifadoPadrao_Le", objOrdemDeProducaoPai.iFilialEmpresa, objNecesProdInfo.sProduto, iAlmoxarifadoPadrao)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 129995
            
            objItemRC.iItem = 1
            
            'Descobre o fator de conversao da UM de estoque para a de Compra
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNecesProdInfo.sUMEstoque, objNecesProdInfo.sUMCompra, dFatorConversao)
            If lErro <> SUCESSO Then gError 129996
            
            objItemRC.dQuantidade = CDbl(Format(dQuantNecesEst * dFatorConversao, FORMATO_ESTOQUE))
            
            objItemRC.iAlmoxarifado = iAlmoxarifadoPadrao
            objItemRC.sDescProduto = objProduto.sDescricao
            objItemRC.sProduto = objNecesProdInfo.sProduto
            objItemRC.sUM = objNecesProdInfo.sUMCompra

            objReqCompras.colItens.Add objItemRC
            
            'Chama Config_ObterAutomatico
            lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROXIMO_CODIGO_RC", "RequisicaoCompra", "Codigo", lCodigo)
            If lErro <> SUCESSO Then gError 129997
            
            'Preenche os demais dados
            objReqCompras.dtData = gdtDataAtual
            objReqCompras.dtDataBaixa = DATA_NULA
            objReqCompras.dtDataEnvio = DATA_NULA
            objReqCompras.dtDataLimite = DATA_NULA
            objReqCompras.iFilialCompra = objOrdemDeProducaoPai.iFilialEmpresa
            objReqCompras.iFilialDestino = objOrdemDeProducaoPai.iFilialEmpresa
            objReqCompras.iFilialEmpresa = objOrdemDeProducaoPai.iFilialEmpresa
            objReqCompras.iTipoDestino = TIPO_DESTINO_EMPRESA
            objReqCompras.lCodigo = lCodigo
            
            objReqCompras.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO
            objReqCompras.sOPCodigo = objOrdemDeProducaoPai.sCodigo
            objReqCompras.lNumIntDocItemOP = objItemOPPai.lNumIntDoc
            
            'Grava a nova ReqCompra gerada
            lErro = CF("RequisicaoCompra_Grava_EmTrans", objReqCompras, True)
            If lErro <> SUCESSO Then gError 129998
        
        End If
        
    End If

    OrdemProducao_Grava_OP_RC = SUCESSO
     
    Exit Function
    
Erro_OrdemProducao_Grava_OP_RC:

    OrdemProducao_Grava_OP_RC = gErr
     
    Select Case gErr
    
        Case 129957, 129990, 129992, ERRO_SEM_MENSAGEM
        
        Case 129969
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOrdemDeProducaoFilha.sCodigo)
        
        Case 129993
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)
                
        Case 129991
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 129994 To 129998
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152662)
     
    End Select

    Exit Function

End Function

Public Function OrdemProducao_CalculaNecessidades(ByVal objProdutoKit As ClassProdutoKit, ByVal objCalcNecesProd As ClassCalcNecesProd, objNecesProdInfo As ClassNecesProdInfo) As Long
'CALCULA A NECESSIDADE DE PRODUÇÃO/COMPRA DE UM PRODUTO DO KIT

Dim lErro As Long
Dim objNecesProdInfoAux As ClassNecesProdInfo
Dim iIndice As Integer

On Error GoTo Erro_OrdemProducao_CalculaNecessidades

    'INICIALIZA DADOS
    objNecesProdInfo.dQuantNecesCalc = 0
    objNecesProdInfo.dQuantNecesCalcEst = 0
    objNecesProdInfo.dQuantNecesInfo = 0

    For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd
    
        If objNecesProdInfoAux.iSeq = objProdutoKit.iSeq And objNecesProdInfoAux.iNivel = objProdutoKit.iNivel And _
            objProdutoKit.sProdutoRaiz = objNecesProdInfoAux.sProdutoRaiz And objProdutoKit.sVersao = objNecesProdInfoAux.sVersaoPai Then
            
            objNecesProdInfo.dQuantNecesCalc = objNecesProdInfoAux.dQuantNecesCalc
            objNecesProdInfo.dQuantNecesCalcEst = objNecesProdInfoAux.dQuantNecesCalcEst
            objNecesProdInfo.dQuantNecesInfo = objNecesProdInfoAux.dQuantNecesInfo
            objNecesProdInfo.dQuantSaldoLivre = objNecesProdInfoAux.dQuantSaldoLivre
            objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfoAux.dQuantSaldoLivreInicial
            objNecesProdInfo.dSaldoDisponivel = objNecesProdInfoAux.dSaldoDisponivel
            objNecesProdInfo.dSaldoEmpenhosDiretos = objNecesProdInfoAux.dSaldoEmpenhosDiretos
            objNecesProdInfo.dSaldoEmpenhosIndiretos = objNecesProdInfoAux.dSaldoEmpenhosIndiretos
            objNecesProdInfo.dSaldoItensPV = objNecesProdInfoAux.dSaldoItensPV
            objNecesProdInfo.dSaldoOPs = objNecesProdInfoAux.dSaldoOPs
            objNecesProdInfo.dSaldoPedCompras = objNecesProdInfoAux.dSaldoPedCompras
            objNecesProdInfo.dSaldoReqCompras = objNecesProdInfoAux.dSaldoReqCompras
            objNecesProdInfo.dSaldoReservado = objNecesProdInfoAux.dSaldoReservado
            objNecesProdInfo.dSaldoTransferencias = objNecesProdInfoAux.dSaldoTransferencias
            objNecesProdInfo.dtData1 = objNecesProdInfoAux.dtData1
            objNecesProdInfo.dtData2 = objNecesProdInfoAux.dtData2
            objNecesProdInfo.iClasseUM = objNecesProdInfoAux.iClasseUM
            objNecesProdInfo.iCompras = objNecesProdInfoAux.iCompras
            objNecesProdInfo.iNivel = objNecesProdInfoAux.iNivel
            objNecesProdInfo.iSeq = objNecesProdInfoAux.iSeq
            objNecesProdInfo.sProduto = objNecesProdInfoAux.sProduto
            objNecesProdInfo.sUMCompra = objNecesProdInfoAux.sUMCompra
            objNecesProdInfo.sUMEstoque = objNecesProdInfoAux.sUMEstoque
            objNecesProdInfo.sUMNecesInfo = objNecesProdInfoAux.sUMNecesInfo
            objNecesProdInfo.sVersao = objNecesProdInfoAux.sVersao
        
            Exit For
            
        End If
    
    Next

    OrdemProducao_CalculaNecessidades = SUCESSO
     
    Exit Function
    
Erro_OrdemProducao_CalculaNecessidades:

    OrdemProducao_CalculaNecessidades = gErr
     
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152663)
     
    End Select

    Exit Function
    
End Function

Function OrdemDeProducao_Grava_Arvore(ByVal colItemOP As Collection, ByVal objOPPai As ClassOrdemDeProducao) As Long
'Lê a Ordem de Produção pelo Código e pela FilialEmpresa

Dim lErro As Long
Dim lTransacao As Long
Dim objItemOP As ClassItemOP
Dim objItemOPBD As ClassItemOP
Dim objOP As New ClassOrdemDeProducao
Dim colOP As New Collection
Dim iCountTela As Integer
Dim iCountBD As Integer
Dim bItemBDNaoTela As Boolean
Dim iIndice As Integer

On Error GoTo Erro_OrdemDeProducao_Grava_Arvore
   
    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 129247
    
    'Le a OP Ancestral
    lErro = CF("OrdemProducao_Le1", objOPPai)
    If lErro <> SUCESSO Then gError 129248
    
    lErro = CF("ItensOrdemProducao_Le", objOPPai)
    If lErro <> SUCESSO Then gError 129949
    
    'Le todas OP descendentes com OP Ancestral
    lErro = CF("OrdemProducao_Le_Filhos", objOPPai, colOP)
    If lErro <> SUCESSO Then gError 129250
    
    'Adiciona a OP ancestral a coleção
    colOP.Add objOPPai
    
    iCountBD = 0
    
    'Para cada OP do BD
    For Each objOP In colOP
                    
        iIndice = 0
                    
       'Para cada item da OP do BD
        For Each objItemOPBD In objOP.colItens
        
            iCountBD = iCountBD + 1
            iIndice = iIndice + 1
            iCountTela = 0
            
            bItemBDNaoTela = True
            
            'Para cada item da Tela
            For Each objItemOP In colItemOP
            
                iCountTela = iCountTela + 1
            
                'Verifica se o Item da OP do BD é igual ao Item da Tela
                If objItemOP.lNumIntDoc = objItemOPBD.lNumIntDoc Then
                    
                    'O Item do BD passa a ser o item da Tela
                    objOP.colItens.Item(iIndice).iSituacao = objItemOP.iSituacao
                    objOP.colItens.Item(iIndice).sJustificativaBloqueio = objItemOP.sJustificativaBloqueio
             
                    bItemBDNaoTela = False
                                    
                End If
            
            Next
            
            'Verifica se existem itens no BD que não estavam na tela
            ' ===> Foi incluído itens no BD
            'If bItemBDNaoTela Then Call Rotina_Aviso(vbOKOnly, "AVISO_ITEM_NAO_TELA")
                    
        Next
        
        'Grava a OP em dentro de transação
        lErro = CF("OrdemDeProducao_Grava0", objOP)
        If lErro <> SUCESSO Then gError 129251
            
    Next
    
    'Verifica se a quantidade de itens da Tela é maior que a do bd
    ' ===> Foi excluido itens do BD
    'If iCountTela > iCountBD Then Call Rotina_Aviso(vbOKOnly, "AVISO_ITEM_NAO_BD")

    'Commit
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 129252

    OrdemDeProducao_Grava_Arvore = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava_Arvore:

    OrdemDeProducao_Grava_Arvore = gErr

    Select Case gErr
    
        Case 129247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 129248 To 129251
        
        Case 129252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152664)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

'##############################################################
'ROTINAS CÓPIAS COM ALTERAÇÕES

Public Function Producao_Calcula_Necessidades_Nivel(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objOP As ClassOrdemDeProducao, ByVal iItemOP As Integer) As Long
'calcula necessidades de producao e/ou compras a partir de necessidades informadas
'CÓPIA DA FUNÇÃO Producao_Calcula_Necessidades COM INCLUSÃO DOS CAMPOS
'iSeq e iNivel NO objNecesProdInfo E NÃO GRUPAMENTO DOS ITENS
'??????? VERIFICAR SE VAI SER REPLICADO PARA FUNÇÃO ORIGINAL?????????


Dim lErro As Long, objNecesProdInfo As ClassNecesProdInfo, dFatorConversao As Double, objProduto As New ClassProduto, dQuantidadeProd As Double
Dim objKit As New ClassKit

On Error GoTo Erro_Producao_Calcula_Necessidades_Nivel

    'percorre a colecao de necessidades informadas
    For Each objNecesProdInfo In objCalcNecesProd.colNecesInfProd
    
        objProduto.sCodigo = objNecesProdInfo.sProduto
        
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
    
        With objNecesProdInfo
            .iCompras = objProduto.iCompras
            .iClasseUM = objProduto.iClasseUM
            .sUMEstoque = objProduto.sSiglaUMEstoque
            .sUMCompra = objProduto.sSiglaUMCompra
        End With
        
        If objNecesProdInfo.sUMNecesInfo = "" Then
            objNecesProdInfo.sUMNecesInfo = objProduto.sSiglaUMEstoque
        End If
        
        'Descobre o fator de conversao da UM da necessidade informada p/UM de estoque do produto (raiz)
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNecesProdInfo.sUMNecesInfo, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 124206

        dQuantidadeProd = objNecesProdInfo.dQuantNecesInfo * dFatorConversao
        objNecesProdInfo.dQuantNecesCalcEst = dQuantidadeProd
        
        If objNecesProdInfo.sVersao = "" Then

            objKit.sProdutoRaiz = objNecesProdInfo.sProduto

            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 124207

            'Se não existe versão padrão p/ o Kit -> Erro.
            If lErro <> SUCESSO Then gError 124281
            
            objNecesProdInfo.sVersao = objKit.sVersao

        End If
        
        If objNecesProdInfo.objItemOP Is Nothing Then
        
            lErro = Producao_Calcula_Necessidades_Nivel2(objCalcNecesProd, objNecesProdInfo, dQuantidadeProd, 0, objOP, iItemOP)
            If lErro <> SUCESSO Then gError 124208
            
        Else
        
            If objNecesProdInfo.objItemOP.colProdutoKitInfo.Count = 0 Then
            
                lErro = Producao_Calcula_Necessidades_Nivel2(objCalcNecesProd, objNecesProdInfo, dQuantidadeProd, 0, objOP, iItemOP)
                If lErro <> SUCESSO Then gError 124208
                
            Else
            
                lErro = Producao_Calcula_Necessidades_Nivel5(objCalcNecesProd, objNecesProdInfo, dQuantidadeProd, 0, objOP, objNecesProdInfo.objItemOP, iItemOP)
                If lErro <> SUCESSO Then gError 124208
                
            End If
        
        End If
            
    Next
    
    'percorre a colecao de necessidades de compra e as coloca na unidade de compras
    For Each objNecesProdInfo In objCalcNecesProd.colNecesCalcProd
    
        objNecesProdInfo.dQuantNecesCalcEst = objNecesProdInfo.dQuantNecesCalc
        
        If objNecesProdInfo.iCompras = PRODUTO_COMPRAVEL Then
        
            'Descobre o fator de conversao da UM de estoque para a de Compra
            lErro = CF("UM_Conversao_Trans", objNecesProdInfo.iClasseUM, objNecesProdInfo.sUMEstoque, objNecesProdInfo.sUMCompra, dFatorConversao)
            If lErro <> SUCESSO Then gError 124215
        
            'armazena a necessidade de compra na unidade de medida de compra
            objNecesProdInfo.dQuantNecesCalc = objNecesProdInfo.dQuantNecesCalcEst * dFatorConversao
                    
        End If
        
    Next
    
    Producao_Calcula_Necessidades_Nivel = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades_Nivel:

    Producao_Calcula_Necessidades_Nivel = gErr
     
    Select Case gErr
          
        Case 124205 To 124208, 124215
        
        Case 124281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152665)
     
    End Select
     
    Exit Function

End Function

Private Function Producao_Calcula_Necessidades_Nivel2(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objNecesProdInfo As ClassNecesProdInfo, ByVal dQuantidadeProd As Double, ByVal iNiveis As Integer, ByVal objOP As ClassOrdemDeProducao, ByVal iItemOP As Integer) As Long

Dim lErro As Long, objNecesProdInfoFilho As ClassNecesProdInfo
Dim objKit As New ClassKit, objNecesProdInfoAux As ClassNecesProdInfo
Dim objProdutoKit As New ClassProdutoKit
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double, bAchou As Boolean
Dim dQuantidadePai As Double, dQuantidadeFilho As Double
Dim dFatorConversaoProdutoRaiz As Double
Dim bAchouPrimeiro As Boolean 'Inserido por Wagner
Dim objNecesProdInfoPrimeiro As ClassNecesProdInfo 'Inserido por Wagner
Dim dQtdSaldoLivre As Double

On Error GoTo Erro_Producao_Calcula_Necessidades_Nivel2

    'Se a quantidade já chegou a 0.1 ou já desceu na recursividade n instâncias
    If (dQuantidadeProd <= 0.1 And iNiveis > 0) Or iNiveis >= 20 Then
    
        Producao_Calcula_Necessidades_Nivel2 = SUCESSO
        Exit Function

    End If

    If iNiveis > 0 Then
    
        bAchou = False
        bAchouPrimeiro = False
        
        'COMENTADO PARA NÃO AGRUPAR ITENS
        For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd

            '############################################
            'Inserido por Wagner
            If Not bAchouPrimeiro And objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto Then
                bAchouPrimeiro = True
                Set objNecesProdInfoPrimeiro = objNecesProdInfoAux
            End If
            '############################################
            
'           If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto then
'
'                bAchou = True
'
'                If objNecesProdInfo.dtData1 <> DATA_NULA Then
'                    If objNecesProdInfoAux.dtData1 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData1 = objNecesProdInfo.dtData1
'                End If
'
'                If objNecesProdInfo.dtData2 <> DATA_NULA Then
'                    If objNecesProdInfoAux.dtData2 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData2 = objNecesProdInfo.dtData2
'                End If
'
'                Exit For
'
'            End If
'
        Next
        
        If bAchou = False Then
            
            If objNecesProdInfo.sVersao = "" And objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then
    
                objKit.sProdutoRaiz = objNecesProdInfo.sProduto
    
                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235
    
                'Se não existe versão padrão p/ o Kit -> Erro.
                If lErro <> SUCESSO Then gError 124282
                
                objNecesProdInfo.sVersao = objKit.sVersao
    
            End If
            
            Set objNecesProdInfoAux = New ClassNecesProdInfo
            
            With objNecesProdInfo
                objNecesProdInfoAux.sProduto = .sProduto
                objNecesProdInfoAux.sVersao = .sVersao
                objNecesProdInfoAux.iCompras = .iCompras
                objNecesProdInfoAux.iClasseUM = .iClasseUM
                objNecesProdInfoAux.sUMNecesInfo = ""
                objNecesProdInfoAux.sUMEstoque = .sUMEstoque
                objNecesProdInfoAux.sUMCompra = .sUMCompra
                objNecesProdInfoAux.dtData1 = .dtData1
                objNecesProdInfoAux.dtData2 = .dtData2
                
                objNecesProdInfoAux.iSeq = .iSeq
                objNecesProdInfoAux.iNivel = .iNivel
                objNecesProdInfoAux.sProdutoRaiz = .sProdutoRaiz
                objNecesProdInfoAux.sVersaoPai = .sVersaoPai
            End With
            
            'obter saldo "livre" de produto
            '############################################
            'Alterado por Wagner
            If Not bAchouPrimeiro Then
                lErro = Producao_Calcula_Necessidades_Nivel3(objCalcNecesProd, objNecesProdInfoAux, objOP, iItemOP)
                If lErro <> SUCESSO Then gError 124209
            End If
            '############################################
            
            'incluir o produto na colecao
            objCalcNecesProd.colNecesCalcProd.Add objNecesProdInfoAux
            
        End If
                
        objNecesProdInfoAux.dQuantNecesCalc = objNecesProdInfoAux.dQuantNecesCalc + dQuantidadeProd
        
        '############################################
        'Alterado por Wagner
        If bAchouPrimeiro Then
            dQtdSaldoLivre = objNecesProdInfoPrimeiro.dQuantSaldoLivre
        Else
            dQtdSaldoLivre = objNecesProdInfoAux.dQuantSaldoLivre
        End If
        '############################################
        
        'se o saldo livre for >= necessario
        If dQtdSaldoLivre >= dQuantidadeProd Then
            
            'abater o necessario do saldo livre
            dQtdSaldoLivre = dQtdSaldoLivre - dQuantidadeProd
            dQuantidadeProd = 0
                    
        Else
                        
            'abater da necessidade o saldo livre
            dQuantidadeProd = dQuantidadeProd - dQtdSaldoLivre
            dQtdSaldoLivre = 0
            
        End If
        
        '############################################
        'Alterado por Wagner/ Corrigido em 30/08/05
        'Se já existe entrada para esse produto
        If bAchouPrimeiro Then
            objNecesProdInfoAux.dQuantSaldoLivreInicial = objNecesProdInfoPrimeiro.dQuantSaldoLivre
            objNecesProdInfoAux.dSaldoEmpenhosDiretos = objNecesProdInfoPrimeiro.dSaldoEmpenhosDiretos + (objNecesProdInfoPrimeiro.dQuantSaldoLivreInicial - objNecesProdInfoPrimeiro.dQuantSaldoLivre)
            objNecesProdInfoPrimeiro.dQuantSaldoLivre = dQtdSaldoLivre
            objNecesProdInfoAux.dSaldoDisponivel = objNecesProdInfoPrimeiro.dSaldoDisponivel
            objNecesProdInfoAux.dQuantSaldoLivre = dQtdSaldoLivre - dQuantidadeProd
        Else
            objNecesProdInfoAux.dQuantSaldoLivre = dQtdSaldoLivre
        End If
        '############################################
        
    End If
    
    If dQuantidadeProd > 0 Then
        
        If objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then
        
            objKit.sVersao = objNecesProdInfo.sVersao
            objKit.sProdutoRaiz = objNecesProdInfo.sProduto
        
            'tem que ler o proprio cara e os seus filhos
            lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
            If lErro <> SUCESSO And lErro <> 103203 Then gError 124210
            
            'Se o produto for um Kit, então leva a recursividade a diante
            If lErro = SUCESSO Then
            
                'percorre os elementos componentes do Kit Lido
                For Each objProdutoKit In objKit.colComponentes
                
                    objProduto.sCodigo = objProdutoKit.sProduto
            
                    'Lê o produto para descobrir as unidades de medidas associadas
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 124211
            
                    'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
                    If lErro <> SUCESSO Then gError 124212
                
                    If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                        
                        dQuantidadePai = (dQuantidadeProd / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                        
                    Else
                        
                        Set objNecesProdInfoFilho = New ClassNecesProdInfo
                
                        '##############################################
                        'Inserido por Wagner 06/10/2005
                        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                            
                            If objProdutoKit.sVersaoKitComp = "" Then
                
                                objKit.sProdutoRaiz = objProdutoKit.sProduto
                
                                lErro = CF("Kit_Le_Padrao", objKit)
                                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235
                
                                'Se não existe versão padrão p/ o Kit -> Erro.
                                If lErro <> SUCESSO Then gError 124282
                
                                objNecesProdInfoFilho.sVersao = objKit.sVersao
                
                            Else
                        
                                objNecesProdInfoFilho.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                
                            End If
                            
                        End If
                        '###############################################
                        
                        With objNecesProdInfoFilho
                            .sProduto = objProdutoKit.sProduto
                            .iCompras = objProduto.iCompras
                            .iClasseUM = objProduto.iClasseUM
                            .sUMEstoque = objProduto.sSiglaUMEstoque
                            .sUMCompra = objProduto.sSiglaUMCompra
                            If objNecesProdInfo.dtData1 <> DATA_NULA Then
                                .dtData2 = objNecesProdInfo.dtData1
                                .dtData1 = .dtData2 - objProduto.iTempoProducao
                            End If
                            
                            .iSeq = objProdutoKit.iSeq
                            .iNivel = objProdutoKit.iNivel
                            .sProdutoRaiz = objProdutoKit.sProdutoRaiz
                            .sVersaoPai = objProdutoKit.sVersao
                        End With
                
                        'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
                        dQuantidadeFilho = dQuantidadePai * (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda) * dFatorConversao)
                          
                        'Calcula de maneira recursiva a quantidade de componentes necessários.
                        lErro = Producao_Calcula_Necessidades_Nivel2(objCalcNecesProd, objNecesProdInfoFilho, dQuantidadeFilho, iNiveis + 1, objOP, iItemOP)
                        If lErro <> SUCESSO Then gError 124213
                
                    End If
                
                Next
            
            End If
    
        End If
        
    End If
    
    Producao_Calcula_Necessidades_Nivel2 = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades_Nivel2:

    Producao_Calcula_Necessidades_Nivel2 = gErr
     
    Select Case gErr
          
        Case 124209 To 124213, 124235
        
        Case 124282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152666)
     
    End Select
     
    Exit Function

End Function

Private Function Producao_Calcula_Necessidades_Nivel3(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objNecesProdInfo As ClassNecesProdInfo, ByVal objOP As ClassOrdemDeProducao, ByVal iItemOP As Integer) As Long
'obtem do bd informacoes sobre saldo de produto

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dEmpenhos As Double, dSaldoEmpenhos As Double
Dim objOPAUx As ClassOrdemDeProducao

On Error GoTo Erro_Producao_Calcula_Necessidades_Nivel3

    '??? falta tratar outros flags referentes a qtde em pedidos de compra, requisicao,..
    
    If (objNecesProdInfo.iCompras = PRODUTO_COMPRAVEL And objCalcNecesProd.iSemSaldoProdComprado = 0) Or (objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL And objCalcNecesProd.iSemSaldoProdProduzido = 0) Then
    
        objEstoqueProduto.sProduto = objNecesProdInfo.sProduto
        
        'Lê as quantidades do produtos nos almoxs da filial
        lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, objCalcNecesProd.iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 52979 Then gError 124214
        
        With objEstoqueProduto
        
            objNecesProdInfo.dQuantSaldoLivreInicial = .dQuantDisponivel + .dQuantRecIndl + .dQuantBenef3 + .dQuantConsig3 '+ .dQuantOP
            If gobjMAT.iOPTipoCalcNecessidade = OP_TIPO_CALC_NECESS_USA_QUANTOP Then
                objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfo.dQuantSaldoLivreInicial + .dQuantOP
            End If
            
            objNecesProdInfo.dSaldoDisponivel = .dQuantDisponivel
            objNecesProdInfo.dSaldoReservado = .dQuantReservada
            
            Set objOPAUx = New ClassOrdemDeProducao
            objOPAUx.colItens.Add objCalcNecesProd.colNecesInfProd.Item(1).objItemOP
            
            'Obtém a quantidade empenhada da própria OP (deve ser desconsiderado)
            lErro = Producao_Calcula_Necessidades_Nivel4(objOP, objNecesProdInfo.sProduto, dEmpenhos, iItemOP)
            If lErro <> SUCESSO Then gError 140453
    
            objNecesProdInfo.dSaldoEmpenhosDiretos = .dQuantEmpenhada - dEmpenhos
    
            'se for p/abater empenhos diretos
            If objCalcNecesProd.iSemSaldoEmpenhosDiretos = 0 Then
                If gobjMAT.iOPTipoCalcNecessidade <> OP_TIPO_CALC_NECESS_QUANTOP_ABATE_EMPENHO Then
                    objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfo.dQuantSaldoLivreInicial - objNecesProdInfo.dSaldoEmpenhosDiretos
                Else
                    'Só considero os empenhos que excedem a produção
                    dSaldoEmpenhos = objNecesProdInfo.dSaldoEmpenhosDiretos - .dQuantOP
                    If dSaldoEmpenhos < QTDE_ESTOQUE_DELTA Then dSaldoEmpenhos = 0 'Tudo que está empenhado já está em produção
                    objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfo.dQuantSaldoLivreInicial - dSaldoEmpenhos
                End If
            End If
            
        End With
    
    Else
    
        objNecesProdInfo.dQuantSaldoLivreInicial = 0
        
    End If
    
    If objNecesProdInfo.dQuantSaldoLivreInicial < 0 Then objNecesProdInfo.dQuantSaldoLivreInicial = 0
    objNecesProdInfo.dQuantSaldoLivre = objNecesProdInfo.dQuantSaldoLivreInicial
    
    Producao_Calcula_Necessidades_Nivel3 = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades_Nivel3:

    Producao_Calcula_Necessidades_Nivel3 = gErr
     
    Select Case gErr
          
        Case 124214, 140453
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152667)
     
    End Select
     
    Exit Function

End Function

Private Function Producao_Calcula_Necessidades_Nivel5(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objNecesProdInfo As ClassNecesProdInfo, ByVal dQuantidadeProd As Double, ByVal iNiveis As Integer, ByVal objOP As ClassOrdemDeProducao, ByVal objItemOP As ClassItemOP, ByVal iItemOP As Integer) As Long

Dim lErro As Long, objNecesProdInfoFilho As ClassNecesProdInfo
Dim objKit As New ClassKit, objNecesProdInfoAux As ClassNecesProdInfo
Dim objProdutoKit As New ClassProdutoKit
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double, bAchou As Boolean
Dim dQuantidadePai As Double, dQuantidadeFilho As Double
Dim dFatorConversaoProdutoRaiz As Double
Dim bAchouPrimeiro As Boolean 'Inserido por Wagner
Dim objNecesProdInfoPrimeiro As ClassNecesProdInfo 'Inserido por Wagner
Dim dQtdSaldoLivre As Double
Dim objProdutoKitInfo As ClassProdutoKitInfo

On Error GoTo Erro_Producao_Calcula_Necessidades_Nivel5

    'Se a quantidade já chegou a 0.1 ou já desceu na recursividade n instâncias
    If (dQuantidadeProd <= 0.1 And iNiveis > 0) Or iNiveis >= 20 Then
    
        Producao_Calcula_Necessidades_Nivel5 = SUCESSO
        Exit Function

    End If

    If iNiveis > 0 Then
    
        '#############################################################
        'Inserido por Wagner 16/08/2006
        'Só inclui se ele tiver sido marcado para gerar requisição
        For Each objProdutoKitInfo In objItemOP.colProdutoKitInfo
        
            If objProdutoKitInfo.objProdutoKit.sProdutoRaiz = objNecesProdInfo.sProdutoRaiz And _
            objProdutoKitInfo.objProdutoKit.sVersao = objNecesProdInfo.sVersaoPai And _
            objProdutoKitInfo.objProdutoKit.iNivel = objNecesProdInfo.iNivel And _
            objProdutoKitInfo.objProdutoKit.iSeq = objNecesProdInfo.iSeq Then
                Exit For
            End If
        Next
        '#############################################################
    
        bAchou = False
        bAchouPrimeiro = False
        
        'COMENTADO PARA NÃO AGRUPAR ITENS
        For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd

            '############################################
            'Inserido por Wagner
            If Not bAchouPrimeiro And objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto Then
                bAchouPrimeiro = True
                Set objNecesProdInfoPrimeiro = objNecesProdInfoAux
            End If
            '############################################
            
'           If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto then
'
'                bAchou = True
'
'                If objNecesProdInfo.dtData1 <> DATA_NULA Then
'                    If objNecesProdInfoAux.dtData1 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData1 = objNecesProdInfo.dtData1
'                End If
'
'                If objNecesProdInfo.dtData2 <> DATA_NULA Then
'                    If objNecesProdInfoAux.dtData2 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData2 = objNecesProdInfo.dtData2
'                End If
'
'                Exit For
'
'            End If
'
        Next
        
        If bAchou = False Then
            
            If objNecesProdInfo.sVersao = "" And objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then
    
                objKit.sProdutoRaiz = objNecesProdInfo.sProduto
    
                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235
    
                'Se não existe versão padrão p/ o Kit -> Erro.
                If lErro <> SUCESSO Then gError 124282
                
                objNecesProdInfo.sVersao = objKit.sVersao
    
            End If
            
            Set objNecesProdInfoAux = New ClassNecesProdInfo
            
            With objNecesProdInfo
                objNecesProdInfoAux.sProduto = .sProduto
                objNecesProdInfoAux.sVersao = .sVersao
                objNecesProdInfoAux.iCompras = .iCompras
                objNecesProdInfoAux.iClasseUM = .iClasseUM
                objNecesProdInfoAux.sUMNecesInfo = ""
                objNecesProdInfoAux.sUMEstoque = .sUMEstoque
                objNecesProdInfoAux.sUMCompra = .sUMCompra
                objNecesProdInfoAux.dtData1 = .dtData1
                objNecesProdInfoAux.dtData2 = .dtData2
                
                objNecesProdInfoAux.iSeq = .iSeq
                objNecesProdInfoAux.iNivel = .iNivel
                objNecesProdInfoAux.sProdutoRaiz = .sProdutoRaiz
                objNecesProdInfoAux.sVersaoPai = .sVersaoPai
            End With
            
            'obter saldo "livre" de produto
            '############################################
            'Alterado por Wagner
            If Not bAchouPrimeiro Then
                lErro = Producao_Calcula_Necessidades_Nivel3(objCalcNecesProd, objNecesProdInfoAux, objOP, iItemOP)
                If lErro <> SUCESSO Then gError 124209
            End If
            '############################################
            
            '#############################################################
            'Alterado por Wagner 16/08/2006
            'Só inclui se ele tiver sido marcado para gerar requisição de compra ou for produzível
            If objProdutoKitInfo.iSelecionado = MARCADO Or objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Then
            'incluir o produto na colecao
                objCalcNecesProd.colNecesCalcProd.Add objNecesProdInfoAux
            End If
            '#############################################################
            
        End If
                
        '#############################################################
        'Alterado por Wagner 16/08/2006
        'Só altera se ele tiver sido marcado para gerar requisição de compra ou for produzível
        If objProdutoKitInfo.iSelecionado = MARCADO Or objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Then
            objNecesProdInfoAux.dQuantNecesCalc = objNecesProdInfoAux.dQuantNecesCalc + dQuantidadeProd
        End If
        '#############################################################
        
        '############################################
        'Alterado por Wagner
        If bAchouPrimeiro Then
            dQtdSaldoLivre = objNecesProdInfoPrimeiro.dQuantSaldoLivre
        Else
            dQtdSaldoLivre = objNecesProdInfoAux.dQuantSaldoLivre
        End If
        '############################################
        
        'se o saldo livre for >= necessario
        If dQtdSaldoLivre >= dQuantidadeProd Then
            
            'abater o necessario do saldo livre
            dQtdSaldoLivre = dQtdSaldoLivre - dQuantidadeProd
            dQuantidadeProd = 0
                    
        Else
                        
            'abater da necessidade o saldo livre
            dQuantidadeProd = dQuantidadeProd - dQtdSaldoLivre
            dQtdSaldoLivre = 0
            
        End If
        
        '############################################
        'Alterado por Wagner/ Corrigido em 30/08/05
        'Se já existe entrada para esse produto
        If bAchouPrimeiro Then
            objNecesProdInfoAux.dQuantSaldoLivreInicial = objNecesProdInfoPrimeiro.dQuantSaldoLivre
            objNecesProdInfoAux.dSaldoEmpenhosDiretos = objNecesProdInfoPrimeiro.dSaldoEmpenhosDiretos + (objNecesProdInfoPrimeiro.dQuantSaldoLivreInicial - objNecesProdInfoPrimeiro.dQuantSaldoLivre)
            objNecesProdInfoPrimeiro.dQuantSaldoLivre = dQtdSaldoLivre
            objNecesProdInfoAux.dSaldoDisponivel = objNecesProdInfoPrimeiro.dSaldoDisponivel
            objNecesProdInfoAux.dQuantSaldoLivre = dQtdSaldoLivre - dQuantidadeProd
        Else
            objNecesProdInfoAux.dQuantSaldoLivre = dQtdSaldoLivre
        End If
        '############################################
        
    End If
    
    If dQuantidadeProd > 0 Then
        
        'Alterado por Wagner 16/08/2006
        'Pode ter que gerar requisições de compra de insumos de kit de produtos comprados
'        If objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then
        
            objKit.sVersao = objNecesProdInfo.sVersao
            objKit.sProdutoRaiz = objNecesProdInfo.sProduto
        
            'tem que ler o proprio cara e os seus filhos
            lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
            If lErro <> SUCESSO And lErro <> 103203 Then gError 124210
            
            'Se o produto for um Kit, então leva a recursividade a diante
            If lErro = SUCESSO Then
            
                'percorre os elementos componentes do Kit Lido
                For Each objProdutoKit In objKit.colComponentes
                
                    objProduto.sCodigo = objProdutoKit.sProduto
            
                    'Lê o produto para descobrir as unidades de medidas associadas
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 124211
            
                    'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProdutoKit.sUnidadeMed, dFatorConversao)
                    If lErro <> SUCESSO Then gError 124212
                
                    If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                        
                        dQuantidadePai = (dQuantidadeProd / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                        
                    Else
                        
                        Set objNecesProdInfoFilho = New ClassNecesProdInfo
                
                        '##############################################
                        'Inserido por Wagner 06/10/2005
                        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                            
                            If objProdutoKit.sVersaoKitComp = "" Then
                
                                objKit.sProdutoRaiz = objProdutoKit.sProduto
                
                                lErro = CF("Kit_Le_Padrao", objKit)
                                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235
                
                                'Se não existe versão padrão p/ o Kit -> Erro.
                                If lErro <> SUCESSO Then gError 124282
                
                                objNecesProdInfoFilho.sVersao = objKit.sVersao
                
                            Else
                        
                                objNecesProdInfoFilho.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                
                            End If
                            
                        End If
                        '###############################################
                        
                        With objNecesProdInfoFilho
                            .sProduto = objProdutoKit.sProduto
                            .iCompras = objProduto.iCompras
                            .iClasseUM = objProduto.iClasseUM
                            .sUMEstoque = objProduto.sSiglaUMEstoque
                            .sUMCompra = objProduto.sSiglaUMCompra
                            If objNecesProdInfo.dtData1 <> DATA_NULA Then
                                .dtData2 = objNecesProdInfo.dtData1
                                .dtData1 = .dtData2 - objProduto.iTempoProducao
                            End If
                            
                            .iSeq = objProdutoKit.iSeq
                            .iNivel = objProdutoKit.iNivel
                            .sProdutoRaiz = objProdutoKit.sProdutoRaiz
                            .sVersaoPai = objProdutoKit.sVersao
                        End With
                
                        'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
                        dQuantidadeFilho = dQuantidadePai * (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda) * dFatorConversao)
                          
                        'Calcula de maneira recursiva a quantidade de componentes necessários.
                        lErro = Producao_Calcula_Necessidades_Nivel5(objCalcNecesProd, objNecesProdInfoFilho, dQuantidadeFilho, iNiveis + 1, objOP, objItemOP, iItemOP)
                        If lErro <> SUCESSO Then gError 124213
                
                    End If
                
                Next
            
            End If
    
        'End If
        
    End If
    
    Producao_Calcula_Necessidades_Nivel5 = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades_Nivel5:

    Producao_Calcula_Necessidades_Nivel5 = gErr
     
    Select Case gErr
          
        Case 124209 To 124213, 124235
        
        Case 124282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152666)
     
    End Select
     
    Exit Function

End Function

Public Sub Produto_Saida_Celula_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal dPrecoUnitario As Double, ByVal sProduto As String, ByVal iGrid_DescProduto_Col As Integer, ByVal sCliente As String, ByVal iFilialCliente As Integer, ByVal dDescontoCli As Double, ByVal iGrid_PercDesc_Col As Integer)

End Sub

Function OrdemProducao_Automatico(sCodigoOP As String, ByVal iFilialEmpresa As Integer) As Long
'retorna o número da proxima OP

Dim lTransacao As Long
Dim lErro As Long
Dim lNumProx As Long

On Error GoTo Erro_OrdemProducao_Automatico

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_PROX_OP", "OrdensProducaoTodas", "CodigoNumerico", lNumProx)  'Alterado por Jorge Specian
    If lErro <> SUCESSO Then gError 131862
    
    sCodigoOP = CStr(lNumProx)
    
    OrdemProducao_Automatico = SUCESSO

    Exit Function

Erro_OrdemProducao_Automatico:

    OrdemProducao_Automatico = gErr

    Select Case gErr

        Case 131862

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152668)

    End Select
    
    'Desfaz a transação
    Call Transacao_Rollback
    
    Exit Function

End Function

'###################################################################
'Inserido por Wagner
'Função copiada da Martineli - Desfaz a baixa de um pedido de venda
Function PedidoVenda_Desfaz_Baixa(ByVal objPedidoVenda As ClassPedidoDeVenda)

Dim alComando(0 To 13) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim tItemPedido As typeItemPedido
Dim lTransacao As Long
Dim bPossuiReserva As Boolean
Dim objBloqueioPV As ClassBloqueioPV
Dim sSelect As String
Dim iItemPVStatus As Integer
Dim iBloqueio  As Integer
Dim bIncluirBloq As Boolean
Dim objItemPV As ClassItemPedido
Dim objProduto As New ClassProduto

On Error GoTo Erro_PedidoVenda_Desfaz_Baixa

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 62208
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 86319

    'Busca o pedido da tabela dos Pedidos de Venda Baixados
    lErro = CF("PedidoDeVendaBaixado_LeMod", alComando(0), objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 62205 Then gError 62211
    
    'Se não encontrou ---> erro
    If lErro <> SUCESSO Then gError 62218
    
    lErro = CF("PedidoDeVenda_Insere", objPedidoVenda, alComando(2))
    If lErro <> SUCESSO Then gError 62220
       
    With tItemPedido
        
        .sProduto = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
        .sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)
        
        'BUsca os itens baixados do pedido
        sSelect = "SELECT NumIntDoc, FilialEmpresa, CodPedido, Produto, Quantidade, PrecoUnitario, PrecoTotal, ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantReservada, QuantFaturada, QuantOP, QuantSC, PercDesc1, TipoDesc1, PercDesc2, TipoDesc2, PercDesc3, TipoDesc3, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco FROM ItensPedidoDeVendaBaixados WHERE  FilialEmpresa = ? AND CodPedido = ? "
        lErro = Comando_ExecutarPos(alComando(3), sSelect, 0, .lNumIntDoc, .iFilialEmpresa, .lCodPedido, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .dQuantFaturadaAMais, .iPrioridade, .dComissao, .iTabelaPreco, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 62214
                        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62215
        
        If lErro <> AD_SQL_SUCESSO Then gError 62217 'Não encontrou itens
        
        'Para cada item do Pedido
        Do While lErro <> AD_SQL_SEM_DADOS
        
            If bPossuiReserva = False Then
                objProduto.sCodigo = .sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO Then gError 11111
                
                If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA Then bPossuiReserva = bPossuiReserva
            End If

            Set objItemPV = objPedidoVenda.colItensPedido.Add(.iFilialEmpresa, .lCodPedido, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, iItemPVStatus, .iControleEstoque, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3, .lNumIntDoc)
            With objItemPV
                .dPercentMaisReceb = tItemPedido.dPercentMaisReceb
                .dPercentMenosReceb = tItemPedido.dPercentMenosReceb
                .iRecebForaFaixa = tItemPedido.iRecebForaFaixa
                .dQuantFaturadaAMais = tItemPedido.dQuantFaturadaAMais
                .iPrioridade = tItemPedido.iPrioridade
                .dComissao = tItemPedido.dComissao
                .iTabelaPreco = tItemPedido.iTabelaPreco
            End With
            
            lErro = CF("ItensPVGrade_Le", objItemPV)
            If lErro <> SUCESSO Then gError 180414
            
            '#################################################
            'Inserido por Wagner 05/07/2006
            lErro = CF("PVItens_Le_Customizado", objPedidoVenda)
            If lErro <> SUCESSO Then gError 180414
            '#################################################
    
            lErro = CF("ItensPV_Insere", alComando(4), objItemPV)
            If lErro <> SUCESSO Then gError 62221
        
            'Exclui o item da tabela dos itens baixados
            lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM ItensPedidoDeVendaBaixados", alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 62222
        
            'BUsca o próximo item
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62269
            
        Loop
        
    End With
        
    'Exclui o pedido de venda da tabela dos baixados
    lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM PedidosDeVendaBaixados", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 62223
    
'    'Desfaz a baixa da tributação do pedido
'    lErro = CF("TributacaoPV_Desfaz_Baixa", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 62245
    
    'Desfaz a baixa das parcelas do pedido
    lErro = CF("ParcelasPV_Desfaz_Baixa", objPedidoVenda, alComando(7), alComando(8))
    If lErro <> SUCESSO Then gError 62248

    'Desfaz a baixa das comissoes do pedido
    lErro = CF("ComissoesPV_Desfaz_Baixa", objPedidoVenda, alComando(9), alComando(10))
    If lErro <> SUCESSO Then gError 62249
    
'    If giTipoVersao = VERSAO_FULL Then
        'Desfaz a baixa dos bloqueios do pedido
        lErro = CF("BloqueiosPV_Desfaz_Baixa", objPedidoVenda, alComando(11), alComando(12))
        If lErro <> SUCESSO Then gError 62250
        
        'seta a variável p\ uma possível criação de 1 bloqueio parcial p\ o PV
        iBloqueio = BLOQUEIO_NAO_RESERVA

        'Se pelo menos um dos produtos possui reserva
        If bPossuiReserva Then
            'Lê os bloqueios do pedido
            lErro = CF("BloqueiosPV_Le", objPedidoVenda)
            If lErro <> SUCESSO Then gError 62246
            
            'Verifica se já possui um bloqueio de estoque
            For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
                If (objPedidoVenda.iTipoPedido = BLOQUEIO_NAO_RESERVA Or objPedidoVenda.iTipoPedido = BLOQUEIO_TOTAL Or objPedidoVenda.iTipoPedido = BLOQUEIO_PARCIAL) And objBloqueioPV.dtDataLib = DATA_NULA Then
                    bIncluirBloq = False
                    Exit For
                End If
            Next
            
            'Se não possui nenum bloqueio de  estoque
            If bIncluirBloq = True Then
                'Gera um bloqueio do tipo passado
                Set objBloqueioPV = New ClassBloqueioPV
                objBloqueioPV.dtData = gdtDataAtual
                objBloqueioPV.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
                objBloqueioPV.lPedidoDeVendas = objPedidoVenda.lCodigo
                objBloqueioPV.iTipoDeBloqueio = iBloqueio
                objBloqueioPV.sCodUsuario = gsUsuario
                objBloqueioPV.dtData = gdtDataAtual
                objBloqueioPV.dtDataLib = DATA_NULA
                
                'Inclui o bloqueio na tabela dos bloqueios de PV
                lErro = CF("BloqueioPV_Insere", alComando(13), objBloqueioPV)
                If lErro <> AD_SQL_SUCESSO Then gError 62247
            End If
        End If
    'End If
    
    'Grava os dados estatísticos - INSERÇÃO
    lErro = CF("PedVenda_Grava_Estatisticas", objPedidoVenda)
    If lErro <> SUCESSO Then gError 52829
    
    lErro = ProdutoFilial_Adiciona_QuantPedida(objPedidoVenda)
    If lErro <> SUCESSO Then gError 51406

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 86320
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoVenda_Desfaz_Baixa = SUCESSO
    
    Exit Function
    
Erro_PedidoVenda_Desfaz_Baixa:

    PedidoVenda_Desfaz_Baixa = gErr
    
    Select Case gErr
    
        Case 62208
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 62211, 62245 To 62250, 180414, 62220, 62221
        
        Case 62214, 62215, 62269
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
        
        Case 62218
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case 62219
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPedidoVenda.lCodigo)
  
        Case 62222
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDIDODEVENDA", gErr, objPedidoVenda.lCodigo)
            
        Case 62223
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDODEVENDA", gErr, objPedidoVenda.lCodigo)
               
        Case 86319
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 86320
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152669)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'###################################################################

'Incluido por Jorge Specian
'-------------------------------------------------------------
Public Function OrcamentoVenda_Grava_EmTrans(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Grava o Orcamento de Venda - deve ser chamada dentro de uma Transação

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0) As Long
Dim objOrcamentoVenda_BD As New ClassOrcamentoVenda
Dim objOrcamentoVendaHist As New ClassOrcamentoVenda
Dim lNumIntPedVenda As Long
Dim lNumIntNFiscal As Long, iVersao As Integer

On Error GoTo Erro_OrcamentoVenda_Grava_EmTrans
    
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 101206
    Next
    
    objOrcamentoVenda.dtDataUltAlt = Date
    objOrcamentoVenda.dHoraUltAlt = CDbl(Time)
    
    'Faz o lock nos objetos vinculados ao Orcamento
    lErro = CF("OVObjetosVinculados_Lock", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101207

    'Lê OrcamentoVenda
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntPedVenda, NumIntNFiscal, Versao FROM OrcamentoVenda WHERE Codigo = ? AND FilialEmpresa = ? ", 0, lNumIntPedVenda, lNumIntNFiscal, iVersao, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 101208

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 101209

    'Se não existir
    If lErro <> AD_SQL_SUCESSO Then
    
        If gobjCRFAT.iOVGravacaoVersao = OV_VERSAO_NAO_GRAVA Then
            objOrcamentoVenda.iVersao = 0
        Else
            objOrcamentoVenda.iVersao = 1
        End If

        'Insere o orcamento novo no Banco de Dados
        lErro = CF("OrcamentoVenda_Insere_BD", objOrcamentoVenda)
        If lErro <> SUCESSO Then gError 101210

    'Se existir
    Else
        
        'Se o orçamento estiver vinculado a um Pedido de Venda, não poderá ser alterado -> Erro.
        If lNumIntPedVenda <> 0 Then gError 101211
        
        'Se o orçamento estiver vinculado a uma Nota Fiscal, não poderá ser alterado -> Erro.
        If lNumIntNFiscal <> 0 Then gError 101212
        
        If gobjCRFAT.iOVGravacaoVersao = OV_VERSAO_NAO_GRAVA Then
            objOrcamentoVenda.iVersao = 0
        Else
            If objOrcamentoVenda.iTrocarVersao = MARCADO Then
            
                objOrcamentoVendaHist.lCodigo = objOrcamentoVenda.lCodigo
                objOrcamentoVendaHist.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
                
                lErro = CF("OrcamentoVenda_Le_Completo", objOrcamentoVendaHist)
                If lErro <> SUCESSO Then gError 101210
                
                objOrcamentoVendaHist.iHistorico = MARCADO
                
                Set objOrcamentoVendaHist.objProjetoInfo = New ClassProjetoInfo
                
                objOrcamentoVendaHist.objProjetoInfo.iTipoOrigem = PRJ_CR_TIPO_OV
                objOrcamentoVendaHist.objProjetoInfo.lNumIntDocOrigem = objOrcamentoVendaHist.lNumIntDoc
                objOrcamentoVendaHist.objProjetoInfo.sCodigoOP = ""
                objOrcamentoVendaHist.objProjetoInfo.iFilialEmpresa = objOrcamentoVendaHist.iFilialEmpresa
            
                'Le as associação gravadas no BD para esse tipo de documento
                lErro = CF("ProjetoInfo_Le", objOrcamentoVendaHist.objProjetoInfo)
                If lErro <> SUCESSO Then gError 101210
                
                objOrcamentoVendaHist.objProjetoInfo.iTipoOrigem = PRJ_CR_TIPO_OVHIST
                
                'Insere o orcamento novo no Banco de Dados
                lErro = CF("OrcamentoVenda_Insere_BD", objOrcamentoVendaHist)
                If lErro <> SUCESSO Then gError 101210
                
                lErro = CF("ProjetoInfo_Grava", objOrcamentoVendaHist.objProjetoInfo, PRJ_CR_TIPO_OVHIST, objOrcamentoVendaHist.lNumIntDoc)
                If lErro <> SUCESSO Then gError 181706
                
                objOrcamentoVendaHist.objInfoAdic.lNumIntDoc = objOrcamentoVendaHist.lNumIntDoc
                objOrcamentoVendaHist.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OVHIST
                lErro = CF("InfoAdicionais_Grava", objOrcamentoVendaHist.objInfoAdic)
                If lErro <> SUCESSO Then gError 181706
                
                objOrcamentoVenda.iVersao = iVersao + 1
            Else
                objOrcamentoVenda.iVersao = iVersao
            End If
        End If
        
        'Altera o Orcamento de Venda e armazena os ítens do Pedido no BD em objOrcamentoVenda_BD
        lErro = CF("OrcamentoVenda_Altera_BD", objOrcamentoVenda, objOrcamentoVenda_BD, alComando(0))
        If lErro <> SUCESSO Then gError 101213

    End If

    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Grava", objOrcamentoVenda.objProjetoInfo, PRJ_CR_TIPO_OV, objOrcamentoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181706
    '############################################
    
    objOrcamentoVenda.objInfoAdic.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OV
    lErro = CF("InfoAdicionais_Grava", objOrcamentoVenda.objInfoAdic)
    If lErro <> SUCESSO Then gError 181706
    
    'se existe algum pedido de cotacao a ser gerado
    If objOrcamentoVenda.objCotacao.colCotacaoProduto.Count <> 0 Then
    
        lErro = OrcamentoVenda_Grava_Cotacao(objOrcamentoVenda)
        If lErro <> SUCESSO Then gError 181706
        
    End If
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrcamentoVenda_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Grava_EmTrans:

    OrcamentoVenda_Grava_EmTrans = gErr

    Select Case gErr

        Case 101207, 101210, 101213, 181706

        Case 101208, 101209
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)

        Case 101206
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101211
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_PEDVENDA", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
            
        Case 101212
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_NFISCAL", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152670)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'----------------------------------------------------------------

Function GeracaoOP_Grava(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_GeracaoOP_Grava

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 36628

    lErro = OP_Gera_Dados(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 138405

    lErro = GeracaoOP_Grava_EmTrans(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 138406

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 36640

    lErro = CF("WorkFlow_Trata_Transacao", "EST", objOrdemDeProducao.objTela, gsUsuario)
    If lErro <> SUCESSO Then gError 178247

    GeracaoOP_Grava = SUCESSO

    Exit Function

Erro_GeracaoOP_Grava:

    GeracaoOP_Grava = gErr

    Select Case gErr

        Case 36628
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 36640
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 138405, 138406, 178247
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152671)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function


End Function

Function OP_Gera_Dados(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataEmissao As Date
Dim iNumItens As Integer
Dim iNumItensBaixados As Integer
Dim objItemOP As New ClassItemOP
Dim alComando(0 To 1) As Long
Dim objKit As New ClassKit
Dim objOP As New ClassOrdemDeProducao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_OP_Gera_Dados

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 36627
    Next

    'verifica se o codigo da OP não coincide com uma OP baixada - Leitura da tabela de Ordens de Produção Baixadas
    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumItens FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 36629

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36630

    'se já existe uma OP baixada com este codigo => erro
    If lErro = AD_SQL_SUCESSO Then gError 36631

    'verifica se o codigo da OP não coincide com uma OP  - Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(alComando(1), "SELECT DataEmissao, NumItens, NumItensBaixados FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, iNumItensBaixados, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 36632

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36633

    'se já existe uma OP com este codigo => erro
    If lErro = AD_SQL_SUCESSO Then gError 36634

    'Se o item tiver como originador o Orcamento de Venda, busca a versao
    lErro = CF("OrdemProducao_Le_Versao_OV", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 131315

    'Se a OP for gerar outra OP
    If objOrdemDeProducao.iGeraOPsArvore <> MARCADO And objOrdemDeProducao.iGeraOP = MARCADO Then
       
        objOP.sCodigo = objOrdemDeProducao.sCodOPGerar
        objOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa

        lErro = CF("OrdemProducao_Le", objOP)
        If lErro <> SUCESSO And lErro <> 30368 Then gError 36636
        If lErro = SUCESSO Then gError 36634

    End If
    
    'inclui os itens da OP
    For Each objItemOP In objOrdemDeProducao.colItens
       
        If objItemOP.sVersao = "" Then
        
            'Verifica se é um kit
            objKit.sProdutoRaiz = objItemOP.sProduto
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 106653
            'If lErro <> SUCESSO Then gError 106654
            
            objItemOP.sVersao = objKit.sVersao

        End If
                
        lErro = Verifica_EstoqueMaximo(objItemOP)
        If lErro <> SUCESSO And lErro <> 41335 Then gError 40663

        If lErro = 41335 Then

            'Confirmar se vai gravar com a Quantidade maior que a quantidade maxima do estoque
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ESTOQUEMAXIMO", objItemOP.sProduto)

            If vbMsgRes = vbNo Then gError 40622

        End If

    Next
        
    lErro = CF("OrdemDeProducao_Gera_Operacoes", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 138406
        
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OP_Gera_Dados = SUCESSO

    Exit Function

Erro_OP_Gera_Dados:

    OP_Gera_Dados = gErr

    Select Case gErr

        Case 36627
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 36629, 36630
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS", gErr)

        Case 36631
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemDeProducao.sCodigo)

        Case 36632, 36633
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)

        Case 36634
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOrdemDeProducao.sCodigo)

        Case 36636, 36637, 106653, 138406, 131315, 40663, 40622

        Case 106654
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz, objKit.sVersao)
        
        Case 126480
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOrdemDeProducao.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152672)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function GeracaoOP_Grava_EmTrans(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDocItem As Long
Dim objItemOP As New ClassItemOP
Dim lEmpenho As Long
Dim alComando(0 To 16) As Long
Dim lCodigoNumerico As Long
Dim objItemOPGrade As ClassItemOP
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim vbMsgRes As VbMsgBoxResult
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_GeracaoOP_Grava_EmTrans

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 36627
    Next

    If IsNumeric(objOrdemDeProducao.sCodigo) Then
        lCodigoNumerico = StrParaLong(objOrdemDeProducao.sCodigo)
    End If
            
    'insere a ordem de produção
    lErro = Comando_Executar(alComando(0), "INSERT INTO OrdensDeProducao (Codigo, FilialEmpresa, DataEmissao, NumItens, NumItensBaixados,CodigoNumerico) VALUES (?,?,?,?,?,?) ", objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.dtDataEmissao, objOrdemDeProducao.colItens.Count, 0, lCodigoNumerico)
    If lErro <> AD_SQL_SUCESSO Then gError 36635

    'Se a OP for gerar RC ou OP de subproduto
    If objOrdemDeProducao.iGeraOPsArvore <> MARCADO And (objOrdemDeProducao.iGeraOP = MARCADO Or objOrdemDeProducao.iGeraReqCompra = MARCADO) Then
        lErro = OrdemProducao_Grava_OPSub_ReqCom(objOrdemDeProducao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    'obtem o proximo codigo para empenhos
    lErro = OrdemDeProducao_Grava3(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    'inclui os itens da OP
    For Each objItemOP In objOrdemDeProducao.colItens
       
        objItemOP.lNumIntDoc = lNumIntDocItem
        
        lErro = CF("OrdemDeProducao_Grava1", objItemOP, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lNumIntDocItem = lNumIntDocItem + 1

        For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
           
            Set objItemOPGrade = New ClassItemOP

            objItemOPGrade.dQuantidade = objItemRomaneioGrade.dQuantidade
            objItemOPGrade.iAlmoxarifado = objItemRomaneioGrade.iAlmoxarifado
            objItemOPGrade.sProdutoPai = objItemOP.sProduto
            objItemOPGrade.sProduto = objItemRomaneioGrade.sProduto
            objItemOPGrade.sVersao = objItemRomaneioGrade.sVersao
            objItemOPGrade.lNumIntDoc = lNumIntDocItem
            objItemOPGrade.lNumItemOP = objItemOP.lNumIntDoc
            
            'deixar o vinculo com os pedidos para itemopitemov
            objItemOPGrade.iFilialPedido = 0 'objItemOP.iFilialPedido
            objItemOPGrade.lCodPedido = 0 'objItemOP.lCodPedido
            
            objItemOPGrade.lNumIntOrigem = objItemOP.lNumIntOrigem
            objItemOPGrade.iOrigemPedido = objItemOP.iOrigemPedido
            objItemOPGrade.iClasseUM = objItemOP.iClasseUM
            objItemOPGrade.sSiglaUMEstoque = objItemOP.sSiglaUMEstoque
            objItemOPGrade.sSiglaUM = objItemOP.sSiglaUM
            objItemOPGrade.iFilialEmpresa = objItemOP.iFilialEmpresa
            objItemOPGrade.sCodigo = objItemOP.sCodigo
            objItemOPGrade.iItem = objItemOP.iItem
            objItemOPGrade.sCcl = objItemOP.sCcl
            objItemOPGrade.dtDataInicioProd = objItemOP.dtDataInicioProd
            objItemOPGrade.dtDataFimProd = objItemOP.dtDataFimProd
            objItemOPGrade.iPrioridade = objItemOP.iPrioridade
            objItemOPGrade.iSituacao = objItemOP.iSituacao
            objItemOPGrade.iDestinacao = objItemOP.iDestinacao
            objItemOPGrade.iBeneficiamento = objItemOP.iBeneficiamento
            objItemOPGrade.lNumIntEquipamento = objItemOP.lNumIntEquipamento
            objItemOPGrade.lNumIntEquipamento2 = objItemOP.lNumIntEquipamento2
            objItemOPGrade.iPossuiGrade = DESMARCADO
            objItemOPGrade.iTipo = objItemOP.iTipo

            'Verifica o estoque máximo do produto utilizado
            lErro = Verifica_EstoqueMaximo(objItemOPGrade)
            If lErro <> SUCESSO And lErro <> 41335 Then gError ERRO_SEM_MENSAGEM

            If lErro = 41335 Then

                'Confirmar se vai gravar com a Quantidade maior que a quantidade maxima do estoque
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ESTOQUEMAXIMO", objItemOPGrade.sProduto)

                If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

            End If

            lErro = CF("OrdemDeProducao_Grava1", objItemOPGrade, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            lNumIntDocItem = lNumIntDocItem + 1

        Next
        
        If objItemOP.lCodPedido <> 0 Then
        
            Set objItemPV = New ClassItemPedido

            objItemPV.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
            objItemPV.lCodPedido = objItemOP.lCodPedido
            objItemPV.sProduto = objItemOP.sProduto
            objItemPV.iTipoOP = OP_TIPO_OP
    
            'aumenta a quantidade em OP do Pedido de Venda
            lErro = OrdemDeProducao_Grava6(alComando(2), alComando(3), objItemPV, objItemOP.dQuantidade)
            If lErro <> SUCESSO And lErro <> 41346 Then gError ERRO_SEM_MENSAGEM
    
            If lErro = 41346 Then gError 202257
        
        Else
        
            lErro = ItemOPItemPV_Insere(objItemOP)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    Next
    
    'atualiza o proximo codigo para lNumIntDocItem e lEmpenho
    lErro = OrdemDeProducao_Grava5(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objOrdemDeProducao.iGeraOPsArvore = MARCADO Then
        lErro = CF("OrdemProducao_Grava_Nodes", objOrdemDeProducao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = CF("OrdemDeProducao_Grava_Operacoes_EmTrans", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    GeracaoOP_Grava_EmTrans = SUCESSO

    Exit Function

Erro_GeracaoOP_Grava_EmTrans:

    GeracaoOP_Grava_EmTrans = gErr

    Select Case gErr

        Case 36627
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 36635
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDEPRODUCAO", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case 202257
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemPV.lCodPedido, objItemPV.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152673)

    End Select

   'Fechamento comando

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'########################################################################
'Inserido por Wagner
Function Producao_Calcula_Necessidades_Nivel4(ByVal objOP As ClassOrdemDeProducao, ByVal sProduto As String, dEmpenhos As Double, ByVal iItemOP As Integer) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim objItemOP As ClassItemOP
Dim dQtdEmp As Double
Dim dQtdReq As Double

On Error GoTo Erro_Producao_Calcula_Necessidades_Nivel4

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140454
    Next
    
    dEmpenhos = 0

    iIndice = 0
    For Each objItemOP In objOP.colItens
    
        'Tem que abater o empenho só do próprio item e dos anteriores que já foram gravados
        iIndice = iIndice + 1
        If iItemOP <= iIndice Then
    
            lErro = Comando_Executar(alComando(1), "SELECT Quantidade, QuantidadeRequisitada FROM Empenho WHERE NumIntDocItemOP = ? AND Produto = ?", dQtdEmp, dQtdReq, objItemOP.lNumIntDoc, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 140455
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140456
    
            Do While lErro = AD_SQL_SUCESSO
    
                dEmpenhos = dEmpenhos + dQtdEmp - dQtdReq
    
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140457
    
            Loop
            
        End If

    Next
        
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Producao_Calcula_Necessidades_Nivel4 = SUCESSO

    Exit Function

Erro_Producao_Calcula_Necessidades_Nivel4:

    Producao_Calcula_Necessidades_Nivel4 = gErr

    Select Case gErr

        Case 140454
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140455 To 140457
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152674)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'########################################################################
'###########################################################
'Inserido por Wagner 09/12/05
Public Function ItensOV_Grava_Customizado(ByVal objItensOV As ClassItemOV) As Long
    ItensOV_Grava_Customizado = SUCESSO
End Function
'###########################################################

Public Function PedidoVenda_Grava_Customizacoes(objPedidoVenda As ClassPedidoDeVenda) As Long

    PedidoVenda_Grava_Customizacoes = SUCESSO
    
End Function

Public Function PedidoVenda_Grava_Customizado(objPedidoVenda As ClassPedidoDeVenda) As Long

    PedidoVenda_Grava_Customizado = SUCESSO
    
End Function

Function ItensPVAux_Grava(ByVal lComando As Long, sProduto As String, iStatus As Integer, lNumIntDoc As Long, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPV As ClassItemPedido) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_ItensPVAux_Grava

    'Localiza o item no Banco de Dados
    'lErro = Comando_ExecutarPos(lComando, "SELECT Produto, Status, NumIntDoc FROM ItensPedidoDeVenda WHERE CodPedido = ? AND FilialEmpresa = ? AND Produto = ? ", 0, sProduto, iStatus, lNumIntDoc, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, objItemPV.sProduto)
    lErro = Comando_ExecutarPos(lComando, "SELECT Produto, Status, NumIntDoc FROM ItensPedidoDeVenda WHERE NumIntDoc = ? ", 0, sProduto, iStatus, lNumIntDoc, objItemPV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 177396
        
    ItensPVAux_Grava = SUCESSO

    Exit Function

Erro_ItensPVAux_Grava:

    ItensPVAux_Grava = gErr

    Select Case gErr
    
        Case 177396

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177395)

    End Select

    Exit Function

End Function

Function InfoAdicDocItem_Insere(ByVal objInfoAdicDocItem As ClassInfoAdicDocItem, ByVal lComando As Long, Optional ByVal lComandoPExp As Long = 0) As Long

Dim lErro As Long

On Error GoTo Erro_InfoAdicDocItem_Insere

    With objInfoAdicDocItem
        lErro = Comando_Executar(lComando, "INSERT INTO InfoAdicDocItem (TipoDoc, NumIntDocItem, DataLimiteFaturamento, NumPedidoCompra, ItemPedCompra, IncluiValorTotal, Msg, Msg2,NumIntDE,NumRE) VALUES (?,?,?,?,?,?,?,?,?,?)", _
            .iTipoDoc, .lNumIntDocItem, .dtDataLimiteFaturamento, .sNumPedidoCompra, .lItemPedCompra, .iIncluiValorTotal, left(.sMsg, STRING_INFOADIC_ITEM_MSG), Mid(.sMsg, STRING_INFOADIC_ITEM_MSG + 1, STRING_INFOADIC_ITEM_MSG), .lNumIntDE, .sNumRE)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 201229
    
    lErro = InfoAdicDocItemDetExp_Insere(objInfoAdicDocItem, lComandoPExp)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    InfoAdicDocItem_Insere = SUCESSO

    Exit Function

Erro_InfoAdicDocItem_Insere:

    InfoAdicDocItem_Insere = gErr

    Select Case gErr
    
        Case 201229
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INFOADICDOCITEM", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201228)

    End Select

    Exit Function

End Function

Private Function InfoAdicDocItemDetExp_Insere(ByVal objInfoAdicDocItem As ClassInfoAdicDocItem, Optional ByVal lComandoPExp As Long = 0) As Long

Dim lErro As Long
Dim objDetExp As ClassInfoAdicDocItemDetExp
Dim lComandoExp As Long, iSeq As Integer

On Error GoTo Erro_InfoAdicDocItemDetExp_Insere

    lComandoExp = 0
    
    If objInfoAdicDocItem.colDetExportacao.Count > 0 Then
    
        If lComandoPExp = 0 Then
            lComandoExp = Comando_Abrir
            If lComandoExp = 0 Then gError 213675
        Else
            lComandoExp = lComandoPExp
        End If
    
        iSeq = 0
        For Each objDetExp In objInfoAdicDocItem.colDetExportacao
                    
            iSeq = iSeq + 1
            With objDetExp
                lErro = Comando_Executar(lComandoExp, "INSERT INTO InfoAdicDocItemDetExport (TipoDoc, NumIntDocItem, Seq, NumDrawback, NumRegistExport, ChvNFe, QuantExport) VALUES (?,?,?,?,?,?,?)", _
                    objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem, iSeq, objDetExp.sNumDrawback, objDetExp.sNumRegistExport, objDetExp.sChvNFe, objDetExp.dQuantExport)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 213676
        
        Next
    
        If lComandoPExp = 0 Then
            Call Comando_Fechar(lComandoExp)
        End If
        
    End If

    InfoAdicDocItemDetExp_Insere = SUCESSO

    Exit Function

Erro_InfoAdicDocItemDetExp_Insere:

    InfoAdicDocItemDetExp_Insere = gErr

    Select Case gErr
    
        Case 213675
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INFOADICDOCITEM", gErr)
        
        Case 213676
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213677)

    End Select
    
    If lComandoPExp = 0 And lComandoExp <> 0 Then
        Call Comando_Fechar(lComandoExp)
    End If

    Exit Function

End Function

Function InfoAdicDocItem_Exclui(ByVal objInfoAdicDocItem As ClassInfoAdicDocItem) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 3) As Long, iIncluiValorTotal As Integer

On Error GoTo Erro_InfoAdicDocItem_Exclui

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201235
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT IncluiValorTotal FROM InfoAdicDocItem WHERE TipoDoc = ? AND NumIntDocItem = ?", 0, iIncluiValorTotal, objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then gError 201237
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201238
    If lErro <> AD_SQL_SUCESSO Then gError 201239
    
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 201240
    
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM InfoAdicDocItem", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 201241
    
    lErro = Comando_Executar(alComando(3), "DELETE FROM InfoAdicDocItemDetExport WHERE TipoDoc = ? AND NumIntDocItem = ?", objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then gError 201241
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    InfoAdicDocItem_Exclui = SUCESSO
    
    Exit Function
    
Erro_InfoAdicDocItem_Exclui:

    InfoAdicDocItem_Exclui = gErr

    Select Case gErr

        Case 201235 To 201241
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INFOADICDOCITEM", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201234)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function InfoAdicDocItem_Altera(ByVal objInfoAdicDocItem As ClassInfoAdicDocItem) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 4) As Long, iIncluiValorTotal As Integer

On Error GoTo Erro_InfoAdicDocItem_Altera

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201235
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT IncluiValorTotal FROM InfoAdicDocItem WHERE TipoDoc = ? AND NumIntDocItem = ?", 0, iIncluiValorTotal, objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then gError 201245
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201246
    If lErro <> AD_SQL_SUCESSO Then gError 201247
    
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 201248
    
    With objInfoAdicDocItem
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE InfoAdicDocItem SET DataLimiteFaturamento = ?, NumPedidoCompra = ?, ItemPedCompra = ?, IncluiValorTotal = ?, Msg = ?, Msg2 = ?,NumIntDE = ?,NumRE = ?", alComando(1), _
            .dtDataLimiteFaturamento, .sNumPedidoCompra, .lItemPedCompra, .iIncluiValorTotal, left(.sMsg, STRING_INFOADIC_ITEM_MSG), Mid(.sMsg, STRING_INFOADIC_ITEM_MSG + 1, STRING_INFOADIC_ITEM_MSG), .lNumIntDE, .sNumRE)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 201249
    
    lErro = Comando_Executar(alComando(3), "DELETE FROM InfoAdicDocItemDetExport WHERE TipoDoc = ? AND NumIntDocItem = ?", objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then gError 201249
    
    lErro = InfoAdicDocItemDetExp_Insere(objInfoAdicDocItem, alComando(4))
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    InfoAdicDocItem_Altera = SUCESSO
    
    Exit Function
    
Erro_InfoAdicDocItem_Altera:

    InfoAdicDocItem_Altera = gErr

    Select Case gErr

        Case 201243 To 201249
            Call Rotina_Erro(vbOKOnly, "ERRO_ATALIZACAO_INFOADICDOCITEM", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201242)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ItensPV_Insere2(ByVal lComando As Long, ByVal objItemPV As ClassItemPedido, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_ItensPV_Insere2

    'Insere o item no BD
    lErro = CF("ItensPV_Insere", lComando, objItemPV)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    objItemPV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMPV
    objItemPV.objInfoAdicDocItem.lNumIntDocItem = objItemPV.lNumIntDoc
    
    lErro = CF("InfoAdicDocItem_Insere", objItemPV.objInfoAdicDocItem, lComando2)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ItensPV_Insere2 = SUCESSO

    Exit Function

Erro_ItensPV_Insere2:

    ItensPV_Insere2 = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177398)

    End Select

    Exit Function

End Function

Function ItensPV_Insere(ByVal lComando As Long, ByVal objItemPV As ClassItemPedido) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_ItensPV_Insere

    'Insere o item no BD
    lErro = Comando_Executar(lComando, "INSERT INTO ItensPedidoDeVenda (NumIntDoc,FilialEmpresa,CodPedido,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objItemPV.lNumIntDoc, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.sLote, objItemPV.iStatus, objItemPV.dValorAbatComissao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.dQuantOP, objItemPV.dQuantSC, _
        objItemPV.dPercentMaisReceb, objItemPV.dPercentMenosReceb, objItemPV.iRecebForaFaixa, objItemPV.dQuantFaturadaAMais, objItemPV.iPrioridade, objItemPV.dComissao, objItemPV.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 177397
        
    ItensPV_Insere = SUCESSO

    Exit Function

Erro_ItensPV_Insere:

    ItensPV_Insere = gErr

    Select Case gErr
    
        Case 177397
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177398)

    End Select

    Exit Function

End Function

Public Function PV_PedidoRepr_Valida(ByVal lPedido As Long, ByVal iFilialEmpresa As Integer, ByVal sPedRepr As String) As Long

Dim lErro As Long

On Error GoTo Erro_PV_PedidoRepr_Valida
    
    PV_PedidoRepr_Valida = SUCESSO
    
    Exit Function
    
Erro_PV_PedidoRepr_Valida:

    PV_PedidoRepr_Valida = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180379)

    End Select

    Exit Function
    
End Function

Public Function PV_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long
    PV_Grava_Customizado = SUCESSO
End Function

Public Function PV_Exclui_Reservas() As Long
    PV_Exclui_Reservas = SUCESSO
End Function

Function BloqueiosPV_Grava_Trans(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere no BD os Bloqueios do Pedido de Venda


Dim lComando As Long
Dim lErro As Long
Dim objBloqueioPV As ClassBloqueioPV
Dim iSequencial As Integer
Dim lTransacao As Long

On Error GoTo Erro_BloqueiosPV_Grava_Trans

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178459

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178461

     'Para cada Bloqueio do Pedido
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV

        'Gera o sequencial automático para o Bloqueio
        lErro = CF("BloqueioPV_Automatico", objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, iSequencial)
        If lErro <> SUCESSO Then gError 178462

        'Atribui Sequencial ANTES de INSERIR
        objBloqueioPV.iSequencial = iSequencial

        'Insere o bloqueio na tabela de Bloqueios de Pedido de Venda
        lErro = Comando_Executar(lComando, "INSERT INTO BloqueiosPV(FilialEmpresa,PedidoDeVenda,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib,DataLib,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial, objBloqueioPV.iTipoDeBloqueio, objBloqueioPV.sCodUsuario, objBloqueioPV.sResponsavel, objBloqueioPV.dtData, objBloqueioPV.sCodUsuarioLib, objBloqueioPV.sResponsavelLib, objBloqueioPV.dtDataLib, objBloqueioPV.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 178463

    Next
     
    'Fecha o comando
    Call Comando_Fechar(lComando)

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 178460

    BloqueiosPV_Grava_Trans = SUCESSO

    Exit Function

Erro_BloqueiosPV_Grava_Trans:

    BloqueiosPV_Grava_Trans = gErr

    Select Case gErr

        Case 178459
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 178460
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 178461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178462

        Case 178463
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BLOQUEIOPV", gErr, objBloqueioPV.iFilialEmpresa, objBloqueioPV.lPedidoDeVendas, iSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178464)

    End Select

    Call Comando_Fechar(lComando)

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function ItensPVEntrega_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'exclui as  datas de entrega do pedido e em seguinda grava as novas

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim objItemPV As ClassItemPedido
Dim objDataEntrega As ClassDataEntrega
Dim lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_ItensPVEntrega_Grava

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183208
    Next

    'Exclui o Pedido da Tabela de Pedidos
    lErro = Comando_Executar(alComando(1), "DELETE FROM ItensPVEntrega WHERE NumIntPV = ?", objPedidoVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183209

     'Para cada Bloqueio do Pedido
    For Each objItemPV In objPedidoVenda.colItensPedido

        For Each objDataEntrega In objItemPV.colItemPVEntrega

            'Gera o número interno automático do Pedido de Venda
            lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ITENSPVENTREGA", lNumIntDoc)
            If lErro <> SUCESSO Then gError 183210

            'Insere o bloqueio na tabela de Bloqueios de Pedido de Venda
            lErro = Comando_Executar(alComando(2), "INSERT INTO ItensPVEntrega(NumIntDoc, NumIntPV, NumIntItemPV, DataEntrega, Quantidade, PedidoCliente) VALUES (?,?,?,?,?,?)", lNumIntDoc, objPedidoVenda.lNumIntDoc, objItemPV.lNumIntDoc, objDataEntrega.dtDataEntrega, objDataEntrega.dQuantidade, objDataEntrega.sPedidoCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 183211

        Next

    Next
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPVEntrega_Grava = SUCESSO

    Exit Function

Erro_ItensPVEntrega_Grava:

    ItensPVEntrega_Grava = gErr

    Select Case gErr

        Case 183208
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183209
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPVENTREGA", gErr)

        Case 183210

        Case 183211
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSPVENTREGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183212)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPVEntrega_Exclui(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'exclui as  datas de entrega do pedido

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_ItensPVEntrega_Exclui

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183213
    Next

    'Exclui o Pedido da Tabela de Pedidos
    lErro = Comando_Executar(alComando(1), "DELETE FROM ItensPVEntrega WHERE NumIntPV = ?", objPedidoVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 183214

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPVEntrega_Exclui = SUCESSO

    Exit Function

Erro_ItensPVEntrega_Exclui:

    ItensPVEntrega_Exclui = gErr

    Select Case gErr

        Case 183213
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183214
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPVENTREGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183215)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Integracao_PV_Obtem_NaturezaOP_Customizado(ByVal objPV As ClassPedidoDeVenda, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco)
'Obtém a natureza de operação para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_NaturezaOP_Customizado

   
    Integracao_PV_Obtem_NaturezaOP_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_NaturezaOP_Customizado:

    Integracao_PV_Obtem_NaturezaOP_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190089)

    End Select
       
    Exit Function

End Function

Public Function Integracao_PV_Obtem_CanalVenda_Customizado(ByVal objPV As ClassPedidoDeVenda, ByVal objTipoCliente As ClassTipoCliente)
'Obtém o canal de venda para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_CanalVenda_Customizado

   
    Integracao_PV_Obtem_CanalVenda_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_CanalVenda_Customizado:

    Integracao_PV_Obtem_CanalVenda_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190089)

    End Select
       
    Exit Function

End Function

Public Function Integracao_PV_Obtem_Frete_Customizado(ByVal objPV As ClassPedidoDeVenda, ByVal objTipoCliente As ClassTipoCliente)
'Obtém o canal de venda para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_Frete_Customizado

   
    Integracao_PV_Obtem_Frete_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_Frete_Customizado:

    Integracao_PV_Obtem_Frete_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190089)

    End Select
       
    Exit Function

End Function

Public Function Integracao_PV_Obtem_UM_Customizado(ByVal objItemPV As ClassItemPedido, ByVal objProduto As ClassProduto)
'Obtém o canal de venda para a importação de pedidos de venda

Dim lErro As Long

On Error GoTo Erro_Integracao_PV_Obtem_UM_Customizado

   
    Integracao_PV_Obtem_UM_Customizado = SUCESSO
    
    Exit Function

Erro_Integracao_PV_Obtem_UM_Customizado:

    Integracao_PV_Obtem_UM_Customizado = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190089)

    End Select
       
    Exit Function

End Function

Public Function Exporta_Dados_SldProd_AlmValido_Cust(ByVal iAlmoxarifado As Integer, bValido As Boolean)
'Verifica se o saldo do almoxarifado deve ser exportado

Dim lErro As Long

On Error GoTo Erro_Exporta_Dados_SldProd_AlmValido_Cust

    bValido = True
   
    Exporta_Dados_SldProd_AlmValido_Cust = SUCESSO
    
    Exit Function

Erro_Exporta_Dados_SldProd_AlmValido_Cust:

    Exporta_Dados_SldProd_AlmValido_Cust = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190118)

    End Select
       
    Exit Function

End Function

Public Function Integracao_Cliente_Obtem_Tipo_Customizado(ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco, ByVal iTipo As Integer, ByVal sTipo As String)
    objCliente.iTipo = iTipo
    Integracao_Cliente_Obtem_Tipo_Customizado = SUCESSO
End Function

Public Function OrdemDeProducao_Exclui_Usu(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
    OrdemDeProducao_Exclui_Usu = SUCESSO
End Function

Private Function ItensOV_Grava_FormPreco(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objItemOV As ClassItemOV) As Long

Dim lErro As Long, objItemFP As ClassItemFormPreco
Dim iIndice As Integer, alComando(1 To 4) As Long, iSituacao As Integer
Dim lNumIntFormPreco As Long, iFPSituacao As Integer, bGerarCotacoes As Boolean
Dim lFornecedor As Long, iFilialForn As Integer

On Error GoTo Erro_ItensOV_Grava_FormPreco

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184330
    Next
    
    For Each objItemFP In objItemOV.colItensFormPreco
    
        objItemFP.iTipoDocOrigem = FPORIGEM_ITEMOV
        objItemFP.lNumIntDocOrigem = objItemOV.lNumIntDoc
        
        bGerarCotacoes = False
        
        'se é um item novo
        If objItemFP.lNumIntDoc = 0 Then
        
            lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ITEM_FORM_PRECO", lNumIntFormPreco)
            If lErro <> SUCESSO Then gError 184336
        
            objItemFP.lNumIntDoc = lNumIntFormPreco
        
        End If
        
        If objItemFP.iSituacao = FPSITUACAO_COTAR Then
            
            'verificar se existe fornecedor associado ao produto
            lErro = Comando_Executar(alComando(4), "SELECT Fornecedor,FilialForn FROM FornecedorProdutoFF WHERE Produto = ? AND FilialEmpresa = ?", _
                lFornecedor, iFilialForn, objItemFP.sProduto, objOrcamentoVenda.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 184760
        
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184761
            If lErro = AD_SQL_SUCESSO Then
            
                objItemFP.iSituacao = FPSITUACAO_EM_COTACAO
                bGerarCotacoes = True
            
            End If
            
        End If
        
        With objItemFP
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensFormPreco (NumIntDoc, TipoDocOrigem, NumIntDocOrigem, Sequencial, Produto, UnidMed, Qtde, CustoUnit, PercentMargem, PrecoUnit, PrecoTotal, Situacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                .lNumIntDoc, .iTipoDocOrigem, .lNumIntDocOrigem, .iSequencial, .sProduto, .sUnidMed, .dQtde, .dCustoUnit, .dPercentMargem, .dPrecoUnit, .dPrecoTotal, .iSituacao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184337
        
        'gerar as cotacoes
        If bGerarCotacoes Then
        
            lErro = ItemFormPreco_Inclui_Cotacoes_OV(objOrcamentoVenda, objItemOV, objItemFP)
            If lErro <> SUCESSO Then gError 184347
            
        End If
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensOV_Grava_FormPreco = SUCESSO
    
    Exit Function
    
Erro_ItensOV_Grava_FormPreco:

    ItensOV_Grava_FormPreco = gErr

    Select Case gErr
    
        Case 184330
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184336, 184347
        
        Case 184337
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSFORMPRECO", gErr)

        Case 184760, 184761
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184327)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ItensOV_Exclui_FormPreco(ByVal objItemOV As ClassItemOV) As Long

Dim lErro As Long, lNumIntFormPreco As Long, iFPSituacao As Integer
Dim iIndice As Integer, alComando(1 To 3) As Long

On Error GoTo Erro_ItensOV_Exclui_FormPreco

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184330
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Situacao FROM ItensFormPreco WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ?", 0, lNumIntFormPreco, iFPSituacao, FPORIGEM_ITEMOV, objItemOV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184331

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184332
    
    Do While lErro = AD_SQL_SUCESSO
    
        If iFPSituacao = FPSITUACAO_COTADO Or iFPSituacao = FPSITUACAO_EM_COTACAO Then
        
            'excluir o vinculo da cotacao
            lErro = Comando_Executar(alComando(3), "DELETE FROM CotacaoProdutoItemFP WHERE NumIntItemFormPreco = ?", lNumIntFormPreco)
            If lErro <> AD_SQL_SUCESSO Then gError 184333
            
        End If
        
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensFormPreco", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 184334
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184335
    
    Loop
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensOV_Exclui_FormPreco = SUCESSO
    
    Exit Function
    
Erro_ItensOV_Exclui_FormPreco:

    ItensOV_Exclui_FormPreco = gErr

    Select Case gErr

        Case 184330
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184331 To 184335
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSFORMPRECO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184328)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ItemFormPreco_Inclui_Cotacoes_OV(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objItemOV As ClassItemOV, ByVal objItemFP As ClassItemFormPreco) As Long

Dim lErro As Long
Dim objCotProduto As ClassCotacaoProduto
Dim bAchou As Boolean

On Error GoTo Erro_ItemFormPreco_Inclui_Cotacoes_OV

    bAchou = False
    For Each objCotProduto In objOrcamentoVenda.objCotacao.colCotacaoProduto
            
        If objCotProduto.sProduto = objItemFP.sProduto Then
            bAchou = True
            Exit For
        End If
        
    Next
    
    If bAchou = False Then
    
        Set objCotProduto = New ClassCotacaoProduto
        
        objCotProduto.dQuantidade = objItemFP.dQtde
        objCotProduto.sProduto = objItemFP.sProduto
        objCotProduto.sUM = objItemFP.sUnidMed
    
        objCotProduto.lNumIntDoc = objItemFP.lNumIntDoc 'provisorio
                    
        'Adiciona em colCotacaoProduto
        objOrcamentoVenda.objCotacao.colCotacaoProduto.Add objCotProduto
        
    Else
    
        objCotProduto.dQuantidade = objCotProduto.dQuantidade + objItemFP.dQtde
    
    End If
    
    objCotProduto.colItemFormPreco.Add objItemFP
    
    ItemFormPreco_Inclui_Cotacoes_OV = SUCESSO
    
    Exit Function
    
Erro_ItemFormPreco_Inclui_Cotacoes_OV:

    ItemFormPreco_Inclui_Cotacoes_OV = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184346)

    End Select
    
    Exit Function

End Function

Public Function OrcamentoVenda_Grava_Cotacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long, lCotacao As Long, iCodigo As Integer, iComprador As Integer
Dim lNumIntProvisorio As Long

On Error GoTo Erro_OrcamentoVenda_Grava_Cotacao

    'Obtem o proximo codigo de Cotacao disponivel
    lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROXIMO_CODIGO_COTACAO", "Cotacao", "Codigo", lCotacao)
    If lErro <> SUCESSO Then gError 184350
    
    'obter comprador
    iComprador = 1 '??? depois trocar para leitura de comprador ativo ou colocar em config o padrao
    
    With objOrcamentoVenda.objCotacao
        .iComprador = iComprador
        .iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
        .sDescricao = "Ref. ao Orçamento de Venda " & CStr(objOrcamentoVenda.lCodigo)
        .dtData = gdtDataAtual
        .lCodigo = lCotacao
        iCodigo = COD_A_VISTA
        .colCondPagtos.Add (iCodigo)
        .iFilialDestino = objOrcamentoVenda.iFilialEmpresa
        .iTipoDestino = TIPO_DESTINO_EMPRESA
    End With
    
    lErro = OrcamentoVenda_PreparaCotacoes(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 184349
    
    lErro = CF("Cotacao_Grava_Pedidos_EmTrans", objOrcamentoVenda.objCotacao, objOrcamentoVenda.colPedidoCotacao)
    If lErro <> SUCESSO Then gError 184349
    
    OrcamentoVenda_Grava_Cotacao = SUCESSO
    
    Exit Function
    
Erro_OrcamentoVenda_Grava_Cotacao:

    OrcamentoVenda_Grava_Cotacao = gErr

    Select Case gErr
    
        Case 184349, 184350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184348)

    End Select
    
    Exit Function

End Function

Function OrcamentoVenda_PreparaCotacoes(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long, bAchou As Boolean
Dim objPedidoCotacao As ClassPedidoCotacao
Dim objCotProduto As ClassCotacaoProduto
Dim objItemPedCotacao As ClassItemPedCotacao
Dim iIndice As Integer, alComando(1 To 3) As Long
Dim lFornecedor As Long, iFilialForn As Integer

On Error GoTo Erro_OrcamentoVenda_PreparaCotacoes

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184352
    Next
    
    For Each objCotProduto In objOrcamentoVenda.objCotacao.colCotacaoProduto
    
        lErro = Comando_Executar(alComando(1), "SELECT Fornecedor,FilialForn FROM FornecedorProdutoFF WHERE Produto = ? AND FilialEmpresa = ?", _
            lFornecedor, iFilialForn, objCotProduto.sProduto, objOrcamentoVenda.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 184353
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184354
    
        Do While lErro = AD_SQL_SUCESSO
                
            bAchou = False
            For Each objPedidoCotacao In objOrcamentoVenda.colPedidoCotacao
            
                If objPedidoCotacao.lFornecedor = lFornecedor And objPedidoCotacao.iFilial = iFilialForn Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If bAchou = False Then
            
                Set objPedidoCotacao = New ClassPedidoCotacao
                
                With objPedidoCotacao
                    .lFornecedor = lFornecedor
                    .iFilial = iFilialForn
                    .iStatus = STATUS_GERADO_NAO_ATUALIZADO
                    .iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
                    .dtData = gdtDataHoje
                    .dtDataEmissao = DATA_NULA
                    .dtDataValidade = DATA_NULA
                End With
                
                objOrcamentoVenda.colPedidoCotacao.Add objPedidoCotacao
                
            End If
            
            'Cria Item de Pedido de Cotação
            Set objItemPedCotacao = New ClassItemPedCotacao
            
            With objItemPedCotacao
                .lCotacaoProduto = objCotProduto.lNumIntDoc
                .sProduto = objCotProduto.sProduto
                .dQuantidade = objCotProduto.dQuantidade
                .sUM = objCotProduto.sUM
            End With
            
            'Adiciona o Item na coleção de Pedido de Cotação
            objPedidoCotacao.colItens.Add objItemPedCotacao
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184355
        
        Loop
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OrcamentoVenda_PreparaCotacoes = SUCESSO
    
    Exit Function
    
Erro_OrcamentoVenda_PreparaCotacoes:

    OrcamentoVenda_PreparaCotacoes = gErr

    Select Case gErr
    
        Case 184352
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 184353 To 184355
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184351)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ItensOV_Exclui_FPAlteracao(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objOrcamentoVenda_BD As ClassOrcamentoVenda) As Long
'exclui itensformpreco que já existiam e continuaram a existir numa alteracao de orcamento de venda

Dim lErro As Long, lNumIntFormPreco As Long, iFPSituacao As Integer
Dim iIndice As Integer, alComando(1 To 3) As Long
Dim objItemOV As ClassItemOV, objItemFP As ClassItemFormPreco

On Error GoTo Erro_ItensOV_Exclui_FPAlteracao

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184330
    Next
    
    For Each objItemOV In objOrcamentoVenda.colItens
    
        For Each objItemFP In objItemOV.colItensFormPreco
        
            'somente para
            If objItemFP.lNumIntDoc <> 0 Then
            
                lErro = Comando_ExecutarPos(alComando(1), "SELECT Situacao FROM ItensFormPreco WHERE NumIntDoc = ?", 0, iFPSituacao, objItemFP.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 184331
            
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184332
                
                Do While lErro = AD_SQL_SUCESSO
                
                    If objItemFP.iSituacao = FPSITUACAO_NAO_COTAR Or objItemFP.iSituacao = FPSITUACAO_COTAR Then
                    
                        'excluir o vinculo da cotacao
                        lErro = Comando_Executar(alComando(3), "DELETE FROM CotacaoProdutoItemFP WHERE NumIntItemFormPreco = ?", lNumIntFormPreco)
                        If lErro <> AD_SQL_SUCESSO Then gError 184333
                        
                    End If
                    
                    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensFormPreco", alComando(1))
                    If lErro <> AD_SQL_SUCESSO Then gError 184334
                    
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184335
                
                Loop
        
            End If
            
        Next
            
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensOV_Exclui_FPAlteracao = SUCESSO
    
    Exit Function
    
Erro_ItensOV_Exclui_FPAlteracao:

    ItensOV_Exclui_FPAlteracao = gErr

    Select Case gErr

        Case 184330
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184331 To 184335
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSFORMPRECO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184328)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function OrdemDeProducao_Grava_Customizado(ByVal objOP As ClassOrdemDeProducao) As Long
    OrdemDeProducao_Grava_Customizado = SUCESSO
End Function

Public Function Integracao_Cliente_Obtem_PadraoCobr_Cust(ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco)
    Integracao_Cliente_Obtem_PadraoCobr_Cust = SUCESSO
End Function

Private Function ItemOPItemPV_Insere(ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim objItemOPItemPV As ClassItemOPItemPV
Dim objItemPV As ClassItemPedido
Dim dQtdFalta As Double
Dim dQtd As Double

On Error GoTo Erro_ItemOPItemPV_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202253
    Next
    
    dQtdFalta = objItemOP.dQuantidade

    iIndice = 0
    For Each objItemOPItemPV In objItemOP.colItemPV
        
        iIndice = iIndice + 1
        Set objItemPV = New ClassItemPedido

        objItemOPItemPV.lNumIntDocItemOP = objItemOP.lNumIntDoc
        objItemOPItemPV.sCodigoOP = objItemOP.sCodigo
        
        If objItemOPItemPV.dQuantidade <= dQtdFalta + QTDE_ESTOQUE_DELTA Then
            dQtd = objItemOPItemPV.dQuantidade
        Else
            dQtd = dQtdFalta
        End If
        dQtdFalta = dQtdFalta - dQtd

        lErro = Comando_Executar(alComando(1), "INSERT INTO ItemOPItemPV( NumIntDocItemOP, NumIntDocItemPV, CodigoOP, CodigoPV, " & _
                "FilialEmpresa, Produto, Quantidade, UM, Prioridade, Seq) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objItemOPItemPV.lNumIntDocItemOP, objItemOPItemPV.lNumIntDocItemPV, objItemOPItemPV.sCodigoOP, objItemOPItemPV.lCodigoPV, _
                objItemOPItemPV.iFilialEmpresa, objItemOPItemPV.sProduto, dQtd, objItemOPItemPV.sUM, objItemOPItemPV.iPrioridade, iIndice)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202254
        
        If dQtd > QTDE_ESTOQUE_DELTA Then
        
            objItemPV.iFilialEmpresa = objItemOPItemPV.iFilialEmpresa
            objItemPV.lCodPedido = objItemOPItemPV.lCodigoPV
            objItemPV.sProduto = objItemOPItemPV.sProduto
            objItemPV.iTipoOP = OP_TIPO_OP
    
            'aumenta a quantidade em OP do Pedido de Venda
            lErro = OrdemDeProducao_Grava6(alComando(2), alComando(3), objItemPV, dQtd)
            If lErro <> SUCESSO And lErro <> 41346 Then gError 202256
    
            If lErro = 41346 Then gError 202257
        
        End If

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPItemPV_Insere = SUCESSO

    Exit Function

Erro_ItemOPItemPV_Insere:

    ItemOPItemPV_Insere = gErr

    Select Case gErr

        Case 202253
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202254
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMOPITEMPV", gErr)
            
        Case 202256
        
        Case 202257
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemPV.lCodPedido, objItemPV.sProduto)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202255)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ItemOPItemPV_Exclui(ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDocItemPV As Integer
Dim objItemPV As ClassItemPedido
Dim iFilialEmpresa As Integer
Dim lCodigoPV As Long
Dim sProduto As String
Dim dQuantidade As Double

On Error GoTo Erro_ItemOPItemPV_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202267
    Next
    
    sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaItemOPItemPV
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocItemPV, FilialEmpresa, CodigoPV, Produto, Quantidade FROM ItemOPItemPV WHERE NumIntDocItemOP= ? ", _
                0, lNumIntDocItemPV, iFilialEmpresa, lCodigoPV, sProduto, dQuantidade, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 202268

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202269

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 202270
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItemOPItemPV", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 202271
        
'        Set objItemPV = New ClassItemPedido
'
'        objItemPV.iFilialEmpresa = iFilialEmpresa
'        objItemPV.lCodPedido = lCodigoPV
'        objItemPV.sProduto = sProduto
'        objItemPV.iTipoOP = OP_TIPO_OP
'
'        'aumenta a quantidade em OP do Pedido de Venda
'        lErro = OrdemDeProducao_Grava6(alComando(2), alComando(3), objItemPV, -dQuantidade)
'        If lErro <> SUCESSO And lErro <> 41346 Then gError 202274
'
'        If lErro = 41346 Then gError 202275
            
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202272
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPItemPV_Exclui = SUCESSO

    Exit Function

Erro_ItemOPItemPV_Exclui:

    ItemOPItemPV_Exclui = gErr

    Select Case gErr

        Case 202267
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202268, 202269, 202272
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPITEMPV", gErr)

        Case 202270
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ITEMOPITEMPV", gErr)

        Case 202271
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOPITEMPV", gErr)
            
        Case 202274
        
        Case 202275
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objItemPV.lCodPedido, objItemPV.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202273)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PVAndamento_Grava(ByVal objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PVAndamento_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205681

    lErro = CF("PVAndamento_Grava_EmTrans", objPVAndamento)
    If lErro <> SUCESSO Then gError 205682

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 205683

    PVAndamento_Grava = SUCESSO

    Exit Function

Erro_PVAndamento_Grava:

    PVAndamento_Grava = gErr

    Select Case gErr

        Case 205681
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205682

        Case 205683
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205684)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PVAndamento_Grava_EmTrans(ByVal objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_PVAndamento_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205686
    Next

    'Le a tabelaPVAndamento
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Auto FROM PVAndamento WHERE Codigo= ? ", _
                0, iAux, objPVAndamento.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 205687

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205688

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 205689

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PVAndamento SET Descricao= ? , Auto= ? , FatorAuto= ? ", alComando(0), _
                objPVAndamento.sDescricao, objPVAndamento.iAuto, objPVAndamento.iFatorAuto)
        If lErro <> AD_SQL_SUCESSO Then gError 205690

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO PVAndamento( Codigo, Descricao, Auto, FatorAuto) VALUES (?,?,?,?)", _
                objPVAndamento.iCodigo, objPVAndamento.sDescricao, objPVAndamento.iAuto, objPVAndamento.iFatorAuto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205691

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PVAndamento_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PVAndamento_Grava_EmTrans:

    PVAndamento_Grava_EmTrans = gErr

    Select Case gErr

        Case 205686
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205687, 205688
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PVANDAMENTO", gErr)

        Case 205689
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PVANDAMENTO", gErr)

        Case 205690
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PVANDAMENTO", gErr)

        Case 205691
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PVANDAMENTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205692)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PVAndamento_Exclui(ByVal objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PVAndamento_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205693

    lErro = CF("PVAndamento_Exclui_EmTrans", objPVAndamento)
    If lErro <> SUCESSO Then gError 205694

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 205695

    PVAndamento_Exclui = SUCESSO

    Exit Function

Erro_PVAndamento_Exclui:

    PVAndamento_Exclui = gErr

    Select Case gErr

        Case 205693
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205694

        Case 205695
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205696)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PVAndamento_Exclui_EmTrans(ByVal objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lPedido As Long, iFilialEmpresa As Integer

On Error GoTo Erro_PVAndamento_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205697
    Next

    'Le a tabela PVHistAndamento
    lErro = Comando_Executar(alComando(2), "SELECT Codigo, FilialEmpresa FROM PVHistAndamento WHERE Andamento= ? ", _
                lPedido, iFilialEmpresa, objPVAndamento.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 205786

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205787

    'Se não existir => ERRO
    If lErro <> AD_SQL_SEM_DADOS Then gError 205788

    'Le a tabela PVAndamento
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Auto FROM PVAndamento WHERE Codigo= ? ", _
                0, iAux, objPVAndamento.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 205698

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205699

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 205700

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 205701

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PVAndamento", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 205702

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PVAndamento_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PVAndamento_Exclui_EmTrans:

    PVAndamento_Exclui_EmTrans = gErr

    Select Case gErr

        Case 205697
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205698, 205699
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PVANDAMENTO", gErr)

        Case 205700
            Call Rotina_Erro(vbOKOnly, "ERRO_PVANDAMENTO_NAO_CADASTRADO", gErr, objPVAndamento.iCodigo)

        Case 205701
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PVANDAMENTO", gErr)

        Case 205702
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PVANDAMENTO", gErr)
            
        Case 205786, 205787
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PVHISTANDAMENTO", gErr)
        
        Case 205788
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PVANDAMENTO_USADO", gErr, lPedido, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205703)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PV_Grava_Andamento(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PV_Grava_Andamento

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 205789

    lErro = CF("PV_Grava_Andamento_EmTrans", objPV, gsUsuario)
    If lErro <> SUCESSO Then gError 205790

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 205791

    PV_Grava_Andamento = SUCESSO

    Exit Function

Erro_PV_Grava_Andamento:

    PV_Grava_Andamento = gErr

    Select Case gErr

        Case 205789
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 205790

        Case 205791
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205792)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PV_Grava_Andamento_EmTrans(ByVal objPV As ClassPedidoDeVenda, ByVal sUsuario As String) As Long

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocPV As Long
Dim lComando As Long
Dim sTabela As String

On Error GoTo Erro_PV_Grava_Andamento_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205793
    Next

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PVHISTANDAMENTO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 205794

    lErro = Comando_Executar(alComando(1), "INSERT INTO PVHistAndamento( NumIntDoc,Codigo, FilialEmpresa, Data, Hora, Usuario, " & _
            "Andamento, DataEntrega, OBS) VALUES (?,?,?,?,?,?,?,?,?)", _
            lNumIntDoc, objPV.lCodigo, objPV.iFilialEmpresa, Date, CDbl(Time), sUsuario, _
            objPV.iAndamento, objPV.dtDataEntrega, objPV.sOBS)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205795
    
    
    lComando = alComando(2)
    sTabela = "PedidosDeVenda"
    'Le a tabela de PV
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM " & sTabela & " WHERE Codigo= ? AND FilialEmpresa = ?", _
                0, lNumIntDocPV, objPV.lCodigo, objPV.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 205796

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205797
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lComando = alComando(3)
        sTabela = "PedidosDeVendaBaixados"
        'Le a tabela de PV Baixados
        lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM " & sTabela & " WHERE Codigo= ? AND FilialEmpresa = ?", _
                    0, lNumIntDocPV, objPV.lCodigo, objPV.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 205798
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205799
    
        If lErro = AD_SQL_SEM_DADOS Then gError 205800
    
    End If
    
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE " & sTabela & " SET Andamento = ?, DataEntrega = ?, OBS = ?", _
                lComando, objPV.iAndamento, objPV.dtDataEntrega, objPV.sOBS)
    If lErro <> AD_SQL_SUCESSO Then gError 205801

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PV_Grava_Andamento_EmTrans = SUCESSO

    Exit Function

Erro_PV_Grava_Andamento_EmTrans:

    PV_Grava_Andamento_EmTrans = gErr

    Select Case gErr

        Case 205793
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205794

        Case 205795
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PVHISTANDAMENTO", gErr)

        Case 205796, 205797
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", gErr)

        Case 205798, 205799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDABAIXADOS", gErr)

        Case 205800
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPV.lCodigo, objPV.iFilialEmpresa)

        Case 205801
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PEDIDOSDEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205802)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PV_Andamento_Auto(ByVal objPV As ClassPedidoDeVenda, Optional ByVal iTipo As Integer = PV_STATUS_ANDAMENTO_TIPO_PV_GRAVA) As Long

Dim lErro As Long
Dim objPVBD As New ClassPedidoDeVenda
Dim bMudou As Boolean, bNovo As Boolean, bReservaCompl As Boolean, bReservaComplAnt As Boolean
Dim bFaturadoTot As Boolean, bFaturadoTotAnt As Boolean, bFaturadoParc As Boolean, bFaturadoParcAnt As Boolean
Dim colPVAnd As New Collection, objItemPV As ClassItemPedido, objPVAnd As ClassPVAndamento
Dim iOrdAtual As Integer, iStatus As Integer

On Error GoTo Erro_PV_Andamento_Auto

    bMudou = False
    bNovo = False
    bReservaComplAnt = False
    bFaturadoParcAnt = False
    bFaturadoTotAnt = False
    
    If iTipo = PV_STATUS_ANDAMENTO_TIPO_PV_GRAVA Or iTipo = PV_STATUS_ANDAMENTO_TIPO_RESERVA_GRAVA Then
        
        objPVBD.lCodigo = objPV.lCodigo
        objPVBD.iFilialEmpresa = objPV.iFilialEmpresa

        lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPVBD, iStatus, False)
        If lErro <> SUCESSO And lErro <> 60304 Then gError 205803
   
        If lErro <> SUCESSO Then bNovo = True
        
        If iTipo = PV_STATUS_ANDAMENTO_TIPO_RESERVA_GRAVA Then Set objPV = objPVBD
    
        If Not bNovo Then
        
            bReservaComplAnt = True
            bFaturadoTotAnt = True
            For Each objItemPV In objPVBD.colItens
                If Abs(objItemPV.dQuantReservada - objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada) > QTDE_ESTOQUE_DELTA Then
                    bReservaComplAnt = False
                End If
                If Abs(objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada) > QTDE_ESTOQUE_DELTA Then
                    bFaturadoTotAnt = False
                End If
                If objItemPV.dQuantFaturada > QTDE_ESTOQUE_DELTA Then
                    bFaturadoParcAnt = True
                End If
            Next
        End If
    
        If objPV.dtDataEntrega <> objPVBD.dtDataEntrega Then
            bMudou = True
        End If
        objPV.sOBS = objPVBD.sOBS
        objPV.iAndamento = objPVBD.iAndamento
    
    End If
           
    bReservaCompl = True
    bFaturadoParc = False
    bFaturadoTot = True
    For Each objItemPV In objPV.colItens
        If Abs(objItemPV.dQuantReservada - objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada) > QTDE_ESTOQUE_DELTA Then
            bReservaCompl = False
        End If
        If Abs(objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada) > QTDE_ESTOQUE_DELTA Then
            bFaturadoTot = False
        End If
        If objItemPV.dQuantFaturada > QTDE_ESTOQUE_DELTA Then
            bFaturadoParc = True
        End If
    Next
   
    lErro = CF("PVAndamento_Le_Todos", colPVAnd)
    If lErro <> SUCESSO Then gError 205805
    
    For Each objPVAnd In colPVAnd
    
        If objPVAnd.iAuto = MARCADO Then
        
            Select Case objPVAnd.iFatorAuto
                
                Case PV_STATUS_ANDAMENTO_RESERVOU_TUDO
                    If bReservaCompl And ((Not bReservaComplAnt) Or bNovo Or iTipo = PV_STATUS_ANDAMENTO_TIPO_RESERVA_GRAVA) And iOrdAtual < PV_STATUS_ANDAMENTO_ORD_RESERVOU_TUDO Then
                        objPV.iAndamento = objPVAnd.iCodigo
                        iOrdAtual = PV_STATUS_ANDAMENTO_ORD_RESERVOU_TUDO
                    End If
                
                Case PV_STATUS_ANDAMENTO_NAO_RESERVOU_TUDO
                    If Not bReservaCompl And (bReservaComplAnt Or bNovo) And iOrdAtual < PV_STATUS_ANDAMENTO_ORD_NAO_RESERVOU_TUDO Then
                        objPV.iAndamento = objPVAnd.iCodigo
                        iOrdAtual = PV_STATUS_ANDAMENTO_ORD_NAO_RESERVOU_TUDO
                    End If
                    
                Case PV_STATUS_ANDAMENTO_FAT_PARCIAL
                    If bFaturadoParc And iTipo = PV_STATUS_ANDAMENTO_TIPO_NF_GRAVA And iOrdAtual < PV_STATUS_ANDAMENTO_ORD_FAT_PARCIAL Then
                        objPV.iAndamento = objPVAnd.iCodigo
                        iOrdAtual = PV_STATUS_ANDAMENTO_ORD_FAT_PARCIAL
                    End If
                    
                Case PV_STATUS_ANDAMENTO_FAT_TOTAL
                    If bFaturadoTot And iTipo = PV_STATUS_ANDAMENTO_TIPO_NF_GRAVA And iOrdAtual < PV_STATUS_ANDAMENTO_ORD_FAT_TOTAL Then
                        objPV.iAndamento = objPVAnd.iCodigo
                        iOrdAtual = PV_STATUS_ANDAMENTO_ORD_FAT_TOTAL
                    End If
            
            End Select
            
        End If
    
    Next
    
    If iTipo = PV_STATUS_ANDAMENTO_TIPO_NF_GRAVA Or iTipo = PV_STATUS_ANDAMENTO_TIPO_RESERVA_GRAVA Then
    
        lErro = CF("PV_Grava_Andamento_EmTrans", objPV, gsUsuario)
         If lErro <> SUCESSO Then gError 205804
   End If

    PV_Andamento_Auto = SUCESSO

    Exit Function

Erro_PV_Andamento_Auto:

    PV_Andamento_Auto = gErr

    Select Case gErr
    
        Case 205803 To 205805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205806)

    End Select

    Exit Function

End Function

Public Function PedidoVenda_Exporta(ByVal iFilialEmpresa As Integer, ByVal lCodigo As Long) As Long
    PedidoVenda_Exporta = SUCESSO
End Function

Public Function BloqueiosPV_Cust_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
    BloqueiosPV_Cust_Adiciona = SUCESSO
End Function

Public Function PV_CalculaPrecoUnitario(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iCodTabelaPreco As Integer, ByVal dtDataPV As Date, dPrecoUnitario As Double) As Long
    PV_CalculaPrecoUnitario = SUCESSO
End Function


Public Function BloqueioPrecoBaixo_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de "preco baixo" de acordo com a analise de margem de contribuicao

Dim lErro As Long, objMargContr As ClassMargContr, objBloqueioPV As ClassBloqueioPV
Dim objItemPV As ClassItemPedido, dFator As Double
Dim sProdutoMascarado As String, lComando As Long, sProdAux As String
Dim objAnaliseVendas As New ClassAnaliseVendaInfo

On Error GoTo Erro_BloqueioPrecoBaixo_Adiciona

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 46037

    If gobjCRFAT.iValidaDescMaxTabPreco <> 0 And objPedidoVenda.colParcelas.Count <> 0 Then
    
        Call objPedidoVenda.AnaliseVenda_Preenche(objAnaliseVendas)
        lErro = CF("AnaliseVenda_ValidaDescMax", objAnaliseVendas, False)
        If lErro <> SUCESSO And lErro <> 201561 Then gError 184250
    
        If lErro <> SUCESSO Then
        
            Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Add(giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PRECO_BAIXO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, "")
            objBloqueioPV.sObservacao = "Preço abaixo do mínimo para a tabela."
            
        End If
    
    End If
    
    lErro = CF("MargContr_InicioCalc", objPedidoVenda.iFilialEmpresaFaturamento, objMargContr, False)
    If lErro <> SUCESSO Then gError 184250
    
    For Each objItemPV In objPedidoVenda.colItensPedido
    
        sProdAux = String(STRING_PRODUTO, 0)
        lErro = Comando_Executar(lComando, "SELECT Produto FROM ProdutoCategoria WHERE Produto = ? AND Categoria = 'Linha' AND Item = '11'", sProdAux, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 201004
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201005
        
        'se nao faz parte da linha 11
        If lErro <> AD_SQL_SUCESSO Then
        
            'obtem preco do produto para o cliente
            lErro = CF("MargContr_DetCalc", objPedidoVenda.lCliente, objPedidoVenda.iFilial, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.sUnidadeMed, objPedidoVenda.iTabelaPreco)
            If lErro <> SUCESSO Then gError 184251
            
            If objItemPV.sUnidadeMed <> objItemPV.sUMEstoque Then
            
                lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 184251
                
                objMargContr.dUltPrecoCli = objMargContr.dUltPrecoCli * dFator
                
            End If
            
            If (objMargContr.dUltPrecoCli - objItemPV.dPrecoUnitario) > 0.01 Then
            
                'Adiciona o bloqueio ao Pedido
                'Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Add(giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PRECO_BAIXO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje)
                Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Add(giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PRECO_BAIXO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, "")
                
                'colocar o produto e precos na observacao
                lErro = Mascara_MascararProduto(objItemPV.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 184252
                
                objBloqueioPV.sObservacao = "produto: " & sProdutoMascarado & " no PV: " & Format(objItemPV.dPrecoUnitario, "standard") & ", do cliente: " & Format(objMargContr.dUltPrecoCli, "standard")
                
            End If
        
        End If
    
    Next
    
    Call CF("MargContr_FimCalc")
    
    Call Comando_Fechar(lComando)
    
    BloqueioPrecoBaixo_Adiciona = SUCESSO
     
    Exit Function
    
Erro_BloqueioPrecoBaixo_Adiciona:

    BloqueioPrecoBaixo_Adiciona = gErr
     
    Select Case gErr
          
        Case 184250, 184251, 184252
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184253)
     
    End Select
     
    If Not (objMargContr Is Nothing) Then Call CF("MargContr_FimCalc")
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function BloqueioPrecoDefasado_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de "preco defasado" se ele se mantem há mais de n dias

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim dtPrecoDesde As Date, objBloqueioPV As ClassBloqueioPV
Dim objItemPV As ClassItemPedido
Dim sProdutoMascarado As String, sProdAux As String
Dim bMargContrIniciado As Boolean, dFator As Double, objMargContr As ClassMargContr

On Error GoTo Erro_BloqueioPrecoDefasado_Adiciona

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184253
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 184253
    
    bMargContrIniciado = False
    
    For Each objItemPV In objPedidoVenda.colItensPedido
    
        sProdAux = String(STRING_PRODUTO, 0)
        lErro = Comando_Executar(lComando, "SELECT Produto FROM ProdutoCategoria WHERE Produto = ? AND Categoria = 'Linha' AND Item = '11'", sProdAux, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 201004
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201005
        
        'se nao faz parte da linha 11
        If lErro <> AD_SQL_SUCESSO Then
        
            '
            lErro = CF("ItemPV_ObtemDtPrecoDesde", objPedidoVenda, objItemPV, dtPrecoDesde, lComando1)
            If lErro <> SUCESSO Then gError 184254
            
            If (gdtDataAtual - dtPrecoDesde) > 120 Then
            
                If bMargContrIniciado = False Then
        
                    lErro = CF("MargContr_InicioCalc", objPedidoVenda.iFilialEmpresaFaturamento, objMargContr, False)
                    If lErro <> SUCESSO Then gError 184254
                    
                    bMargContrIniciado = True
                
                End If
                
                'obtem preco do produto para o cliente
                lErro = CF("MargContr_DetCalc", objPedidoVenda.lCliente, objPedidoVenda.iFilial, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.sUnidadeMed, objPedidoVenda.iTabelaPreco)
                If lErro <> SUCESSO Then gError 184254
                
                If objItemPV.sUnidadeMed <> objItemPV.sUMEstoque Then
                
                    lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.sUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 184254
                    
                    objMargContr.dUltPrecoCli = objMargContr.dUltPrecoCli * dFator
                    
                End If
                
                'Adiciona o bloqueio ao Pedido
                'Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Add(giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PRECO_DEFASADO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje)
                Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Add(giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_PRECO_DEFASADO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, "")
                
                'colocar o produto e data na observacao
                lErro = Mascara_MascararProduto(objItemPV.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 184255
                
                objBloqueioPV.sObservacao = "prod.: " & sProdutoMascarado & " s/reaj. desde: " & Format(dtPrecoDesde, "dd/mm/yy") & " (PV " & Format(objItemPV.dPrecoUnitario, "standard") & " Cli " & Format(objMargContr.dUltPrecoCli, "standard") & ")"
                
            End If
        
        End If
        
    Next
    
    If bMargContrIniciado Then Call CF("MargContr_FimCalc")

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    BloqueioPrecoDefasado_Adiciona = SUCESSO
     
    Exit Function
    
Erro_BloqueioPrecoDefasado_Adiciona:

    BloqueioPrecoDefasado_Adiciona = gErr
     
    Select Case gErr
          
        Case 184254, 184255
        
        Case 184253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184256)
     
    End Select
     
    If bMargContrIniciado Then Call CF("MargContr_FimCalc")
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
     
    Exit Function
    
End Function


Public Function OrcamentoVendaHist_Exclui(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Exclui do BD as Parcelas do Orcamento

Dim lErro As Long
Dim iIndice As Integer
Dim iVersao As Integer
Dim lNumIntDocItem As Long, lNumIntDocGrade As Long, lNumIntDocOV As Long
Dim alComando(1 To 8) As Long, objItemOV As ClassItemOV

On Error GoTo Erro_OrcamentoVendaHist_Exclui

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206831
    Next

    'Le as versões
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Versao FROM OrcamentoVendaHist WHERE Codigo = ? AND FilialEmpresa = ?", 0, lNumIntDocOV, iVersao, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 206832

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206833

    'Para cada Versao
    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_OVHIST, lNumIntDocOV)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_OVHIST, lNumIntDocOV)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("InfoAdicionais_Exclui", TIPODOC_INFOADIC_OVHIST, lNumIntDocOV)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM OrcamentoVendaHist", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 206834
           
        'Busca a Próxima
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206835

    Loop
    
    'Le os itens
    lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM ItensOrcamentoVendaHist WHERE CodOrcamento = ? AND FilialEmpresa = ?", 0, lNumIntDocItem, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 206836

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206837

    'Para cada item
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui a parcela da tabela de parcelas de Orcamento de Venda
        lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ItensOrcamentoVendaHist", alComando(3))
        If lErro <> AD_SQL_SUCESSO Then gError 206838
        
        Set objItemOV = New ClassItemOV
        
        objItemOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOVHIST
        objItemOV.objInfoAdicDocItem.lNumIntDocItem = lNumIntDocItem

        lErro = CF("InfoAdicDocItem_Exclui", objItemOV.objInfoAdicDocItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Le a grade do item
        lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc FROM ItensOrcamentoVendaGradeHist WHERE NumIntItemOV = ?", 0, lNumIntDocGrade, lNumIntDocItem)
        If lErro <> AD_SQL_SUCESSO Then gError 206839
    
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206840
        
        'Para cada grade
        Do While lErro <> AD_SQL_SEM_DADOS
    
            'Exclui a parcela da tabela de parcelas de Orcamento de Venda
            lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM ItensOrcamentoVendaGradeHist", alComando(5))
            If lErro <> AD_SQL_SUCESSO Then gError 206841
        
            'Busca a Próxima
            lErro = Comando_BuscarProximo(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206842
    
        Loop
    
        'Busca a Próxima
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206843

    Loop
    
    'Le as versões
    lErro = Comando_ExecutarPos(alComando(7), "SELECT Versao FROM ParcelasOVHist WHERE OrcamentoVenda = ? AND FilialEmpresa = ?", 0, iVersao, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 206844

    lErro = Comando_BuscarPrimeiro(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206845

    'Para cada Versao
    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(alComando(8), "DELETE FROM ParcelasOVHist", alComando(7))
        If lErro <> AD_SQL_SUCESSO Then gError 206846
   
        'Busca a Próxima
        lErro = Comando_BuscarProximo(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206847

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrcamentoVendaHist_Exclui = SUCESSO

    Exit Function

Erro_OrcamentoVendaHist_Exclui:

    OrcamentoVendaHist_Exclui = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206831
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206832, 206833, 206835
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 206834
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 206836, 206837, 206843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 206838
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 206839, 206840, 206842
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORCAMENTOVENDAGRADE", gErr, objOrcamentoVenda.lCodigo)

        Case 206841
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSORCAMENTOVENDAGRADE", gErr, objOrcamentoVenda.lCodigo)

        Case 206844, 206845, 206847
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASOV", gErr, objOrcamentoVenda.lCodigo)

        Case 206846
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASOV", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206848)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PV_Grava_AnalisaCliente(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'impede a gravacao de pv de cliente em atraso

Dim lErro As Long
Dim lComando As Long
Dim dtDataLimite As Date
Dim lNumIntDoc As Long

On Error GoTo Erro_PV_Grava_AnalisaCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94359
    
     'Alterado por Cyntia em 21/05/02
    lErro = CF("Dias_Uteis_Antes", gdtDataHoje, 1, dtDataLimite)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objPedidoVenda.iFilialEmpresa > 50 Then

        'Seleciona parcelasrec do Cliente recebido em objPedidoVenda que esteja em atraso de acordo com os dias máximos p/ atraso definido p/ Bloqueios por atraso de pagamento.
        lErro = Comando_Executar(lComando, "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND ParcelasRec.DataVencimentoReal <= ? AND TitulosRec.Cliente = ? AND TitulosRec.FilialEmpresa > 50", lNumIntDoc, STATUS_EXCLUIDO, STATUS_BAIXADO, dtDataLimite, objPedidoVenda.lCliente)
        If lErro <> SUCESSO Then gError 94366
    
    Else
    
        'Seleciona parcelasrec do Cliente recebido em objPedidoVenda que esteja em atraso de acordo com os dias máximos p/ atraso definido p/ Bloqueios por atraso de pagamento.
        lErro = Comando_Executar(lComando, "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND ParcelasRec.DataVencimentoReal <= ? AND TitulosRec.Cliente = ?", lNumIntDoc, STATUS_EXCLUIDO, STATUS_BAIXADO, dtDataLimite, objPedidoVenda.lCliente)
        If lErro <> SUCESSO Then gError 94366
    
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94367
    
    'se existe pelo menos uma parcela em atraso
    If lErro = SUCESSO Then gError 201211
       
    Call Comando_Fechar(lComando)
    
    PV_Grava_AnalisaCliente = SUCESSO
     
    Exit Function
    
Erro_PV_Grava_AnalisaCliente:

    PV_Grava_AnalisaCliente = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 201211 'a gravacao nao será permitida
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATRASO_PAGTO_IMPEDE_PV", gErr)
        
        Case 94366, 94367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
        
        Case 94359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152623)
     
    End Select
     
    Call Comando_Fechar(lComando)
     
    Exit Function
    
End Function

Public Function Importa_Dados_Cliente_Arq_Funcao_Cust(sNomeFuncao As String) As Long
    sNomeFuncao = ""
    Importa_Dados_Cliente_Arq_Funcao_Cust = SUCESSO
End Function

Public Function Importa_Dados_PV_Arq_Funcao_Cust(sNomeFuncao As String) As Long
    sNomeFuncao = ""
    Importa_Dados_PV_Arq_Funcao_Cust = SUCESSO
End Function

Function OrcamentoVenda_Atualiza_Envio(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lComando As Long, sTabela As String
Dim lTransacao As Long
Dim iIndice As Integer
Dim iVersao As Integer

On Error GoTo Erro_OrcamentoVenda_Atualiza_Envio

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211470
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211471
    Next
        
    If objOrcamentoVenda.iHistorico <> MARCADO Then
    
        lComando = alComando(0)
        sTabela = "OrcamentoVenda"
    
        lErro = Comando_ExecutarPos(lComando, "SELECT Versao FROM " & sTabela & " WHERE FilialEmpresa = ? AND Codigo = ?", 0, iVersao, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 211472
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211473
    
        'se nao encontrou => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 211474
    
    Else

        lComando = alComando(1)
        sTabela = "OrcamentoVendaHist"
    
        lErro = Comando_ExecutarPos(lComando, "SELECT Versao FROM " & sTabela & " WHERE FilialEmpresa = ? AND Codigo = ? AND Versao = ? ", 0, iVersao, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iVersao)
        If lErro <> AD_SQL_SUCESSO Then gError 211475
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211476
    
        'se nao encontrou => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 211477

    End If

    lErro = Comando_ExecutarPos(alComando(2), "UPDATE " & sTabela & " SET DataEnvio = ?", lComando, objOrcamentoVenda.dtDataEnvio)
    If lErro <> AD_SQL_SUCESSO Then gError 211478

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211479

    OrcamentoVenda_Atualiza_Envio = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Atualiza_Envio:

    OrcamentoVenda_Atualiza_Envio = gErr

    Select Case gErr

        Case 211470
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211471
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211472, 211473, 211475, 211476
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 211474, 211477
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)
        
        Case 211478
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORCAMENTOVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case 211479
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211480)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Private Function OrdemDeProducao_Grava42(ByVal objOrdemDeProducao As ClassOrdemDeProducao, ByVal objItemOP As ClassItemOP, tItemOP As typeItemOP) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 20) As Long
Dim objItemPV As New ClassItemPedido
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objEmpenho As ClassEmpenho
Dim colEmpenhoAnt As New Collection
Dim sProduto As String, iAlmoxarifado As Integer, dQuantEmpenhada As Double, dQuantRequisitada As Double
Dim dSaldo As Double, lNumIntDocItem As Long, lEmpenho As Long, objProduto As New ClassProduto

On Error GoTo Erro_OrdemDeProducao_Grava42

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211729
    Next
    
    'Se alterou a quantidade e não baixou
    'Tem que acertar a quantidade empenhada mantendo a quantidade requisitada
    If objItemOP.iPossuiGrade = DESMARCADO And objItemOP.iTipo = OP_TIPO_OP And Abs(objItemOP.dQuantidade - tItemOP.dQuantidade) > QTDE_ESTOQUE_DELTA And objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then

        sProduto = String(STRING_PRODUTO, 0)

        'EXCLUI OS EMPENHOS
        'tenta seleciona empenhos associados a um Item da OP
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Produto, Almoxarifado, Quantidade, QuantidadeRequisitada FROM Empenho WHERE FilialEmpresa = ? AND NumIntDocItemOP = ?", 0, sProduto, iAlmoxarifado, dQuantEmpenhada, dQuantRequisitada, objItemOP.iFilialEmpresa, objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 211730
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211731
    
        Do While lErro <> AD_SQL_SEM_DADOS
    
            'exclui empenho
            lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Empenho ", alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 211732
            
            Set objEstoqueProduto = New ClassEstoqueProduto
            Set objEmpenho = New ClassEmpenho
            colEmpenhoAnt.Add objEmpenho
            objEmpenho.sProduto = sProduto
            objEmpenho.iAlmoxarifado = iAlmoxarifado
            objEmpenho.dQuantidade = dQuantEmpenhada
            objEmpenho.dQuantidadeRequisitada = dQuantRequisitada
    
            objEstoqueProduto.sProduto = sProduto
            objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
    
            dSaldo = dQuantRequisitada - dQuantEmpenhada
            
            'Só abate do saldo empenhado quando a qtd requisitada for menor que
            'a empenhada
            If dSaldo > 0 Then dSaldo = 0
            
            lErro = CF("EstoqueProduto_AlterarEmpenho", objEstoqueProduto, dSaldo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211733
    
        Loop

        'REFAZ OS EMPENHOS
        'cria os empenhos
        
        'obtem o proximo codigo para empenhos e para NumIntItemOP
        lErro = OrdemDeProducao_Grava3(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objProduto.sCodigo = objItemOP.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
        objItemOP.iClasseUM = objProduto.iClasseUM
    
        lErro = CF("OrdemDeProducao_Grava2", objItemOP, alComando, objOrdemDeProducao.dtDataEmissao, lEmpenho)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'atualiza o proximo codigo para lNumIntDocItem e lEmpenho
        lErro = OrdemDeProducao_Grava5(alComando(12), alComando(13), lNumIntDocItem, lEmpenho)
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        sProduto = String(STRING_PRODUTO, 0)
        
        'ACERTA A QUANTIDADE REQUISITADA
        'tenta seleciona empenhos associados a um Item da OP
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Produto, Almoxarifado, Quantidade, QuantidadeRequisitada FROM Empenho WHERE FilialEmpresa = ? AND NumIntDocItemOP = ?", 0, sProduto, iAlmoxarifado, dQuantEmpenhada, dQuantRequisitada, objItemOP.iFilialEmpresa, objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 211734
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211735
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            For Each objEmpenho In colEmpenhoAnt
                If objEmpenho.sProduto = sProduto Then
                    If objEmpenho.dQuantidadeRequisitada > 0 Then
                        'Altera
                        lErro = Comando_ExecutarPos(alComando(3), "UPDATE Empenho SET QuantidadeRequisitada = ? ", alComando(1), objEmpenho.dQuantidadeRequisitada)
                        If lErro <> AD_SQL_SUCESSO Then gError 211736
                    
                        If objEmpenho.dQuantidadeRequisitada >= dQuantEmpenhada Then
                            dSaldo = -dQuantEmpenhada
                        Else
                            dSaldo = -objEmpenho.dQuantidadeRequisitada
                        End If
            
                        objEstoqueProduto.sProduto = sProduto
                        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
                        lErro = CF("EstoqueProduto_AlterarEmpenho", objEstoqueProduto, dSaldo)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    End If
                    Exit For
                End If
            Next

            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211737
    
        Loop
    
    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava42 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava42:

    OrdemDeProducao_Grava42 = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 211729
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211730 To 211737
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211738)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoVenda_Grava_Parc(ByVal objPVOrig As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPV As ClassPedidoDeVenda
Dim objItem As ClassItemPedido
Dim objItemOrig As ClassItemPedido
Dim iItem As Integer, lNumPedido As Long
Dim objTelaPedido As Object
Dim objTelaPedidoOCX As Object
Dim bLoad As Boolean, iCancel As Integer
Dim objPVNovo As ClassPedidoDeVenda

On Error GoTo Erro_PedidoVenda_Grava_Parc

    Set objTelaPedidoOCX = CreateObject("TelasFAT.PedidoVendaOCX")
    Set objTelaPedido = objTelaPedidoOCX.objCTtela

    bLoad = False
    Call objTelaPedidoOCX.Form_Load
    bLoad = True
    Call objTelaPedido.Form_Activate
    
    objTelaPedido.iImportacao = MARCADO
    
    'TEM QUE EXCLUIR TODAS QUE JÁ EXISTEM PARA O CÓDIGO BASE
    lErro = CF("PedidoVendaParc_Exclui", objPVOrig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'RECRIA ALTERANDO A TRIBUTAÇÃO/CÓDIGO
    For iIndice = 2 To objPVOrig.iParc

        Set objPV = New ClassPedidoDeVenda
        Set objPVNovo = New ClassPedidoDeVenda
               
        lErro = CF("Config_ObterAutomatico_EmTrans", "FatConfig", "NUM_PROX_CODIGO_PEDVENDA", "PedVenTodos", "Codigo", lNumPedido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objPV.lCodigo = objPVOrig.lCodigo
        objPV.iFilialEmpresa = objPVOrig.iFilialEmpresa
        
        'Traz o pedido original para tela
        lErro = objTelaPedido.Traz_Pedido_Tela(objPV)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Tira a tributação e recalcula
        objPV.objTributacao.iTipoTributacao = 32
        objPV.objTributacao.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        
        'objTelaPedido.Codigo.promptinclude = False
        objTelaPedido.Codigo.Text = CStr(lNumPedido)
        'objTelaPedido.Codigo.promptinclude = True
        
        lErro = objTelaPedido.ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Obtém como o pedid ficou em um novo obj
        lErro = objTelaPedido.Move_Pedido_Memoria(objPVNovo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objPVNovo.lCodigoBase = objPVOrig.lCodigo
        
        Set objPVNovo.colBloqueiosPV = New ColBloqueioPV
        
        lErro = CF("PedidoVenda_Grava_EmTrans", objPVNovo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next
    
    Call objTelaPedidoOCX.Form_Unload(iCancel)
    bLoad = False

    Set objTelaPedidoOCX = Nothing
    Set objTelaPedido = Nothing

    PedidoVenda_Grava_Parc = SUCESSO

    Exit Function

Erro_PedidoVenda_Grava_Parc:

    PedidoVenda_Grava_Parc = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211763)

    End Select
    
    If bLoad Then Call objTelaPedidoOCX.Form_Unload(iCancel)
    
    Set objTelaPedidoOCX = Nothing
    Set objTelaPedido = Nothing

    Exit Function

End Function

Function PedidoVendaParc_Exclui(ByVal objPVOrig As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long, lCodigo As Long
Dim objPV As ClassPedidoDeVenda

On Error GoTo Erro_PedidoVendaParc_Exclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211758
    Next
    
    If objPVOrig.lCodigo <> 0 Then
    
        'Lê o Pedido com o Código e a filialempresa passados
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Codigo FROM PedidosDeVenda WHERE FilialEmpresa = ? AND CodigoBase = ?", 0, lNumIntDoc, lCodigo, objPVOrig.iFilialEmpresa, objPVOrig.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 211759
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211760
         
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objPV = New ClassPedidoDeVenda
            
            objPV.lNumIntDoc = lNumIntDoc
            objPV.lCodigo = lCodigo
            objPV.iFilialEmpresa = objPVOrig.iFilialEmpresa
    
            lErro = CF("PedidoVenda_Exclui_Trans", objPV)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211761
            
        Loop
        
    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoVendaParc_Exclui = SUCESSO

    Exit Function

Erro_PedidoVendaParc_Exclui:

    PedidoVendaParc_Exclui = gErr

    Select Case gErr

        Case 211758
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211759 To 211761
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", gErr, objPVOrig.lCodigo)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211762)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OrdemDeProducao_Grava31(ByVal lComando1 As Long, lEmpenho As Long) As Long
'obtem o proximo codigo para lEmpenho

Dim lErro As Long, sConteudo As String
Dim iIndice As Integer

On Error GoTo Erro_OrdemDeProducao_Grava31

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(lComando1, "SELECT Conteudo FROM MatConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, "NUM_PROX_EMPENHO", giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 22959

    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 22960

    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 22961

    lEmpenho = CLng(sConteudo)

    OrdemDeProducao_Grava31 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava31:

    OrdemDeProducao_Grava31 = Err

    Select Case Err

        Case 22959, 22960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", Err, "NUM_PROX_EMPENHO")

        Case 22961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG", Err, "NUM_PROX_EMPENHO")

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152593)

    End Select

    Exit Function

End Function

Private Function OrdemDeProducao_Grava51(ByVal lComando1 As Long, lEmpenho As Long) As Long
'atualiza numeração de lNumIntDocItem e lEmpenho

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_OrdemDeProducao_Grava51

    'Abre novos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 22943
    Next

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE MATConfig SET Conteudo = ?", lComando1, lEmpenho)
    If lErro <> AD_SQL_SUCESSO Then Error 22980

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava51 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava51:

    OrdemDeProducao_Grava51 = Err

    Select Case Err

        Case 22943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", Err, "NUM_PROX_EMPENHO")

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152592)

    End Select

    'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OrdemDeProducao_Grava7(ByVal lNumIntDocItemOP As Long) As Long
'ajusta qtde requisitada de empenhos de itens de op que acabou de ser reativada

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iFilialEmpresa As Integer, lCodigo As Long, sProduto As String, iAlmoxarifado As Integer
Dim dQtdRequisitada As Double
Dim dQtdAtualizar As Double
Dim dQtdMov As Double, dQtdMovEst As Double, dFator As Double
Dim dQtdEmpenho As Double, iTipoMov As Integer, objProduto As New ClassProduto
Dim objEstoqueProduto As New ClassEstoqueProduto, sSiglaUM As String

On Error GoTo Erro_OrdemDeProducao_Grava7

    'Abre novos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201456
    Next

    sProduto = String(STRING_PRODUTO, 0)
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa, Codigo, Produto, Almoxarifado, Quantidade FROM Empenho WHERE NumIntDocItemOP = ? ORDER BY Codigo", 0, _
        iFilialEmpresa, lCodigo, sProduto, iAlmoxarifado, dQtdEmpenho, lNumIntDocItemOP)
    If lErro <> AD_SQL_SUCESSO Then gError 201457
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201458

    Do While lErro <> AD_SQL_SEM_DADOS
    
        dQtdRequisitada = 0
    
        objProduto.sCodigo = sProduto
            
        lErro = CF("Produto_le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        sSiglaUM = String(STRING_UM_SIGLA, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT SiglaUM, Quantidade, TipoMov FROM MovimentoEstoque WHERE NumIntDocEst = 0 AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND Produto = ? AND TipoMov IN (?,?) ORDER BY NumIntDoc", _
            sSiglaUM, dQtdMov, iTipoMov, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP, sProduto, MOV_EST_REQ_PRODUCAO, MOV_EST_DEV_MATERIAL_PRODUCAO)
        If lErro <> AD_SQL_SUCESSO Then gError 201459
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201460
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            If UCase(sSiglaUM) = UCase(objProduto.sSiglaUMEstoque) Then
                dQtdMovEst = dQtdMov
            Else
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                dQtdMovEst = Arredonda_Estoque(dQtdMov * dFator)
            
            End If
            
            Select Case iTipoMov
            
                Case MOV_EST_REQ_PRODUCAO
                    dQtdRequisitada = Arredonda_Estoque(dQtdRequisitada + dQtdMovEst)
                
                Case MOV_EST_DEV_MATERIAL_PRODUCAO
                    dQtdRequisitada = Arredonda_Estoque(dQtdRequisitada - dQtdMovEst)
                
            End Select
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201461
        
        Loop
        
        If dQtdRequisitada <> 0 Then
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE Empenho SET QuantidadeRequisitada = ?", alComando(0), dQtdRequisitada)
            If lErro <> AD_SQL_SUCESSO Then gError 201462
        
            If dQtdEmpenho > dQtdRequisitada Then
                dQtdAtualizar = -dQtdRequisitada
            Else
                dQtdAtualizar = -dQtdEmpenho
            End If
            
            objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
            objEstoqueProduto.sProduto = sProduto
            
            'Ajustar a qtde do empenho
            lErro = CF("EstoqueProduto_AlterarEmpenho", objEstoqueProduto, dQtdAtualizar)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201463
    
    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Grava7 = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava7:

    OrdemDeProducao_Grava7 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201457 To 201463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP7", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152592)

    End Select

    'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function BloqueioMargemBaixa_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de "preco baixo" de acordo com a analise de margem de contribuicao

Dim lErro As Long, objMargContr As ClassMargContr, objBloqueioPV As ClassBloqueioPV
Dim objItemPV As ClassItemPedido, dFator As Double, dPrecoVenda As Double
Dim sProdutoMascarado As String, lComando As Long, sProdAux As String
Dim objAnaliseVendas As New ClassAnaliseVendaInfo, dPercMarg As Double, dFrete As Double

On Error GoTo Erro_BloqueioMargemBaixa_Adiciona

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 46037

    lErro = CF("MargContr_InicioCalc", objPedidoVenda.iFilialEmpresaFaturamento, objMargContr, False)
    If lErro <> SUCESSO Then gError 184250
    
    For Each objItemPV In objPedidoVenda.colItensPedido
    
        sProdAux = String(STRING_PRODUTO, 0)
        lErro = Comando_Executar(lComando, "SELECT Produto FROM ProdutoCategoria WHERE Produto = ? AND Categoria = 'Linha' AND Item = '11'", sProdAux, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 201004
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201005
        
        'se nao faz parte da linha 11
        If lErro <> AD_SQL_SUCESSO Then
        
            dPrecoVenda = objItemPV.dPrecoUnitario
            dFrete = objItemPV.objTributacaoItemPV.dValorFreteItem / IIf(objItemPV.dQuantidade <> 0, objItemPV.dQuantidade, 1)
            dPrecoVenda = dPrecoVenda + dFrete
            'obtem a margem para o preço do itempv
            lErro = CF("MargContr_DetCalcMarg", dPercMarg, objItemPV.dQuantidade, objPedidoVenda.lCliente, objPedidoVenda.iFilial, objPedidoVenda.iFilialEmpresaFaturamento, objItemPV.sProduto, objPedidoVenda.iTabelaPreco, objItemPV.sUnidadeMed, objPedidoVenda.iCondicaoPagto, objItemPV.objTributacaoItemPV.sNaturezaOp, dPrecoVenda, dFrete)
            If lErro <> SUCESSO Then gError 184251
            
            If dPercMarg < 0.15 Then
            
                'Adiciona o bloqueio ao Pedido
                Set objBloqueioPV = objPedidoVenda.colBloqueiosPV.Add(giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_MARGEM_BAIXA, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, "")
                
                'colocar o produto e precos na observacao
                lErro = Mascara_MascararProduto(objItemPV.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 184252
                
                objBloqueioPV.sObservacao = "produto: " & sProdutoMascarado & " no PV: " & Format(objItemPV.dPrecoUnitario, "standard") & ", margem: " & Format(dPercMarg, "standard")
                
            End If
        
        End If
    
    Next
    
    Call CF("MargContr_FimCalc")
    
    Call Comando_Fechar(lComando)
    
    BloqueioMargemBaixa_Adiciona = SUCESSO
     
    Exit Function
    
Erro_BloqueioMargemBaixa_Adiciona:

    BloqueioMargemBaixa_Adiciona = gErr
     
    Select Case gErr
          
        Case 184250, 184251, 184252
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184253)
     
    End Select
     
    If Not (objMargContr Is Nothing) Then Call CF("MargContr_FimCalc")
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

