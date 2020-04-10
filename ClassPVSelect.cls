VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'??? talvez devesse estar em grava e nao em select. Jones 16/01/2002
Function Reserva_Automatico1(ByVal lComando As Long, lCodigo As Long, ByVal iFilialEmpresa As Integer) As Long
'descobre o próximo código de reserva disponível e loca-o
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Reserva_Automatico1

    sConteudo = String(STRING_CONTEUDO, 0)
    
    'Pesquisa o número da Reserva gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, NUM_PROX_RESERVA, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 36674

    'le o registro com o número da Reserva automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36675

    'loca o registro que guarda o numero da Reserva gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36676

    lCodigo = CLng(sConteudo)

    Reserva_Automatico1 = SUCESSO

    Exit Function

Erro_Reserva_Automatico1:

    Reserva_Automatico1 = Err

    Select Case Err

        Case 36674, 36675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", Err, sConteudo)
    
        Case 36676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CPRCONFIG", Err, sConteudo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152676)

    End Select

    Exit Function

End Function

'??? inutil e errada pois nao considera a filialempresa
'Function Reserva_Automatico2(lComando1 As Long, lCodigo As Long) As Long
''descobre o proximo codigo de reserva disponivel
'
'Dim lErro As Long
'Dim lCodigo_Reserva As Long
'
'On Error GoTo Erro_Reserva_Automatico2
'
'    'verifica se este número ainda não foi utilizado
'    'Pesquisa as numerações de Reservas maiores ou iguais a numeração gerada automaticamente
'    lErro = Comando_Executar(lComando1, "SELECT Codigo FROM Reserva WHERE Codigo >= ? ORDER BY Codigo", lCodigo_Reserva, lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then Error 36679
'
'    'le o primeiro codigo >= ao codigo automatico
'    lErro = Comando_BuscarPrimeiro(lComando1)
'    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36680
'
'    'pesquisa o proximo número de codigo disponível
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        If lCodigo_Reserva < lCodigo Then Exit Do
'
'        If lCodigo_Reserva = lCodigo Then
'            lCodigo = lCodigo + 1
'        End If
'
'        lErro = Comando_BuscarProximo(lComando1)
'        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36681
'
'    Loop
'
'    Reserva_Automatico2 = SUCESSO
'
'    Exit Function
'
'Erro_Reserva_Automatico2:
'
'    Reserva_Automatico2 = Err
'
'    Select Case Err
'
'        Case 36679, 36680, 36681
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152677)
'
'    End Select
'
'    Exit Function
'
'End Function

'mario
Function ItemPV_ObtemQtdeReservaAlmox(ByVal objItemPV As ClassItemPedido, ByVal iAlmoxarifado As Integer, dQuantReserva As Double) As Long
'preenche em dQuantReserva a qtde reservada do item no almoxarifado

Dim lErro As Long
Dim objReserva As ClassReserva

On Error GoTo Erro_ItemPV_ObtemQtdeReservaAlmox

    dQuantReserva = 0

    For Each objReserva In objItemPV.ColReserva

        If objReserva.iAlmoxarifado = iAlmoxarifado Then

            dQuantReserva = objReserva.dQuantidade
            Exit For

        End If

    Next

    ItemPV_ObtemQtdeReservaAlmox = SUCESSO

    Exit Function

Erro_ItemPV_ObtemQtdeReservaAlmox:

    ItemPV_ObtemQtdeReservaAlmox = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152678)

    End Select

    Exit Function

End Function

Function ReservasItem_Le(ByVal objItemPedido As ClassItemPedido, ByVal colReservaBD As colReservaItem) As Long
'Lê nas tabelas Reserva e Almoxarifado as reservas do item

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim tReserva As typeReserva
Dim sAlmoxarifado As String
Dim iTipoOrigem As Integer

On Error GoTo Erro_ReservasItem_Le

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23720

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 23721
        
    If objItemPedido.iPossuiGrade = MARCADO Then
        iTipoOrigem = TIPO_PEDIDO_GRADE
    Else
        iTipoOrigem = TIPO_PEDIDO_COD
    End If
        
    With tReserva
    
        sAlmoxarifado = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sResponsavel = String(STRING_NOME, 0)  'Responsável pela reserva
    
        'Verifica Reserva do Item da própria filial
        lErro = Comando_Executar(lComando, "SELECT Almoxarifado, Quantidade, NomeReduzido, Responsavel, DataValidade FROM Almoxarifado, Reserva WHERE Almoxarifado.Codigo = Reserva.Almoxarifado AND Reserva.Produto = ? AND Reserva.DocOrigem = ? AND Reserva.FilialEmpresa = ? AND Reserva.TipoDoc = ?", _
            .iAlmoxarifado, .dQuantidade, sAlmoxarifado, .sResponsavel, .dtDataValidade, objItemPedido.sProduto, objItemPedido.lCodPedido, giFilialEmpresa, iTipoOrigem)
        If lErro <> AD_SQL_SUCESSO Then Error 23722
    
        'Tenta ler Reserva da própria filial
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23724
    
        'Coloca na coleção caso não esteja vazio
        Do While lErro <> AD_SQL_SEM_DADOS
    
            colReservaBD.Add .iAlmoxarifado, .dQuantidade, .sResponsavel, sAlmoxarifado, .dtDataValidade, 0
    
            'le o proximo registro da tabela
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23725
    
        Loop
    
        sAlmoxarifado = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sResponsavel = String(STRING_NOME, 0)  'Responsável pela reserva
        
        'Verifica Reserva do restante da empresa
        lErro = Comando_Executar(lComando2, "SELECT Almoxarifado, Quantidade, NomeReduzido, Responsavel, DataValidade FROM Almoxarifado, Reserva WHERE Almoxarifado.Codigo=Reserva.Almoxarifado AND Reserva.Produto=?  AND Reserva.DocOrigem=? AND Reserva.FilialEmpresa<>? AND Reserva.TipoDoc=1  ORDER BY Reserva.FilialEmpresa, Almoxarifado.NomeReduzido", _
        .iAlmoxarifado, .dQuantidade, sAlmoxarifado, .sResponsavel, .dtDataValidade, objItemPedido.sProduto, objItemPedido.lCodPedido, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 23723
    
        'Tenta ler Reserva do restante da empresa
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23726
    
        Do While lErro <> AD_SQL_SEM_DADOS
    
            colReservaBD.Add .iAlmoxarifado, .dQuantidade, .sResponsavel, sAlmoxarifado, .dtDataValidade, 0
    
            'le o proximo registro da tabela
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23727
    
        Loop
        
    End With
    
    If colReservaBD.Count = 0 Then Error 30099

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    ReservasItem_Le = SUCESSO

    Exit Function

Erro_ReservasItem_Le:

    ReservasItem_Le = Err

    Select Case Err

        Case 23720, 23721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 23722, 23723, 23724, 23725, 23726, 23727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVAITEMBD", Err)

        Case 30099
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_RESERVAS", Err, objItemPedido.lCodPedido, objItemPedido.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152679)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function ItemPedido_Le(ByVal objItemPedido As ClassItemPedido) As Long
'Pesquisa no Banco de Dados a ItemPedido cujo código do Pedido de Venda e o ítem foram passados no objeto

Dim lErro As Long
Dim lComando As Long
Dim tItemPedido As typeItemPedido

On Error GoTo Erro_ItemPedido_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23968

    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tItemPedido.sDescricao = String(STRING_DESCRICAO_ITEM, 0)
    tItemPedido.sLote = String(STRING_LOTE_ITEM, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)

    'Pesquisa no Banco de Dados o Item Pedido
    If objItemPedido.lNumIntDoc = 0 Then
        lErro = Comando_Executar(lComando, "SELECT CodPedido , Produto, QuantOP, QuantSC, PercDesc1, TipoDesc1, PercDesc2, TipoDesc2, PercDesc3, TipoDesc3, FilialEmpresa, Quantidade, PrecoUnitario, PrecoTotal, ClasseUM, UnidadeMed, ValorDesconto, DataEntrega , Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantFaturada, QuantReservada, NumIntDoc, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais,Prioridade, Comissao, TabelaPreco FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? and CodPedido = ? AND Produto = ?", _
            tItemPedido.lCodPedido, tItemPedido.sProduto, tItemPedido.dQuantOP, tItemPedido.dQuantSC, tItemPedido.dPercDesc1, tItemPedido.iTipoDesc1, tItemPedido.dPercDesc2, tItemPedido.iTipoDesc2, tItemPedido.dPercDesc3, tItemPedido.iTipoDesc3, tItemPedido.iFilialEmpresa, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus _
            , tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.dQuantReservada, tItemPedido.lNumIntDoc, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRecebForaFaixa, tItemPedido.dQuantFaturadaAMais, tItemPedido.iPrioridade, tItemPedido.dComissao, tItemPedido.iTabelaPreco, objItemPedido.iFilialEmpresa, objItemPedido.lCodPedido, objItemPedido.sProduto)
    Else
        lErro = Comando_Executar(lComando, "SELECT CodPedido , Produto, QuantOP, QuantSC, PercDesc1, TipoDesc1, PercDesc2, TipoDesc2, PercDesc3, TipoDesc3, FilialEmpresa, Quantidade, PrecoUnitario, PrecoTotal, ClasseUM, UnidadeMed, ValorDesconto, DataEntrega , Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantFaturada, QuantReservada, NumIntDoc, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco FROM ItensPedidoDeVenda WHERE NumIntDoc = ?", _
            tItemPedido.lCodPedido, tItemPedido.sProduto, tItemPedido.dQuantOP, tItemPedido.dQuantSC, tItemPedido.dPercDesc1, tItemPedido.iTipoDesc1, tItemPedido.dPercDesc2, tItemPedido.iTipoDesc2, tItemPedido.dPercDesc3, tItemPedido.iTipoDesc3, tItemPedido.iFilialEmpresa, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus _
            , tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.dQuantReservada, tItemPedido.lNumIntDoc, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRecebForaFaixa, tItemPedido.dQuantFaturadaAMais, tItemPedido.iPrioridade, tItemPedido.dComissao, tItemPedido.iTabelaPreco, objItemPedido.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 23969

    'Verifica se há Item Pedido
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23970

    'Não encontrou o Item Pedido
    If lErro = AD_SQL_SEM_DADOS Then Error 23971

    'Preenche objItemPedido
    With objItemPedido

        .iFilialEmpresa = tItemPedido.iFilialEmpresa
        .sUnidadeMed = tItemPedido.sUnidadeMed
        .dPrecoTotal = tItemPedido.dPrecoTotal
        .dPrecoUnitario = tItemPedido.dPrecoUnitario
        .dQuantCancelada = tItemPedido.dQuantCancelada
        .dQuantFaturada = tItemPedido.dQuantFaturada
        .dQuantidade = tItemPedido.dQuantidade
        .dQuantReservada = tItemPedido.dQuantReservada
        .dtDataEntrega = tItemPedido.dtDataEntrega
        .dValorAbatComissao = tItemPedido.dValorAbatComissao
        .dValorDesconto = tItemPedido.dValorDesconto
        .iClasseUM = tItemPedido.iClasseUM
        .iStatus = tItemPedido.iStatus
        .sDescricao = tItemPedido.sDescricao
        .sLote = tItemPedido.sLote
        .lNumIntDoc = tItemPedido.lNumIntDoc
        .dPercentMaisReceb = tItemPedido.dPercentMaisReceb
        .dPercentMenosReceb = tItemPedido.dPercentMenosReceb
        .iRecebForaFaixa = tItemPedido.iRecebForaFaixa
        .dQuantFaturadaAMais = tItemPedido.dQuantFaturadaAMais
        .iPrioridade = tItemPedido.iPrioridade
        .dComissao = tItemPedido.dComissao
        .iTabelaPreco = tItemPedido.iTabelaPreco
        .lCodPedido = tItemPedido.lCodPedido
        .sProduto = tItemPedido.sProduto
        .dQuantOP = tItemPedido.dQuantOP
        .dQuantSC = tItemPedido.dQuantSC
        .dPercDesc1 = tItemPedido.dPercDesc1
        .iTipoDesc1 = tItemPedido.iTipoDesc1
        .dPercDesc2 = tItemPedido.dPercDesc2
        .iTipoDesc2 = tItemPedido.iTipoDesc2
        .dPercDesc3 = tItemPedido.dPercDesc3
        .iTipoDesc3 = tItemPedido.iTipoDesc3
    
    End With

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ItemPedido_Le = SUCESSO

    Exit Function

Erro_ItemPedido_Le:

    ItemPedido_Le = Err

    Select Case Err

        Case 23968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23969, 23970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 23971 'Tratado na Rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152680)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function ItemPedido_QuantFaltaReservar(ByVal objItemPedido As ClassItemPedido, dQuantidadeFaltaReservar As Double) As Long
'Calcula a quantidade que falta reservar deste ítem no Pedido.
'Desconta o que já tem reservado do total a reservar.
'O ítem deve ser um Produto de Reserva
'Devolve em dQuantidadeFaltaReservar a quantidade na Unidade de Estoque.
    
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
    
On Error GoTo Erro_ItemPedido_QuantFaltaReservar
    
    'Quantidade Reservar na UM de Venda
    dQuantidadeFaltaReservar = objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada - objItemPedido.dQuantReservada
   
    'Lê o Produto
    objProduto.sCodigo = objItemPedido.sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 25070
    If lErro = 28030 Then Error 25071
    
    'Calcula o fator de conversão de UM de Venda para UM de Estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 25072
    
    'Quantidade a Reservar na UM de Estoque
    dQuantidadeFaltaReservar = dFator * dQuantidadeFaltaReservar
    
    ItemPedido_QuantFaltaReservar = SUCESSO
    
    Exit Function

Erro_ItemPedido_QuantFaltaReservar:

    ItemPedido_QuantFaltaReservar = Err

    Select Case Err

        Case 25070, 25072
        
        Case 25071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152681)

    End Select

    Exit Function

End Function

Function ItemPedido_QuantReservar(ByVal objItemPedido As ClassItemPedido, dQuantidadeReservar As Double) As Long
'Calcula a quantidade total que se deve reservar deste ítem no Pedido.
'Ignora o que já tem reservado.
'O ítem deve ser um Produto de Reserva
'Devolve em dQuantidadeReservar a quantidade na Unidade de Estoque.
    
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
    
On Error GoTo Erro_ItemPedido_QuantReservar
    
    'Quantidade Reservar na UM de Venda
    dQuantidadeReservar = objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada
   
    'Lê o Produto
    objProduto.sCodigo = objItemPedido.sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 25064
    If lErro = 28030 Then Error 25065
    
    'Calcula o fator de conversão de UM de Venda para UM de Estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 25063
    
    'Quantidade a Reservar na UM de Estoque
    dQuantidadeReservar = dFator * dQuantidadeReservar
    
    ItemPedido_QuantReservar = SUCESSO
    
    Exit Function

Erro_ItemPedido_QuantReservar:

    ItemPedido_QuantReservar = Err

    Select Case Err

        Case 25063, 25064
        
        Case 25065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152682)

    End Select

    Exit Function

End Function

Function PedidoDeVenda_Le_Itens(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê no BD a coleção de ítens do Pedido de Venda e devolve em colItensPedido dentro de objPedidoVenda

Dim lErro As Long
'Dim lComando As Long
'Dim tItemPedido As typeItemPedido
'Dim iItem As Integer
'Dim objItemPV As ClassItemPedido

On Error GoTo Erro_PedidoDeVenda_Le_Itens

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 26523
'
'    tItemPedido.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
'    tItemPedido.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
'    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
'    tItemPedido.sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
'    tItemPedido.sProdutoNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
'    tItemPedido.sUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA, 0)
'
'    lErro = Comando_Executar(lComando, "SELECT ItensPedidoDeVenda.NumIntDoc, ItensPedidoDeVenda.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensPedidoDeVenda.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensPedidoDeVenda.Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantReservada, QuantFaturada, Produtos.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, ItensPedidoDeVenda.QuantSC, ItensPedidoDeVenda.PercentMaisReceb, ItensPedidoDeVenda.PercentMenosReceb, ItensPedidoDeVenda.RecebForaFaixa, ItensPedidoDeVenda.QuantFaturadaAMais, ItensPedidoDeVenda.Prioridade  FROM ItensPedidoDeVenda, Produtos WHERE ItensPedidoDeVenda.FilialEmpresa = ? AND CodPedido = ? AND ItensPedidoDeVenda.Produto = Produtos.Codigo ORDER BY ItensPedidoDeVenda.NumIntDoc", tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
'        tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus, tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantReservada, tItemPedido.dQuantFaturada, tItemPedido.sProdutoDescricao, tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, tItemPedido.dQuantSC, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRecebForaFaixa, tItemPedido.dQuantFaturadaAMais, tItemPedido.iPrioridade, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 26524
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26525
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 26526
'
'    iItem = 0
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        iItem = iItem + 1
'        'Adiciona o ItemPedido na coleção de objPedidoVenda
'        With tItemPedido
'            Set objItemPV = objPedidoVenda.colItensPedido.Add(objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iItem, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc)
'            objItemPV.iPrioridade = .iPrioridade
'        End With
'
'        objItemPV.dPercentMaisReceb = tItemPedido.dPercentMaisReceb
'        objItemPV.dPercentMenosReceb = tItemPedido.dPercentMenosReceb
'        objItemPV.iRecebForaFaixa = tItemPedido.iRecebForaFaixa
'        objItemPV.dQuantFaturadaAMais = tItemPedido.dQuantFaturadaAMais
'
'        lErro = CF("ItensPVGrade_Le", objPedidoVenda.colItensPedido(objPedidoVenda.colItensPedido.Count))
'        If lErro <> SUCESSO Then gError 86358
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26527
'    Loop

    'XXXX
    'Tem que criar essa função e trocar nas funções já existentes para ter apenas um lugar para ler os itens do PV
    lErro = CF("ItensPedido_Le_Inclusive_Baixados", objPedidoVenda, STATUS_ABERTO)
    If lErro <> SUCESSO Then gError 86358
'
'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("PVItens_Le_Customizado", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 177040
'    '############################################

'    Call Comando_Fechar(lComando)

    PedidoDeVenda_Le_Itens = SUCESSO

    Exit Function

Erro_PedidoDeVenda_Le_Itens:

    PedidoDeVenda_Le_Itens = gErr

    Select Case gErr

'        Case 26523
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 26524, 26525, 26527
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA2", gErr, objPedidoVenda.lCodigo)
'
'        Case 26526
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_SEM_ITENS", gErr, objPedidoVenda.lCodigo)
'
        Case 86358 ', 177040

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152683)

    End Select

'    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoVenda_Lock(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long) As Long
'Lê e faz o Lock no Pedido de Venda

Dim lErro As Long
Dim tPedVenda As typePedidoVenda
Dim lCodigo As Long

On Error GoTo Erro_PedidoVenda_Lock

    'Lê o Pedido com o Código e a filialempresa passados
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, ValorTotal, FilialEmpresaFaturamento FROM PedidosDeVenda WHERE FilialEmpresa = ? and Codigo = ?", 0, tPedVenda.lNumIntDoc, tPedVenda.dValorTotal, tPedVenda.iFilialEmpresaFaturamento, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 42986

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42987
    If lErro = AD_SQL_SEM_DADOS Then Error 42988 'Não encontrou

    'Faz o "Lock" no Pedido de Venda
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 42989

    objPedidoVenda.lNumIntDoc = tPedVenda.lNumIntDoc
    objPedidoVenda.dValorTotal = tPedVenda.dValorTotal
    objPedidoVenda.iFilialEmpresaFaturamento = tPedVenda.iFilialEmpresaFaturamento

    PedidoVenda_Lock = SUCESSO

    Exit Function

Erro_PedidoVenda_Lock:

    PedidoVenda_Lock = Err

    Select Case Err

        Case 42986, 42987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedidoVenda.lCodigo)

        Case 42988

        Case 42989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152684)

    End Select

    Exit Function

End Function

Function ItensPV_Le_Lock(ByVal objPedVenda As ClassPedidoDeVenda) As Long
'Faz lock exclusive nos itens de PedVenda e lê os NumIntDoc desses itens

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItemPV As typeItemPedido
Dim tItemRomaneio As typeItemRomaneioGrade
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ItensPV_Le_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 42990

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 42990

    tItemPV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPV.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemPV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    'Lê os itens do pedido passado
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ?", 0, tItemPV.lNumIntDoc, tItemPV.sProduto, tItemPV.dQuantidade, tItemPV.dPrecoUnitario, tItemPV.dPrecoTotal, tItemPV.iClasseUM, tItemPV.sUnidadeMed, tItemPV.dValorDesconto, tItemPV.dtDataEntrega, tItemPV.sDescricao, tItemPV.sLote, tItemPV.iStatus, tItemPV.dValorAbatComissao, tItemPV.dQuantCancelada, tItemPV.dQuantReservada, tItemPV.dQuantFaturada, tItemPV.dQuantOP, tItemPV.dQuantSC, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 42991

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 42992
    If lErro = AD_SQL_SEM_DADOS Then gError 42993 'Não encontrou

    Do While lErro = AD_SQL_SUCESSO
        'faz o lock do item
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 42994

        With tItemPV
            Set tItemPV.colReservaItem = New ColReserva

            .iFilialEmpresa = objPedVenda.iFilialEmpresa
            .lCodPedido = objPedVenda.lCodigo

            'Adiciona o item a colecao de itens do pedido de Venda
            objPedVenda.colItensPedido.Add .iFilialEmpresa, .lCodPedido, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, "", .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, "", .sUnidadeMed, .sLote, "", .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc
        End With
        
        With tItemRomaneio
            .sProduto = String(STRING_PRODUTO, 0)
    
            lErro = Comando_Executar(lComando1, "SELECT NumIntDoc,NumIntItemPV,Produto,Quantidade,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC FROM ItensPedidoDeVendaGrade WHERE NumIntItemPV = ?", _
            .lNumIntDoc, .lNumIntItemPV, .sProduto, .dQuantidade, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, tItemPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 86343
        
        End With
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86344
        
        Do While lErro <> AD_SQL_SEM_DADOS

            Set objItemRomaneio = New ClassItemRomaneioGrade
            
            objItemRomaneio.dQuantAFaturar = tItemRomaneio.dQuantAFaturar
            objItemRomaneio.dQuantCancelada = tItemRomaneio.dQuantCancelada
            objItemRomaneio.dQuantFaturada = tItemRomaneio.dQuantFaturada
            objItemRomaneio.dQuantidade = tItemRomaneio.dQuantidade
            objItemRomaneio.dQuantOP = tItemRomaneio.dQuantOP
            objItemRomaneio.dQuantReservada = tItemRomaneio.dQuantReservada
            objItemRomaneio.dQuantSC = tItemRomaneio.dQuantSC
            objItemRomaneio.lNumIntDoc = tItemRomaneio.lNumIntDoc
            objItemRomaneio.sProduto = tItemRomaneio.sProduto
            
            objPedVenda.colItensPedido(objPedVenda.colItensPedido.Count).colItensRomaneioGrade.Add objItemRomaneio
            
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86345
            
        Loop
        
        'Busca o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 42995

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensPV_Le_Lock = SUCESSO

    Exit Function

Erro_ItensPV_Le_Lock:

    ItensPV_Le_Lock = gErr

    Select Case gErr

        Case 42990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42991, 42992, 42995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA2", gErr, objPedVenda.lCodigo)

        Case 42993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_SEM_ITENS", gErr, objPedVenda.lCodigo)

        Case 42994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDODEVENDA", gErr, objPedVenda.lCodigo)

        Case 86343 To 86345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDVENDAGRADE", gErr, objPedVenda.colItensPedido(objPedVenda.colItensPedido.Count).sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152685)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BloqueioEstoquePV_Le(ByVal objBloqueioPV As ClassBloqueioPV) As Long
'Lê o Bloqueio de Estoque correspondente ao Pedido que não esteja liberado

Dim lErro As Long
Dim lComando As Long
Dim tBloqueioPV As typeBloqueioPV

On Error GoTo Erro_BloqueioEstoquePV_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46029

    tBloqueioPV.sCodUsuario = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sCodUsuarioLib = String(STRING_BLOQUEIOSPV_COD_USUARIO, 0)
    tBloqueioPV.sResponsavel = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)
    tBloqueioPV.sResponsavelLib = String(STRING_BLOQUEIOSPV_RESPONSAVEL, 0)

    'Lê o bloqueio
    lErro = Comando_Executar(lComando, "SELECT Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,ResponsavelLib FROM BloqueiosPV WHERE PedidoDeVenda = ? AND FilialEmpresa = ? AND (TipoDeBloqueio =? OR TipoDeBloqueio = ? OR TipoDeBloqueio = ?) AND DataLib =?", tBloqueioPV.iSequencial, tBloqueioPV.iTipoDeBloqueio, tBloqueioPV.sCodUsuario, tBloqueioPV.sResponsavel, tBloqueioPV.dtData, tBloqueioPV.sCodUsuarioLib, tBloqueioPV.sResponsavelLib, objBloqueioPV.lPedidoDeVendas, giFilialEmpresa, BLOQUEIO_TOTAL, BLOQUEIO_PARCIAL, BLOQUEIO_NAO_RESERVA, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then Error 46030

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46031
    If lErro <> AD_SQL_SUCESSO Then Error 46032 'Não encontrou

    'Armazena os dados do bloqueio
    objBloqueioPV.dtData = tBloqueioPV.dtData
    objBloqueioPV.dtDataLib = DATA_NULA
    objBloqueioPV.iFilialEmpresa = giFilialEmpresa
    objBloqueioPV.iSequencial = tBloqueioPV.iSequencial
    objBloqueioPV.iTipoDeBloqueio = tBloqueioPV.iTipoDeBloqueio
    objBloqueioPV.sCodUsuario = tBloqueioPV.sCodUsuario
    objBloqueioPV.sCodUsuarioLib = tBloqueioPV.sCodUsuarioLib
    objBloqueioPV.sResponsavel = tBloqueioPV.sResponsavel
    objBloqueioPV.sResponsavelLib = tBloqueioPV.sResponsavelLib

    'Fechar comando
    Call Comando_Fechar(lComando)

    BloqueioEstoquePV_Le = SUCESSO

    Exit Function

Erro_BloqueioEstoquePV_Le:

    BloqueioEstoquePV_Le = Err

    Select Case Err

        Case 46029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46030, 46031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", Err, objBloqueioPV.lPedidoDeVendas)

        Case 46032 'Não cadastrado. A ser tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152686)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemOP_Le_ItemPV(ByVal objItemOP As ClassItemOP, ByVal objItemPV As ClassItemPedido) As Long
'Lê o item de Ordem de Produção vinculado ao Item de Pedido de Venda passado

Dim lErro As Long
Dim lComando As Long
Dim tItemOP As typeItemOP

On Error GoTo Erro_ItemOP_Le_ItemPV

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46072

    tItemOP.sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
    tItemOP.sCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
    tItemOP.sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
    tItemOP.sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)

    'Lê Item OP a partir do número interno do item de Pedido Venda (NumIntOrigem)
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc,FilialEmpresa,Codigo,Item,Produto,FilialPedido,CodPedido,Almoxarifado,Ccl,SiglaUM,Quantidade,DataInicioProd,DataFimProd,Prioridade,QuantidadeProd,Situacao,Destinacao FROM ItensOPTodos WHERE NumIntOrigem = ?", tItemOP.lNumIntDoc, tItemOP.iFilialEmpresa, tItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, objItemPV.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 46076

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 46073
    If lErro = AD_SQL_SEM_DADOS Then Error 46074 'Não encontrou

    'Armazena os dados lidos
    With tItemOP
        objItemOP.dQuantidade = .dQuantidade
        objItemOP.dQuantidadeProd = .dQuantidadeProd
        objItemOP.dtDataFimProd = .dtDataFimProd
        objItemOP.dtDataInicioProd = .dtDataInicioProd
        objItemOP.iAlmoxarifado = .iAlmoxarifado
        objItemOP.iDestinacao = .iDestinacao
        objItemOP.iFilialEmpresa = .iFilialEmpresa
        objItemOP.iFilialPedido = .iFilialPedido
        objItemOP.iItem = .iItem
        objItemOP.iPrioridade = .iPrioridade
        objItemOP.iSituacao = .iSituacao
        objItemOP.lCodPedido = .lCodPedido
        objItemOP.lNumIntDoc = .lNumIntDoc
        objItemOP.lNumIntOrigem = objItemPV.lNumIntDoc
        objItemOP.sCcl = .sCcl
        objItemOP.sCodigo = .sCodigo
        objItemOP.sProduto = .sProduto
        objItemOP.sSiglaUM = .sSiglaUM

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemOP_Le_ItemPV = SUCESSO

    Exit Function

Erro_ItemOP_Le_ItemPV:

    ItemOP_Le_ItemPV = Err

    Select Case Err

        Case 46072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46073, 46076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", Err)

        Case 46074
            'tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152687)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoVendaBaixado_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda Baixado com o Código e a Filial passados

Dim lErro As Long
Dim tPedidoVenda As typePedidoVenda
Dim lComando As Long

On Error GoTo Erro_PedidoVendaBaixado_Le

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46132

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

    'Busca o Pedido de Venda Baixado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,VolumeNumero,FreteRespons, ValorProdutos, CodTranspRedesp, DetPagFrete, DataEntrega, Moeda, TaxaMoeda, VolumeTotal,PedidoRepresentante, Andamento, OBS, " & _
    "ValorDescontoTit,ValorDescontoItens,ValorItens, CodigoBase, Parc, Emitente, UsuarioUltAlteracao, DataInclusao, DataAlteracao, HoraInclusao, HoraAlteracao FROM PedidosDeVendaBaixados WHERE FilialEmpresa = ? AND Codigo = ?", tPedidoVenda.iFilialEmpresaFaturamento, tPedidoVenda.lCliente, tPedidoVenda.iFilial, tPedidoVenda.iFilialEntrega, tPedidoVenda.iCodTransportadora, tPedidoVenda.iCondicaoPagto, tPedidoVenda.dPercAcrescFinanceiro, tPedidoVenda.dtDataEmissao, tPedidoVenda.dtDataReferencia, tPedidoVenda.sMensagemNota, _
        tPedidoVenda.sNaturezaOp, tPedidoVenda.dValorTotal, tPedidoVenda.dValorFrete, tPedidoVenda.dValorDesconto, tPedidoVenda.dValorSeguro, tPedidoVenda.dValorOutrasDespesas, tPedidoVenda.sPedidoCliente, tPedidoVenda.iCanalVenda, tPedidoVenda.iTabelaPreco, tPedidoVenda.iProxSeqBloqueio, tPedidoVenda.iFaturaIntegral, tPedidoVenda.iComissaoAutomatica, tPedidoVenda.iCobrancaAutomatica, tPedidoVenda.lNumIntDoc, tPedidoVenda.sPlaca, tPedidoVenda.sPlacaUF, tPedidoVenda.lVolumeQuant, tPedidoVenda.lVolumeEspecie, tPedidoVenda.lVolumeMarca, tPedidoVenda.dPesoLiq, tPedidoVenda.dPesoBruto, tPedidoVenda.sVolumeNumero, tPedidoVenda.iFreteRespons, tPedidoVenda.dValorProdutos, tPedidoVenda.iCodTranspRedesp, tPedidoVenda.iDetPagFrete, tPedidoVenda.dtDataEntrega, tPedidoVenda.iMoeda, tPedidoVenda.dTaxaMoeda, tPedidoVenda.dVolumeTotal, tPedidoVenda.sPedidoRepresentante, tPedidoVenda.iAndamento, tPedidoVenda.sOBS, _
        tPedidoVenda.dValorDescontoTit, tPedidoVenda.dValorDescontoItens, tPedidoVenda.dValorItens, tPedidoVenda.lCodigoBase, tPedidoVenda.iParc, tPedidoVenda.sEmitente, tPedidoVenda.sUsuarioUltAlteracao, tPedidoVenda.dtDataInclusao, tPedidoVenda.dtDataAlteracao, tPedidoVenda.dHoraInclusao, tPedidoVenda.dHoraAlteracao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 46133

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46134
    If lErro = AD_SQL_SEM_DADOS Then Error 46135 'Não encontrou

    With tPedidoVenda
        'Armazena os dados lidos
        objPedidoVenda.dPercAcrescFinanceiro = .dPercAcrescFinanceiro
        objPedidoVenda.dtDataEmissao = .dtDataEmissao
        objPedidoVenda.dtDataReferencia = .dtDataReferencia
        objPedidoVenda.dValorDesconto = .dValorDesconto
        objPedidoVenda.dValorFrete = .dValorFrete
        objPedidoVenda.dValorOutrasDespesas = .dValorOutrasDespesas
        objPedidoVenda.dValorSeguro = .dValorSeguro
        objPedidoVenda.dValorTotal = .dValorTotal
        objPedidoVenda.iCanalVenda = .iCanalVenda
        objPedidoVenda.iCobrancaAutomatica = .iCobrancaAutomatica
        objPedidoVenda.iCodTransportadora = .iCodTransportadora
        objPedidoVenda.iCodTranspRedesp = .iCodTranspRedesp
        objPedidoVenda.iDetPagFrete = .iDetPagFrete
        objPedidoVenda.dtDataEntrega = .dtDataEntrega
        objPedidoVenda.iComissaoAutomatica = .iComissaoAutomatica
        objPedidoVenda.iCondicaoPagto = .iCondicaoPagto
        objPedidoVenda.iFaturaIntegral = .iFaturaIntegral
        objPedidoVenda.iFilial = .iFilial
        objPedidoVenda.iFilialEmpresaFaturamento = .iFilialEmpresaFaturamento
        objPedidoVenda.iFilialEntrega = .iFilialEntrega
        objPedidoVenda.iProxSeqBloqueio = .iProxSeqBloqueio
        objPedidoVenda.iTabelaPreco = .iTabelaPreco
        objPedidoVenda.iTipoPedido = .iTipoPedido
        objPedidoVenda.lCliente = .lCliente
        objPedidoVenda.lNumIntDoc = .lNumIntDoc
        objPedidoVenda.sMensagemNota = .sMensagemNota
        objPedidoVenda.sNaturezaOp = .sNaturezaOp
        objPedidoVenda.sPedidoCliente = .sPedidoCliente
        objPedidoVenda.sPlaca = .sPlaca
        objPedidoVenda.sPlacaUF = .sPlacaUF
        objPedidoVenda.lVolumeQuant = .lVolumeQuant
        objPedidoVenda.lVolumeEspecie = .lVolumeEspecie
        objPedidoVenda.lVolumeMarca = .lVolumeMarca
        objPedidoVenda.dPesoLiq = .dPesoLiq
        objPedidoVenda.dPesoBruto = .dPesoBruto
        objPedidoVenda.iFreteRespons = .iFreteRespons
        objPedidoVenda.sVolumeNumero = .sVolumeNumero
        objPedidoVenda.dValorProdutos = .dValorProdutos
        objPedidoVenda.sPedidoRepresentante = .sPedidoRepresentante
        objPedidoVenda.iMoeda = .iMoeda
        objPedidoVenda.dTaxaMoeda = .dTaxaMoeda
        objPedidoVenda.dVolumeTotal = .dVolumeTotal
        objPedidoVenda.sOBS = .sOBS
        objPedidoVenda.iAndamento = .iAndamento
        objPedidoVenda.dValorDescontoTit = .dValorDescontoTit
        objPedidoVenda.dValorItens = .dValorItens
        objPedidoVenda.dValorDescontoItens = .dValorDescontoItens
        objPedidoVenda.lCodigoBase = .lCodigoBase
        objPedidoVenda.iParc = .iParc
        objPedidoVenda.sEmitente = .sEmitente
        objPedidoVenda.sUsuarioUltAlteracao = .sUsuarioUltAlteracao
        objPedidoVenda.dtDataInclusao = .dtDataInclusao
        objPedidoVenda.dtDataAlteracao = .dtDataAlteracao
        objPedidoVenda.dHoraInclusao = .dHoraInclusao
        objPedidoVenda.dHoraAlteracao = .dHoraAlteracao

    End With

    lErro = CF("PV_Le_Customizado", objPedidoVenda)
    If lErro <> SUCESSO Then Error 46133
        
    'Fecha o comando
    Call Comando_Fechar(lComando)

    PedidoVendaBaixado_Le = SUCESSO

    Exit Function

Erro_PedidoVendaBaixado_Le:

    PedidoVendaBaixado_Le = Err

    Select Case Err

        Case 46132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46133, 43134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDABAIXADOS", Err, objPedidoVenda.lCodigo)

        Case 46135
            'Tratado pela rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152688)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ReservasPedido_Le(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as Reservas dos itens do Pedido passado

Dim lComando As Long
Dim lErro As Long
Dim tReserva As typeReserva
Dim objReserva As ClassReserva
Dim objItemPV As ClassItemPedido
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ReservasPedido_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46049

    'Para cada item do Pedido
    For Each objItemPV In objPedidoVenda.colItensPedido

        tReserva.sCodUsuario = String(STRING_RESERVA_USUARIO, 0)
        tReserva.sProduto = String(STRING_PRODUTO, 0)
        tReserva.sResponsavel = String(STRING_RESERVA_RESPONSAVEL, 0)

        'Seleciona reservas
        lErro = Comando_Executar(lComando, "SELECT Codigo,Produto,Almoxarifado,NumIntOrigem,Quantidade,DataReserva,DataValidade,CodUsuario,Responsavel,NumIntDoc FROM Reserva WHERE Filialempresa = ? AND TipoDoc = ? AND DocOrigem = ? AND NumIntOrigem = ?", tReserva.lCodigo, tReserva.sProduto, tReserva.iAlmoxarifado, tReserva.lNumIntOrigem, tReserva.dQuantidade, tReserva.dtDataReserva, tReserva.dtDataValidade, tReserva.sCodUsuario, tReserva.sResponsavel, tReserva.lNumIntDoc, objPedidoVenda.iFilialEmpresa, PEDIDO_DE_VENDA, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 46056

        'Tenta ler a primeira
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46057

        Do While lErro = AD_SQL_SUCESSO

            'Adiciona a reserva na coleção de reserva do item
            With tReserva
                objItemPV.ColReserva.Add objPedidoVenda.iFilialEmpresa, 0, .sProduto, .iAlmoxarifado, 0, 0, 0, .dQuantidade, DATA_NULA, .dtDataValidade, "", .sResponsavel, .lNumIntDoc, ""
            End With

            'Busca próxima reserva
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 46058

        Loop
        
        If objItemPV.colItensRomaneioGrade.Count > 0 Then
            
            For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
                lErro = ReservasItemGrade_Le(objItemRomaneio)
                If lErro <> SUCESSO Then Error 11111
            Next
        End If

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ReservasPedido_Le = SUCESSO

    Exit Function

Erro_ReservasPedido_Le:

    ReservasPedido_Le = Err

    Select Case Err

        Case 46049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46056, 46057, 46058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152689)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function ItensPV_ItensOP_Vinculados(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'verifica se existe alguma Item de Ordem de Produção vincula do a algum Item do Pedido de Venda passado

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_ItensPV_ItensOP_Vinculados

    'Para casa item de Pedido de Venda
    For Each objItemPV In objPedidoVenda.colItensPedido
        'Lê o item de Ordem de Produção vinculado
        lErro = CF("ItemOP_Le_ItemPV", objItemOP, objItemPV)
        If lErro <> SUCESSO And lErro <> 46074 Then Error 46071
        If lErro = SUCESSO Then Error 46077 'Se encontrou --> Erro

    Next

    ItensPV_ItensOP_Vinculados = SUCESSO

    Exit Function

Erro_ItensPV_ItensOP_Vinculados:

    ItensPV_ItensOP_Vinculados = Err

    Select Case Err

        Case 46071

        Case 46077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_OP_VINCULADO_ITEM_PV", Err, objPedidoVenda.lCodigo, objItemPV.lNumIntDoc, objItemOP.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152690)

    End Select

    Exit Function

End Function

Function PVObjetosVinculados_Lock(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Faz "lock" nos objetos associados ao Pedido de Venda

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim objProduto As New ClassProduto
Dim objItemPV As ClassItemPedido
Dim objReserva As ClassReserva
Dim iIndice As Integer
Dim objBloqueioPV As Object
Dim objComissao As ClassComissaoPedVendas

On Error GoTo Erro_PVObjetosVinculados_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 42926
    Next

    'Loca a FilialCliente
    lErro = CF("FilialCliente_Lock", objPedidoVenda.lCliente, objPedidoVenda.iFilial)
    If lErro <> SUCESSO And lErro <> 43050 Then Error 42925
    If lErro <> SUCESSO Then Error 42927

    'Loca a NaturezaOp
    lErro = CF("NaturezaOperacao_Lock", objPedidoVenda.sNaturezaOp, alComando(0))
    If lErro <> SUCESSO And lErro <> 40551 Then Error 42928
    If lErro <> SUCESSO Then Error 42929

    If objPedidoVenda.iFilialEntrega <> 0 Then
        'Locka a FilialEntrega
        lErro = CF("FilialCliente_Lock", objPedidoVenda.lCliente, objPedidoVenda.iFilialEntrega)
        If lErro <> SUCESSO And lErro <> 43050 Then Error 42930
        If lErro <> SUCESSO Then Error 42931
    End If

    If objPedidoVenda.iCanalVenda <> 0 Then
        'Locka o Canal de Venda
        lErro = CF("CanalVenda_Lock", alComando(2), objPedidoVenda.iCanalVenda)
        If lErro <> SUCESSO And lErro <> 43026 Then Error 42932
        If lErro <> SUCESSO Then Error 43170   'Se não achou o Canal de Venda --> erro
    End If

    If objPedidoVenda.iCodTransportadora <> 0 Then
        'Locka a Transportadora
        lErro = CF("Transportadora_Lock", alComando(3), objPedidoVenda.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 42933
        If lErro <> SUCESSO Then Error 42934 'Se não achou a Transportadora --> erro
    End If

    If objPedidoVenda.iCodTranspRedesp <> 0 Then
        'Locka a Transportadora de redespacho
        lErro = CF("Transportadora_Lock", alComando(3), objPedidoVenda.iCodTranspRedesp)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 42933
        If lErro <> SUCESSO Then Error 42934 'Se não achou a Transportadora --> erro
    End If

    If objPedidoVenda.iTabelaPreco <> 0 Then
        'Locka a TabelaPreco
        lErro = CF("TabelaDePreco_Lock", objPedidoVenda.iTabelaPreco)
        If lErro <> SUCESSO And lErro <> 43128 Then Error 42935
        If lErro <> SUCESSO Then Error 43171 'Se não achou a Tabela de Preço --> erro
    End If

    If objPedidoVenda.iCondicaoPagto <> 0 Then
        'Locka a Condição de Pagamento
        lErro = CF("CondicaoPagto_Lock", alComando(0), objPedidoVenda.iCondicaoPagto)
        If lErro <> SUCESSO Then Error 42936
    End If

    For Each objComissao In objPedidoVenda.colComissoes
        lErro = CF("Vendedor_Lock", alComando(4), objComissao.iCodVendedor)
        If lErro <> SUCESSO Then Error 51409
    Next
    
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        If UCase(TypeName(objBloqueioPV)) = "CLASSBLOQUEIOPV" Then
            lErro = TipoBloqueioPV_Lock(objBloqueioPV.iTipoDeBloqueio)
            If lErro <> SUCESSO Then Error 51410
        End If
    Next

    'Para cada item de Pedido Venda
    For Each objItemPV In objPedidoVenda.colItensPedido

        objProduto.sCodigo = objItemPV.sProduto
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(5), objProduto)
        If lErro <> SUCESSO Then Error 42937

        'Armazena os dados do produto para que possam ser usados depois
        objItemPV.iControleEstoque = objProduto.iControleEstoque
        objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque 'usa nas reservas
        objItemPV.iClasseUM = objProduto.iClasseUM
        objItemPV.iApropriacaoProd = objProduto.iApropriacaoCusto

        'Para casa reserva do item
        For Each objReserva In objItemPV.ColReserva
            'Locka o Almoxarifado
            lErro = CF("Almoxarifado_Lock", objReserva.iAlmoxarifado, alComando(1))
            If lErro <> SUCESSO Then Error 42938
        Next
        
    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

     PVObjetosVinculados_Lock = SUCESSO

    Exit Function

Erro_PVObjetosVinculados_Lock:

    PVObjetosVinculados_Lock = Err

    Select Case Err

        Case 42925, 42928, 42930, 42932, 42933, 42935, 42936, 42937, 42938, 51409, 51410

        Case 42926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 42927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objPedidoVenda.iFilial, objPedidoVenda.lCliente)

        Case 42929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objPedidoVenda.sNaturezaOp)

        Case 42931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", objPedidoVenda.lCliente, objPedidoVenda.iFilialEntrega)

        Case 42934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objPedidoVenda.iCodTransportadora)

        Case 43170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_CADASTRADO", Err, objPedidoVenda.iCanalVenda)

        Case 43171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objPedidoVenda.iTabelaPreco)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152691)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueiosPV_Credito_Aprovado_Testa(ByVal objPedidoVenda As ClassPedidoDeVenda, iCreditoAprovado As Integer) As Long
'Verifica se tem bloqueio de crédito liberado

'Função alterada por Leo em 26/02/02. Agora só chama uma função mais genérica passando o tipo de bloqueio a ser verificado.
Dim lErro As Long

On Error GoTo Erro_BloqueiosPV_Credito_Aprovado_Testa

    'verifica se o pedido tem bloqueio por credito aprovado (função genérica, basta passar o tipo de Bloqueio)
    lErro = CF("BloqueiosPV_Aprovado_Testa", objPedidoVenda, iCreditoAprovado, BLOQUEIO_CREDITO)
    If lErro <> SUCESSO Then gError 94387

    BloqueiosPV_Credito_Aprovado_Testa = SUCESSO

    Exit Function

Erro_BloqueiosPV_Credito_Aprovado_Testa:

    BloqueiosPV_Credito_Aprovado_Testa = gErr

    Select Case gErr

        Case 94387
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152692)

    End Select

    Exit Function

End Function

Function SldDiaFat_Le_Lock(ByVal lComando As Long, ByVal objSldDiaFat As ClassSldDiaFat) As Long
'Lê todos atributos da tabela SldDiaFat e faz lock exclusive
'Filtro: filial/produto/data passados em objSldDiaFat

Dim lErro As Long
Dim tSldDiaFat As typeSldDiaFat

On Error GoTo Erro_SldDiaFat_Le_Lock

    'Le a tabela de SldDiaFat para a filial/produto/data em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT ValorFaturado, TotalDescontos, QuantPedida, QuantFaturada, ValorDevolvido, QuantDevolvida, QuantPedidaSRV FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data = ?", 0, tSldDiaFat.dValorFaturado, tSldDiaFat.dTotalDescontos, tSldDiaFat.dQuantPedida, tSldDiaFat.dQuantFaturada, tSldDiaFat.dValorDevolvido, tSldDiaFat.dQuantDevolvida, tSldDiaFat.dQuantPedidaSRV, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 52727

    'leitura do registro referente a filial/produto/data em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52728
    
    If lErro = AD_SQL_SEM_DADOS Then Error 52729
    
    'Lock da SldDiaFat
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 52730
    
    objSldDiaFat.dValorFaturado = tSldDiaFat.dValorFaturado
    objSldDiaFat.dTotalDescontos = tSldDiaFat.dTotalDescontos
    objSldDiaFat.dQuantPedida = tSldDiaFat.dQuantPedida
    objSldDiaFat.dQuantFaturada = tSldDiaFat.dQuantFaturada
    objSldDiaFat.dValorDevolvido = tSldDiaFat.dValorDevolvido
    objSldDiaFat.dQuantDevolvida = tSldDiaFat.dQuantDevolvida
    objSldDiaFat.dQuantPedidaSRV = tSldDiaFat.dQuantPedidaSRV
    
    SldDiaFat_Le_Lock = SUCESSO
    
    Exit Function
    
Erro_SldDiaFat_Le_Lock:
    
    SldDiaFat_Le_Lock = Err
    
    Select Case Err
        
        Case 52727, 52728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFAT", Err, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
         
        Case 52729 'Tratado na rotina chamadora
        
        Case 52730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDDIAFAT", Err, objSldDiaFat.iFilialEmpresa, objSldDiaFat.sProduto, objSldDiaFat.dtData)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152693)

    End Select

    Exit Function

End Function

Function SldMesFat_Le_Lock(ByVal lComando As Long, ByVal objSldMesFat As ClassSldMesFat) As Long
'Lê todos atributos da tabela SldMesFat e faz lock exclusive
'Filtro: filial/produto/ano passados em objSldMesFat

Dim lErro As Long
Dim tSldMesFat As typeSldMesFat
Dim iIndice As Integer
Dim sIndice As String
Dim sComando_SqlPrimario As String
Dim sComando_SQL As String

On Error GoTo Erro_SldMesFat_Le_Lock
        
    'Le a tabela de SldMesFat para a filial/produto/ano em questão
    With tSldMesFat
        lErro = Comando_ExecutarPos(lComando, "SELECT ValorFaturado1,ValorDevolvido1, TotalDescontos1, QuantPedida1, QuantFaturada1, QuantDevolvida1, QuantPedidaSRV1, ValorFaturado2,ValorDevolvido2, TotalDescontos2, QuantPedida2, QuantFaturada2, QuantDevolvida2, QuantPedidaSRV2, ValorFaturado3,ValorDevolvido3, TotalDescontos3, QuantPedida3, QuantFaturada3, QuantDevolvida3, QuantPedidaSRV3, ValorFaturado4,ValorDevolvido4, TotalDescontos4, QuantPedida4, QuantFaturada4, QuantDevolvida4, QuantPedidaSRV4, ValorFaturado5,ValorDevolvido5, TotalDescontos5, QuantPedida5, QuantFaturada5, QuantDevolvida5, QuantPedidaSRV5, ValorFaturado6,ValorDevolvido6, TotalDescontos6, QuantPedida6, QuantFaturada6, QuantDevolvida6, QuantPedidaSRV6, ValorFaturado7,ValorDevolvido7, TotalDescontos7, QuantPedida7, QuantFaturada7, QuantDevolvida7, QuantPedidaSRV7, " _
        & "ValorFaturado8,ValorDevolvido8, TotalDescontos8, QuantPedida8, QuantFaturada8, QuantDevolvida8, QuantPedidaSRV8, ValorFaturado9,ValorDevolvido9, TotalDescontos9, QuantPedida9, QuantFaturada9, QuantDevolvida9, QuantPedidaSRV9,   ValorFaturado10,ValorDevolvido10, TotalDescontos10, QuantPedida10, QuantFaturada10, QuantDevolvida10, QuantPedidaSRV10,   ValorFaturado11,ValorDevolvido11, TotalDescontos11, QuantPedida11, QuantFaturada11, QuantDevolvida11, QuantPedidaSRV11, ValorFaturado12,ValorDevolvido12, TotalDescontos12, QuantPedida12, QuantFaturada12, QuantDevolvida12, QuantPedidaSRV12  FROM SldMesFat WHERE FilialEmpresa = ? AND Produto = ? AND Ano = ?", 0, _
        .adValorFaturado(1), .adValorDevolvido(1), .adTotalDescontos(1), .adQuantPedida(1), .adQuantFaturada(1), .adQuantDevolvida(1), .adQuantPedidaSRV(1), .adValorFaturado(2), .adValorDevolvido(2), .adTotalDescontos(2), .adQuantPedida(2), .adQuantFaturada(2), .adQuantDevolvida(2), .adQuantPedidaSRV(2), .adValorFaturado(3), .adValorDevolvido(3), .adTotalDescontos(3), .adQuantPedida(3), .adQuantFaturada(3), .adQuantDevolvida(3), .adQuantPedidaSRV(3), .adValorFaturado(4), .adValorDevolvido(4), .adTotalDescontos(4), .adQuantPedida(4), .adQuantFaturada(4), .adQuantDevolvida(4), .adQuantPedidaSRV(4), .adValorFaturado(5), .adValorDevolvido(5), .adTotalDescontos(5), .adQuantPedida(5), .adQuantFaturada(5), .adQuantDevolvida(5), .adQuantPedidaSRV(5), .adValorFaturado(6), .adValorDevolvido(6), .adTotalDescontos(6), .adQuantPedida(6), .adQuantFaturada(6), .adQuantDevolvida(6), .adQuantPedidaSRV(6), _
        .adValorFaturado(7), .adValorDevolvido(7), .adTotalDescontos(7), .adQuantPedida(7), .adQuantFaturada(7), .adQuantDevolvida(7), .adQuantPedidaSRV(7), _
        .adValorFaturado(8), .adValorDevolvido(8), .adTotalDescontos(8), .adQuantPedida(8), .adQuantFaturada(8), .adQuantDevolvida(8), .adQuantPedidaSRV(8), .adValorFaturado(9), .adValorDevolvido(9), .adTotalDescontos(9), .adQuantPedida(9), .adQuantFaturada(9), .adQuantDevolvida(9), .adQuantPedidaSRV(9), .adValorFaturado(10), .adValorDevolvido(10), .adTotalDescontos(10), .adQuantPedida(10), .adQuantFaturada(10), .adQuantDevolvida(10), .adQuantPedidaSRV(10), .adValorFaturado(11), .adValorDevolvido(11), .adTotalDescontos(11), .adQuantPedida(11), .adQuantFaturada(11), .adQuantDevolvida(11), .adQuantPedidaSRV(11), .adValorFaturado(12), .adValorDevolvido(12), .adTotalDescontos(12), .adQuantPedida(12), .adQuantFaturada(12), .adQuantDevolvida(12), .adQuantPedidaSRV(12), objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto, objSldMesFat.iAno)
        If lErro <> AD_SQL_SUCESSO Then Error 52731
    
    End With
    
    'leitura do registro referente a filial/produto/ano em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52732
    
    If lErro = AD_SQL_SEM_DADOS Then Error 52733
    
    'Lock da SldMesFat
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 52734
    
    'Preenche o objMesFat
    For iIndice = 1 To 12
    
        objSldMesFat.dValorFaturado(iIndice) = tSldMesFat.adValorFaturado(iIndice)
        objSldMesFat.dValorDevolvido(iIndice) = tSldMesFat.adValorDevolvido(iIndice)
        objSldMesFat.dTotalDescontos(iIndice) = tSldMesFat.adTotalDescontos(iIndice)
        objSldMesFat.dQuantPedida(iIndice) = tSldMesFat.adQuantPedida(iIndice)
        objSldMesFat.dQuantFaturada(iIndice) = tSldMesFat.adQuantFaturada(iIndice)
        objSldMesFat.dQuantDevolvida(iIndice) = tSldMesFat.adQuantDevolvida(iIndice)
        objSldMesFat.dQuantPedidaSRV(iIndice) = tSldMesFat.adQuantPedidaSRV(iIndice)
    
    Next
                
    SldMesFat_Le_Lock = SUCESSO
    
    Exit Function
    
Erro_SldMesFat_Le_Lock:

    SldMesFat_Le_Lock = Err
    
    Select Case Err
        
        Case 52731, 52732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", Err, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case 52733 'Tratado na rotina chamadora
        
        Case 52734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESFAT", Err, objSldMesFat.iAno, objSldMesFat.iFilialEmpresa, objSldMesFat.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152694)

    End Select

    Exit Function

End Function

'Inicio da Leitura de Cliente e FilialCliente com relacao a Estatística

Function Cliente_Le_Estendida(ByVal objCliente As ClassCliente, ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le todos os Dados do cliente inclusive os dados de Estatistica

Dim lErro As Long

On Error GoTo Erro_Cliente_Le_Estendida
    
    'Faz a Leitura na Tabela de cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 52544
    
    'Se não encontrou  - - -> Erro
    If lErro = 12293 Then Error 52545
    
    objClienteEstatistica.lCodCliente = objCliente.lCodigo
    objClienteEstatistica.iFilialEmpresa = giFilialEmpresa
    
    'Faz a Leitura dos Dados de Estatistica
    lErro = CF("Cliente_Le_Estatistica", objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52546
    
    Cliente_Le_Estendida = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estendida:
    
    Cliente_Le_Estendida = Err
    
    Select Case Err
        
        Case 52544, 52546  'Tratado na rotina chamada
        
        Case 52545 'Tratado na rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152695)

    End Select

    Exit Function
        
End Function

Function Cliente_Le_Estatistica_Credito(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le os Dados de Estatistica do Cliente para Credito
'Independe de Filial Empresa

Dim lErro As Long

On Error GoTo Erro_Cliente_Le_Estatistica_Credito
        
    'Le a soma dos saldos de TitulosReceber
    lErro = Cliente_Le_Est_TitRec_Global(objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 58876 Then Error 58870

    'Le a soma de NotasFiscais de Venda (faturáveis) que não estão associadas a Fatura
    lErro = Cliente_Le_Est_NFsNaoFaturadas_Global(objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 58880 Then Error 58871
    
    'Le os dados de estatistica com relação ao PV
    lErro = Cliente_Le_Est_PV_Global(objClienteEstatistica)
    If lErro <> SUCESSO Then Error 58872
    
    Cliente_Le_Estatistica_Credito = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_Credito:
    
    Cliente_Le_Estatistica_Credito = Err
    
    Select Case Err
        
        Case 58870, 58871, 58872 'Tratado nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152696)

    End Select

    Exit Function
        
End Function

Private Function Cliente_Le_Est_TitRec_Global(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le o Saldo dos Titulos do Cliente em questão
'Filtro: objClienteEstatistica.lCodCliente
'Independe de Filial Empresa

Dim lErro As Long
Dim dSaldoTitulos As Double
Dim lComando As Long
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Est_TitRec_Global
            
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 58873
    
    lErro = CF("FilialEmpresa_Le_Filtro_Filial2", sFiltro)
    If lErro <> SUCESSO Then gError 202730
               
    'le o saldo dos Titulos em TitulosRec
    lErro = Comando_Executar(lComando, "SELECT SUM(Saldo) FROM TitulosRec WHERE Cliente =? AND Status <> ?" & sFiltro, dSaldoTitulos, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 58874

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58875

    If lErro = AD_SQL_SEM_DADOS Then gError 58876
    
    objClienteEstatistica.dSaldoTitulos = dSaldoTitulos
    
    'Libera comandos
    Call Comando_Fechar(lComando)
        
    Cliente_Le_Est_TitRec_Global = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Est_TitRec_Global:
    
    Cliente_Le_Est_TitRec_Global = Err

    Select Case gErr
            
        Case 58873
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 58874, 58875
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSARECEBER", gErr)
        
        Case 58876 'tratado na rotina chamadora
        
        Case 202730
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152697)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Cliente_Le_Est_NFsNaoFaturadas_Global(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le a soma dos valores das NFs faturáveis não faturadas do Cliente em questão
'Filtro: objClienteEstatistica.lCodCliente
'Independe de Filial Empresa

Dim lErro As Long
Dim dValorNFsNaoFaturadas As Double
Dim lComando As Long
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Est_NFsNaoFaturadas_Global
            
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 58877

    lErro = CF("FilialEmpresa_Le_Filtro_Filial2", sFiltro, "NFiscal")
    If lErro <> SUCESSO Then gError 202730
    
    'Le o saldo das NFs (faturáveis) não faturadas em NFiscal
    lErro = Comando_Executar(lComando, "SELECT SUM(ValorTotal) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturavel = ? AND NumIntDocCPR = 0 AND Cliente =? AND Status <> ?" & sFiltro, dValorNFsNaoFaturadas, TIPODOCINFO_FATURAVEL, objClienteEstatistica.lCodCliente, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 58878

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58879
    If lErro = AD_SQL_SEM_DADOS Then gError 58880
    
    objClienteEstatistica.dValorNFsNaoFaturadas = dValorNFsNaoFaturadas
    
    'Libera comandos
    Call Comando_Fechar(lComando)
        
    Cliente_Le_Est_NFsNaoFaturadas_Global = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Est_NFsNaoFaturadas_Global:
    
    Cliente_Le_Est_NFsNaoFaturadas_Global = gErr

    Select Case gErr
            
        Case 58877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 58878, 58879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
        
        Case 58880 'tratado na rotina chamadora
        
        Case 202730
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152698)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Cliente_Le_Est_PV_Global(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Esta função lê estatísticas com relação aos pedidos de vendas
'Independe de Filial Empresa

Dim lErro As Long
Dim tItemPedido As typeItemPedido
Dim colItemPedido As New colItemPedido
Dim objItemPedido As New ClassItemPedido
Dim colCodPedido As New Collection
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim colPedidoVendaCompl As New Collection
Dim lCodPedBloqueio As Long
Dim lComando As Long
Dim lComando1 As Long
Dim vCodPedido As Variant
Dim iFilialEmpresa As Integer, iFilEmpAnt As Integer
Dim lCodPedido As Long, lCodPedidoAnt As Long
Dim sNumIntItemPV As String, lNumIntPV As Long
Dim dValorTotal As Double, dValorProdutos As Double, dValorDesconto As Double
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Est_PV_Global
    
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 58881
       
    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 58882
    
    lErro = CF("FilialEmpresa_Le_Filtro_Filial2", sFiltro, "PedidosDeVenda")
    If lErro <> SUCESSO Then gError 202730
       
    'Seleciona todos os Pedidos de Venda do Cliente
    lErro = Comando_Executar(lComando, "SELECT PedidosDeVenda.NumIntDoc, PedidosDeVenda.ValorTotal, PedidosDeVenda.ValorProdutos, PedidosDeVenda.ValorDesconto, PedidosDeVenda.FilialEmpresa, CodPedido,PrecoTotal,Quantidade,QuantCancelada,QuantFaturada, ItensPedidoDeVenda.NumIntDoc FROM PedidosDeVenda,ItensPedidoDeVenda WHERE PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.Cliente = ? " & sFiltro & " ORDER BY PedidosDeVenda.FilialEmpresa, PedidosDeVenda.Codigo", lNumIntPV, dValorTotal, dValorProdutos, dValorDesconto, iFilialEmpresa, tItemPedido.lCodPedido, tItemPedido.dPrecoTotal, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.lNumIntDoc, objClienteEstatistica.lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 58883

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58884

    Do While lErro <> AD_SQL_SEM_DADOS
        
        'se é o 1o item do pedido
        If iFilialEmpresa <> iFilEmpAnt Or tItemPedido.lCodPedido <> lCodPedidoAnt Then
        
            If Abs(dValorTotal - dValorProdutos - dValorDesconto) > DELTA_VALORMONETARIO Then
                Set objPedidoDeVenda = New ClassPedidoDeVenda
                With objPedidoDeVenda
                    .iFilialEmpresa = iFilialEmpresa
                    .lCodigo = tItemPedido.lCodPedido
                    .dValorTotal = dValorTotal - dValorProdutos - dValorDesconto
                    .lNumIntDoc = lNumIntPV
                End With
                colPedidoVendaCompl.Add objPedidoDeVenda, CStr(lNumIntPV)
            End If
            
            iFilEmpAnt = iFilialEmpresa
            lCodPedidoAnt = tItemPedido.lCodPedido
            
        End If
        
        'key para a coleção
        sNumIntItemPV = CStr(tItemPedido.lNumIntDoc)
        
        'Joga na coleção
        colItemPedido.Add objClienteEstatistica.iFilialEmpresa, tItemPedido.lCodPedido, 0, "", tItemPedido.dQuantidade, 0, tItemPedido.dPrecoTotal, 0, DATA_NULA, "", 0, tItemPedido.dQuantCancelada, 0, Nothing, "", "", "", "", 0, tItemPedido.dQuantFaturada, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, tItemPedido.lNumIntDoc, sNumIntItemPV
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58885

    Loop
    
    'Seleciona todos os pedidos com bloqueio de crédito
    lErro = Comando_Executar(lComando1, "SELECT Codigo FROM  PedidosDeVenda,BloqueiosPV WHERE PedidosDeVenda.FilialEmpresa = BloqueiosPV.FilialEmpresa AND PedidosDeVenda.Codigo = BloqueiosPV.PedidoDeVenda AND PedidosDeVenda.Cliente = ? AND BloqueiosPV.TipoDeBloqueio =? AND BloqueiosPV.DataLib =?" & sFiltro, lCodPedBloqueio, objClienteEstatistica.lCodCliente, BLOQUEIO_CREDITO, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 58886

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58887

    Do While lErro <> AD_SQL_SEM_DADOS
        
        'preenche na coleção
        colCodPedido.Add lCodPedBloqueio
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 58888
    Loop

    'Retira da coleção de pedidos os pedidos com bloqueio de credito
    For Each vCodPedido In colCodPedido
    
        'Código do Pedido bloqueado por crédito
        lCodPedido = CLng(vCodPedido)
        
        For Each objItemPedido In colItemPedido

            If objItemPedido.lCodPedido = lCodPedido Then
                
                sNumIntItemPV = CStr(objItemPedido.lNumIntDoc)
                colItemPedido.Remove (sNumIntItemPV)
            
            End If
        
        Next
        
        For Each objPedidoDeVenda In colPedidoVendaCompl
            If objPedidoDeVenda.lCodigo = lCodPedido Then
                colPedidoVendaCompl.Remove (CStr(objPedidoDeVenda.lNumIntDoc))
            End If
        Next
        
    Next

    objClienteEstatistica.dSaldoPedidosLiberados = 0
    
    'Calcula o saldo dos pedidos liberados
    For Each objItemPedido In colItemPedido
        
        If (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) > 0 Then
            objClienteEstatistica.dSaldoPedidosLiberados = objClienteEstatistica.dSaldoPedidosLiberados + objItemPedido.dPrecoTotal * (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) / (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada)
        End If
    
    Next
    
    'considera os valores de frete/seguro/outras/descontos/ipi...
    For Each objPedidoDeVenda In colPedidoVendaCompl
        objClienteEstatistica.dSaldoPedidosLiberados = objClienteEstatistica.dSaldoPedidosLiberados + objPedidoDeVenda.dValorTotal
    Next

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Cliente_Le_Est_PV_Global = SUCESSO
        
    Exit Function
    
Erro_Cliente_Le_Est_PV_Global:

    Cliente_Le_Est_PV_Global = gErr

    Select Case gErr
        
        Case 58881, 58882
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 58883, 58884, 58885
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDO_PEDIDOVENDA", gErr)
         
        Case 58886, 58887, 58888
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_BLOQUEIOSPV", gErr)
            
        Case 202730
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152699)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
        
End Function

Function Cliente_Le_Estatistica(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le os Dados de Estatistica do Cliente
'Depende de Filial Empresa

Dim lErro As Long

On Error GoTo Erro_Cliente_le_Estatistica
        
    'Le a soma dos saldos de TitulosReceber
    lErro = Cliente_Le_Estatistica_TitRec(objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52686 Then Error 52682

    'Le a soma de NotasFiscais de Venda (faturáveis) que não estão associadas a Fatura
    lErro = Cliente_Le_Estatistica_NFsNaoFaturadas(objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 25650 Then Error 25646
    
    'Le o Saldo de Duplicata do Cliente em questão
    lErro = Cliente_Le_Estatistica_Duplicatas(objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52555 Then Error 52547
    
    'Le os dados de estatistica com relação ao PV
    lErro = Cliente_Le_Estatistica_PV(objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52548
    
    'Le os dados estatisticos do cliente com relação à atraso
    lErro = CF("Cliente_Le_Estatistica_Atraso", objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52549
    
    'Le os dados estatisticos com relação a cheques devolvidos
    lErro = Cliente_Le_Estatistica_ChequeDev(objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52550
        
    'Le os dados estatisticos com relação a compras
    lErro = Cliente_Le_Estatistica_Compras(objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52551
    
    Cliente_Le_Estatistica = SUCESSO
    
    Exit Function
    
Erro_Cliente_le_Estatistica:
    
    Cliente_Le_Estatistica = Err
    
    Select Case Err
        
        Case 25646, 52547, 52548, 52549, 52550, 52551, 52682 'Tratado nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152700)

    End Select

    Exit Function
        
End Function


Private Function Cliente_Le_Estatistica_TitRec(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le o Saldo dos Titulos do Cliente em questão
'Filtro: objClienteEstatistica.lCodCliente , objClienteEstatistica.iFilialEmpresa

Dim lErro As Long
Dim dSaldoTitulos As Double
Dim iIndice As Integer
Dim alComando(1 To 5) As Long
Dim sFiltro As String, dValorTotalCR As Double
Dim dValorTotalCRComProtesto As Double, dValorTotalCREmCartorio As Double
Dim dValorTotalNaoPagoAteVenc As Double

On Error GoTo Erro_Cliente_Le_Estatistica_TitRec
            
    'Inicializa comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 52683
    Next
               
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
    
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
    End If
    
    'le o saldo dos Titulos em TitulosRec
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(1), "SELECT SUM(Saldo) FROM TitulosRec WHERE Cliente =? AND Status <> ?" & sFiltro, dSaldoTitulos, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO)
    Else
        lErro = Comando_Executar(alComando(1), "SELECT SUM(Saldo) FROM TitulosRec WHERE Cliente =? AND FilialEmpresa= ? AND Status <> ?", dSaldoTitulos, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52684

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52685

    If lErro = AD_SQL_SEM_DADOS Then gError 52686
    
    objClienteEstatistica.dSaldoTitulos = dSaldoTitulos
    
    'le o valor dos Titulos em TitulosRec
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) FROM TitulosRecTodos WHERE Cliente =? AND Status <> ?" & sFiltro, dValorTotalCR, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO)
    Else
        lErro = Comando_Executar(alComando(2), "SELECT SUM(Valor) FROM TitulosRecTodos WHERE Cliente =? AND FilialEmpresa= ? AND Status <> ?", dValorTotalCR, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52684

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52685

    If lErro = AD_SQL_SEM_DADOS Then gError 52686
    
    objClienteEstatistica.dValorTotalCR = dValorTotalCR
    
    'le o valor dos Titulos com protesto
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(3), "SELECT SUM(Valor) FROM TitRecBaixaComProtesto WHERE Cliente =? " & sFiltro, dValorTotalCRComProtesto, objClienteEstatistica.lCodCliente)
    Else
        lErro = Comando_Executar(alComando(3), "SELECT SUM(Valor) FROM TitRecBaixaComProtesto WHERE Cliente =? AND FilialEmpresa= ? ", dValorTotalCRComProtesto, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52684

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52685

    If lErro = AD_SQL_SEM_DADOS Then gError 52686
    
    objClienteEstatistica.dValorTotalCRComProtesto = dValorTotalCRComProtesto
    
    'le o valor dos Titulos em cartório
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(alComando(4), "SELECT SUM(Valor) FROM TitRecBaixaComCartorio WHERE Cliente =? " & sFiltro, dValorTotalCREmCartorio, objClienteEstatistica.lCodCliente)
    Else
        lErro = Comando_Executar(alComando(4), "SELECT SUM(Valor) FROM TitRecBaixaComCartorio WHERE Cliente =? AND FilialEmpresa= ? ", dValorTotalCREmCartorio, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52684

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52685

    If lErro = AD_SQL_SEM_DADOS Then gError 52686
    
    objClienteEstatistica.dValorTotalCREmCartorio = dValorTotalCREmCartorio
'
'    'le o valor dos Titulos Não pagos até o vencimento
'    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
'        lErro = Comando_Executar(alComando(5), "SELECT SUM(Valor) FROM TitRecBaixaAux01 WHERE Cliente =? AND Atraso > 0 " & sFiltro, dValorTotalNaoPagoAteVenc, objClienteEstatistica.lCodCliente)
'    Else
'        lErro = Comando_Executar(alComando(5), "SELECT SUM(Valor) FROM TitRecBaixaAux01 WHERE Cliente =? AND FilialEmpresa= ? AND Atraso > 0 ", dValorTotalNaoPagoAteVenc, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
'    End If
'    If lErro <> AD_SQL_SUCESSO Then gError 52684
'
'    lErro = Comando_BuscarPrimeiro(alComando(5))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52685
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 52686
'
'    objClienteEstatistica.dValorTotalNaoPagoAteVenc = dValorTotalNaoPagoAteVenc
    
    'Libera comandos
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
        
    Cliente_Le_Estatistica_TitRec = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_TitRec:
    
    Cliente_Le_Estatistica_TitRec = gErr

    Select Case gErr
            
        Case 52683
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 52684, 52685
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC3", gErr, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
        
        Case 52686 'tratado na rotina chamadora
        
        Case 177597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152701)

    End Select
    
    'Libera comandos
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Cliente_Le_Estatistica_NFsNaoFaturadas(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le a soma dos valores das NFs faturáveis não faturadas do Cliente em questão
'Filtro: objClienteEstatistica.lCodCliente , objClienteEstatistica.iFilialEmpresa

Dim lErro As Long
Dim dValorNFsNaoFaturadas As Double
Dim lComando As Long
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Estatistica_NFsNaoFaturadas
            
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 25647
               
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        'Le o saldo das NFs (faturáveis) não faturadas em NFiscal
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorTotal) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturavel = ? AND NumIntDocCPR = 0 AND Cliente =? AND Status <> ? AND Status <> ? AND Status <> ?" & sFiltro, dValorNFsNaoFaturadas, TIPODOCINFO_FATURAVEL, objClienteEstatistica.lCodCliente, STATUS_BAIXADO, STATUS_CANCELADO, STATUS_EXCLUIDO)
    Else
        'Le o saldo das NFs (faturáveis) não faturadas em NFiscal
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorTotal) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturavel = ? AND NumIntDocCPR = 0 AND Cliente =? AND FilialEmpresa= ? AND Status <> ? AND Status <> ? AND Status <> ?", dValorNFsNaoFaturadas, TIPODOCINFO_FATURAVEL, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa, STATUS_BAIXADO, STATUS_CANCELADO, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 25648
    '#########################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25649
    If lErro = AD_SQL_SEM_DADOS Then gError 25650
    
    objClienteEstatistica.dValorNFsNaoFaturadas = dValorNFsNaoFaturadas
    
    'Libera comandos
    Call Comando_Fechar(lComando)
        
    Cliente_Le_Estatistica_NFsNaoFaturadas = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_NFsNaoFaturadas:
    
    Cliente_Le_Estatistica_NFsNaoFaturadas = gErr

    Select Case gErr
            
        Case 25647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 25648, 25649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
        
        Case 25650 'tratado na rotina chamadora
        
        Case 177597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152702)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Private Function Cliente_Le_Estatistica_Duplicatas(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le o Saldo de Duplicata do Cliente em questão
'Filtro: objClienteEstatistica.lCodCliente , objClienteEstatistica.iFilialEmpresa

Dim lErro As Long
Dim dSaldoDuplicatas As Double
Dim lComando As Long
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Estatistica_Duplicatas
            
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 52552
               
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
    
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
    
    'le o saldo
        lErro = Comando_Executar(lComando, "SELECT SUM(ParcelasRec.Saldo) FROM TitulosRec, ParcelasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND TitulosRec.Cliente =? AND ParcelasRec.CarteiraCobranca = ? " & sFiltro, dSaldoDuplicatas, objClienteEstatistica.lCodCliente, CARTEIRA_DESCONTADA)
    Else
    'le o saldo
        lErro = Comando_Executar(lComando, "SELECT SUM(ParcelasRec.Saldo) FROM TitulosRec, ParcelasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND TitulosRec.Cliente =? AND TitulosRec.FilialEmpresa= ?  AND ParcelasRec.CarteiraCobranca = ?", dSaldoDuplicatas, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa, CARTEIRA_DESCONTADA)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52553
    '#########################################

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52554

    If lErro = AD_SQL_SEM_DADOS Then gError 52555
    
    objClienteEstatistica.dSaldoDuplicatas = dSaldoDuplicatas
    
    'Libera comandos
    Call Comando_Fechar(lComando)
        
    Cliente_Le_Estatistica_Duplicatas = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_Duplicatas:
    
    Cliente_Le_Estatistica_Duplicatas = gErr

    Select Case gErr
            
        Case 52552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 52553, 52554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC_TITULOSREC_SALDO", gErr, objClienteEstatistica.lCodCliente)
        
        Case 52555 'tratado na rotina chamadora
        
        Case 177597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152703)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function

Private Function Cliente_Le_Estatistica_PV(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Esta função lê estatísticas com relação aos pedidos de vendas

Dim lErro As Long
Dim tItemPedido As typeItemPedido
Dim colItemPedido As New colItemPedido
Dim objItemPedido As New ClassItemPedido
Dim colCodPedido As New Collection
Dim lCodPedBloqueio As Long
Dim lComando As Long
Dim lComando1 As Long
Dim vCodPedido As Variant
Dim lCodPedido As Long
Dim sNumIntItemPV As String
Dim sFiltro As String
Dim sFiltro2 As String

On Error GoTo Erro_Cliente_Le_Estatistica_PV
    
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 52556
       
    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 52557
       
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        If Len(Trim(sFiltro)) <> 0 Then
            sFiltro = " AND PedidosDeVenda." & right(sFiltro, Len(sFiltro) - 5)
        End If
        
        'Seleciona todos os Pedidos de Venda do Cliente
        lErro = Comando_Executar(lComando, "SELECT CodPedido,PrecoTotal,Quantidade,QuantCancelada,QuantFaturada, ItensPedidoDeVenda.NumIntDoc FROM  PedidosDeVenda,ItensPedidoDeVenda WHERE PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.Cliente = ?" & sFiltro, tItemPedido.lCodPedido, tItemPedido.dPrecoTotal, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.lNumIntDoc, objClienteEstatistica.lCodCliente)
    Else
        'Seleciona todos os Pedidos de Venda do Cliente e da objClienteEstatistica.iFilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT CodPedido,PrecoTotal,Quantidade,QuantCancelada,QuantFaturada, ItensPedidoDeVenda.NumIntDoc FROM  PedidosDeVenda,ItensPedidoDeVenda WHERE PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.Cliente = ? AND PedidosDeVenda.FilialEmpresa = ?", tItemPedido.lCodPedido, tItemPedido.dPrecoTotal, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.lNumIntDoc, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52558
    '#########################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52559

    Do While lErro <> AD_SQL_SEM_DADOS
        
        'key para a coleção
        sNumIntItemPV = CStr(tItemPedido.lNumIntDoc)
        
        'Joga na coleção
        colItemPedido.Add objClienteEstatistica.iFilialEmpresa, tItemPedido.lCodPedido, 0, "", tItemPedido.dQuantidade, 0, tItemPedido.dPrecoTotal, 0, DATA_NULA, "", 0, tItemPedido.dQuantCancelada, 0, Nothing, "", "", "", "", 0, tItemPedido.dQuantFaturada, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, tItemPedido.lNumIntDoc, sNumIntItemPV
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52560

    Loop
    
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
    
        'Seleciona todos os pedidos com bloqueio de crédito
        lErro = Comando_Executar(lComando1, "SELECT Codigo FROM  PedidosDeVenda,BloqueiosPV WHERE PedidosDeVenda.FilialEmpresa = BloqueiosPV.FilialEmpresa AND PedidosDeVenda.Codigo = BloqueiosPV.PedidoDeVenda AND PedidosDeVenda.Cliente = ? AND BloqueiosPV.TipoDeBloqueio =? AND BloqueiosPV.DataLib =?" & sFiltro2, lCodPedBloqueio, objClienteEstatistica.lCodCliente, BLOQUEIO_CREDITO, DATA_NULA)
    Else
        'Seleciona todos os pedidos com bloqueio de crédito
        lErro = Comando_Executar(lComando1, "SELECT Codigo FROM  PedidosDeVenda,BloqueiosPV WHERE PedidosDeVenda.FilialEmpresa = BloqueiosPV.FilialEmpresa AND PedidosDeVenda.Codigo = BloqueiosPV.PedidoDeVenda AND PedidosDeVenda.Cliente = ? AND PedidosDeVenda.FilialEmpresa = ? AND BloqueiosPV.TipoDeBloqueio =? AND BloqueiosPV.DataLib =?", lCodPedBloqueio, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa, BLOQUEIO_CREDITO, DATA_NULA)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52561
    '#########################################

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52562

    Do While lErro <> AD_SQL_SEM_DADOS
        
        'preenche na coleção
        colCodPedido.Add lCodPedBloqueio
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52563
    Loop

    'Retira da coleção de pedidos os pedidos com bloqueio
    For Each vCodPedido In colCodPedido
    
        'Código do Pedido bloqueado por crédito
        lCodPedido = CLng(vCodPedido)
        
        For Each objItemPedido In colItemPedido

            If objItemPedido.lCodPedido = lCodPedido Then
                
                sNumIntItemPV = CStr(objItemPedido.lNumIntDoc)
                colItemPedido.Remove (sNumIntItemPV)
            
            End If
        
        Next
    Next

    objClienteEstatistica.dSaldoPedidosLiberados = 0
    
    'Calcula o saldo dos pedidos liberados
    For Each objItemPedido In colItemPedido
        
        If (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) > 0 Then
            objClienteEstatistica.dSaldoPedidosLiberados = objClienteEstatistica.dSaldoPedidosLiberados + objItemPedido.dPrecoTotal * (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) / (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada)
        End If
    
    Next
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Cliente_Le_Estatistica_PV = SUCESSO
        
    Exit Function
    
Erro_Cliente_Le_Estatistica_PV:

    Cliente_Le_Estatistica_PV = gErr

    Select Case gErr
        
        Case 52556, 52557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 52558, 52559, 52560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDO_PEDIDOVENDA", gErr)
         
        Case 52561, 52562, 52563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_BLOQUEIOSPV", gErr)
         
        Case 177597
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152704)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
        
End Function

Public Function Cliente_Le_Estatistica_Atraso(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Função que calcula a estatística relacionada com atraso do cliente

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim dValPagtosAtraso1 As Double
Dim dValPagtosAtraso2 As Double
Dim dDiasAtrasoValor As Double, dDiasAtrasoValor1 As Double, dDiasAtrasoValor2 As Double
Dim dValPagtosEmDia As Double, dValPagtosEmDia1 As Double, dValPagtosEmDia2 As Double
Dim lMaiorAtraso1 As Long, lMaiorAtraso2 As Long
Dim dSaldoAtrasados As Double
Dim iIndice As Integer, bSuportaTimeStampDiff As Boolean
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Estatistica_Atraso

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 52564
    Next
    
    lErro = Conexao_Suporta_TimeStampDiff(bSuportaTimeStampDiff, GL_lConexao)
    If lErro <> SUCESSO Then gError 59284
    
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        'Seleciona nas baixas de Títulos Receber
        'soma(valores baixados), maior atraso(em dias), soma(dias de atraso * valores atrasados)
        If bSuportaTimeStampDiff Then
            lErro = Comando_Executar(alComando(0), "SELECT SUM(ValorBaixado), MAX( {fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) } ), SUM( {fn TIMESTAMPDIFF(SQL_TSI_DAY, DataVencimentoReal, Data ) } * ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND TitulosRec.Cliente = ? AND ParcelasRec.DataVencimento < BaixasRec.Data AND TitulosRec.Status <> ? AND BaixasParcRec.Status <> ?" & sFiltro, dValPagtosAtraso1, lMaiorAtraso1, dDiasAtrasoValor1, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        Else
            lErro = Comando_Executar(alComando(0), "SELECT SUM(ValorBaixado), MAX( Data - DataVencimentoReal ), SUM( ( Data - DataVencimentoReal ) * ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND TitulosRec.Cliente = ? AND ParcelasRec.DataVencimento < BaixasRec.Data AND TitulosRec.Status <> ? AND BaixasParcRec.Status <> ?" & sFiltro, dValPagtosAtraso1, lMaiorAtraso1, dDiasAtrasoValor1, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        End If
    Else
        'Seleciona nas baixas de Títulos Receber
        'soma(valores baixados), maior atraso(em dias), soma(dias de atraso * valores atrasados)
        If bSuportaTimeStampDiff Then
            lErro = Comando_Executar(alComando(0), "SELECT SUM(ValorBaixado), MAX( {fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) } ), SUM( {fn TIMESTAMPDIFF(SQL_TSI_DAY, DataVencimentoReal, Data ) } * ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND ParcelasRec.DataVencimento < BaixasRec.Data AND TitulosRec.Status <> ? AND BaixasParcRec.Status <> ?", dValPagtosAtraso1, lMaiorAtraso1, dDiasAtrasoValor1, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        Else
            lErro = Comando_Executar(alComando(0), "SELECT SUM(ValorBaixado), MAX( Data - DataVencimentoReal ), SUM( ( Data - DataVencimentoReal ) * ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND ParcelasRec.DataVencimento < BaixasRec.Data AND TitulosRec.Status <> ? AND BaixasParcRec.Status <> ?", dValPagtosAtraso1, lMaiorAtraso1, dDiasAtrasoValor1, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52565
    '#########################################
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52566

    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        'Seleciona nas baixas de Títulos Receber Baixados
        'soma(valores baixados), maior atraso(em dias), soma(dias de atraso * valores atrasados)
        If bSuportaTimeStampDiff Then
            lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorBaixado), MAX( {fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) } ), SUM({fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) }* ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND TitulosRecBaixados.Cliente = ? AND ParcelasRecBaixadas.DataVencimento < BaixasRec.Data AND TitulosRecBaixados.Status <> ? AND BaixasParcRec.Status <> ?" & sFiltro, dValPagtosAtraso2, lMaiorAtraso2, dDiasAtrasoValor2, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        Else
            lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorBaixado), MAX( Data - DataVencimentoReal ), SUM( ( Data - DataVencimentoReal ) * ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND TitulosRecBaixados.Cliente = ? AND ParcelasRecBaixadas.DataVencimento < BaixasRec.Data AND TitulosRecBaixados.Status <> ? AND BaixasParcRec.Status <> ?" & sFiltro, dValPagtosAtraso2, lMaiorAtraso2, dDiasAtrasoValor2, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        End If
    Else
        'Seleciona nas baixas de Títulos Receber Baixados
        'soma(valores baixados), maior atraso(em dias), soma(dias de atraso * valores atrasados)
        If bSuportaTimeStampDiff Then
            lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorBaixado), MAX( {fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) } ), SUM({fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) }* ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRecBaixados.Cliente = ? AND ParcelasRecBaixadas.DataVencimento < BaixasRec.Data AND TitulosRecBaixados.Status <> ? AND BaixasParcRec.Status <> ?", dValPagtosAtraso2, lMaiorAtraso2, dDiasAtrasoValor2, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        Else
            lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorBaixado), MAX( Data - DataVencimentoReal ), SUM( ( Data - DataVencimentoReal ) * ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRecBaixados.Cliente = ? AND ParcelasRecBaixadas.DataVencimento < BaixasRec.Data AND TitulosRecBaixados.Status <> ? AND BaixasParcRec.Status <> ?", dValPagtosAtraso2, lMaiorAtraso2, dDiasAtrasoValor2, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 52567
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52567
    '#########################################

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52568
        
    objClienteEstatistica.dValPagtosAtraso = dValPagtosAtraso1 + dValPagtosAtraso2
    
    dDiasAtrasoValor = dDiasAtrasoValor1 + dDiasAtrasoValor2
    
    If lMaiorAtraso1 > lMaiorAtraso2 Then
        objClienteEstatistica.lMaiorAtraso = lMaiorAtraso1
    Else
        objClienteEstatistica.lMaiorAtraso = lMaiorAtraso2
    End If
    
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        'Seleciona soma(valores baixados) de Titulos Receber que foram pagos sem atraso
        lErro = Comando_Executar(alComando(2), "SELECT SUM(ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND TitulosRec.Cliente = ? AND DataVencimento >= Data AND TitulosRec.Status <> ? AND BaixasParcRec.Status <> ?" & sFiltro, dValPagtosEmDia1, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
    Else
        'Seleciona soma(valores baixados) de Titulos Receber que foram pagos sem atraso
        lErro = Comando_Executar(alComando(2), "SELECT SUM(ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND DataVencimento >= Data AND TitulosRec.Status <> ? AND BaixasParcRec.Status <> ?", dValPagtosEmDia1, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52569
    '#########################################
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52570

    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        'Seleciona soma(valores baixados) de Titulos Receber Baixados que foram pagos sem atraso
        lErro = Comando_Executar(alComando(3), "SELECT SUM(ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND Cliente = ? AND DataVencimento >= Data AND TitulosRecBaixados.Status <> ? AND BaixasParcRec.Status <> ?" & sFiltro, dValPagtosEmDia1, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
    Else
        'Seleciona soma(valores baixados) de Titulos Receber Baixados que foram pagos sem atraso
        lErro = Comando_Executar(alComando(3), "SELECT SUM(ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND Cliente = ? AND DataVencimento >= Data AND TitulosRecBaixados.Status <> ? AND BaixasParcRec.Status <> ?", dValPagtosEmDia1, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52571
    '#########################################

    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52572
    
    dValPagtosEmDia = dValPagtosEmDia1 + dValPagtosEmDia2
    
    If dDiasAtrasoValor <> 0 Then
        objClienteEstatistica.lMediaAtraso = dDiasAtrasoValor / (objClienteEstatistica.dValPagtosAtraso + dValPagtosEmDia)
    Else
        objClienteEstatistica.lMediaAtraso = 0
    End If
    
    '#########################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        'Seleciona o saldo de títulos receber atrasados
        lErro = Comando_Executar(alComando(4), "SELECT SUM(ParcelasRec.Saldo) FROM TitulosRec, ParcelasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND TitulosRec.Cliente = ? AND TitulosRec.Saldo > 0  AND ParcelasRec.DataVencimentoReal < ? AND TitulosRec.Status <> ?" & sFiltro, dSaldoAtrasados, objClienteEstatistica.lCodCliente, gdtDataAtual, STATUS_EXCLUIDO)
    Else
        'Seleciona o saldo de títulos receber atrasados
        lErro = Comando_Executar(alComando(4), "SELECT SUM(ParcelasRec.Saldo) FROM TitulosRec, ParcelasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND TitulosRec.Saldo > 0  AND ParcelasRec.DataVencimentoReal < ? AND TitulosRec.Status <> ?", dSaldoAtrasados, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, gdtDataAtual, STATUS_EXCLUIDO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52573
    '#########################################
       
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52574
       
    objClienteEstatistica.dSaldoAtrasados = dSaldoAtrasados
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Cliente_Le_Estatistica_Atraso = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_Atraso:
    
    Cliente_Le_Estatistica_Atraso = gErr
    
    Select Case gErr
                    
        Case 52564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 52565, 52566, 52569, 52570, 52573, 52574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC3", gErr, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
        
        Case 52567, 52568, 52571, 52572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS1", gErr, objClienteEstatistica.lCodCliente, objClienteEstatistica.iFilialEmpresa)
                        
        Case 59284
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152705)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Private Function Cliente_Le_Estatistica_ChequeDev(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Função que calcula estatística com relação a cheques devolvidos

Dim lErro As Long
Dim dtDataUltChequeDevolvido  As Date
Dim lNumChequesDevolvidos As Long
Dim lComando  As Long
Dim lComando2  As Long
Dim sFiltro As String

On Error GoTo Erro_Cliente_Le_Estatistica_ChequeDev
    
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 52575
        
    'Inicializa comandos
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 52576
    
    '#######################################################
    'Alterado por Wagner
    If objClienteEstatistica.iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        'le o saldo dos Titulos em TitulosRec
        'Seleciona para BaixasTitRec data do último cheque devolvido e a quantidade de cheque devolvidos
        'Seleciona a partir MovCC_BaixaPR_TitRec
        'lErro = Comando_Executar(lComando, "SELECT MAX(DataMovimento), COUNT(*) FROM MovCC_BaixaPR_TitRec WHERE Cliente = ? AND Status = ?" & sFiltro, dtDataUltChequeDevolvido, lNumChequesDevolvidos, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO)
        lErro = Comando_Executar(lComando, "SELECT MAX(Data), COUNT(*) FROM ChequesDevolvidos WHERE Cliente = ? " & sFiltro, dtDataUltChequeDevolvido, lNumChequesDevolvidos, objClienteEstatistica.lCodCliente)
    
    Else
        'le o saldo dos Titulos em TitulosRec
        'Seleciona para BaixasTitRec data do último cheque devolvido e a quantidade de cheque devolvidos
        'Seleciona a partir MovCC_BaixaPR_TitRec
        'lErro = Comando_Executar(lComando, "SELECT MAX(DataMovimento), COUNT(*) FROM MovCC_BaixaPR_TitRec WHERE FilialEmpresa =? AND Cliente = ? AND Status = ?", dtDataUltChequeDevolvido, lNumChequesDevolvidos, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente, STATUS_EXCLUIDO)
        lErro = Comando_Executar(lComando, "SELECT MAX(Data), COUNT(*) FROM ChequesDevolvidos WHERE FilialEmpresa = ? AND Cliente = ?", dtDataUltChequeDevolvido, lNumChequesDevolvidos, objClienteEstatistica.iFilialEmpresa, objClienteEstatistica.lCodCliente)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52577
    '#######################################################
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52578
    
    objClienteEstatistica.dtDataUltChequeDevolvido = dtDataUltChequeDevolvido
    objClienteEstatistica.iNumChequesDevolvidos = CInt(lNumChequesDevolvidos)
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Cliente_Le_Estatistica_ChequeDev = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_ChequeDev:

    Cliente_Le_Estatistica_ChequeDev = gErr
    
    Select Case gErr
        
        Case 52575, 52576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 52577, 52578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVCC_BAIXAPR_TITREC", gErr)
            
        Case 177597
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152706)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
        
End Function

Private Function Cliente_Le_Estatistica_Compras(ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Função calcula a estatística de compras

Dim lErro As Long
Dim objCliFilEmp As New ClassFilCliFilEmp

On Error GoTo Erro_Cliente_Le_Estatistica_Compras
        
    objCliFilEmp.lCodCliente = objClienteEstatistica.lCodCliente
    objCliFilEmp.iFilialEmpresa = objClienteEstatistica.iFilialEmpresa
    
    'Lê na tabela FilialClienteFilEmp
    lErro = CliFilEmp_Le(objCliFilEmp)
    If lErro <> SUCESSO Then Error 52581
    
    objClienteEstatistica.lNumeroCompras = objCliFilEmp.lNumeroCompras
    objClienteEstatistica.dMediaCompra = objCliFilEmp.dMediaCompra
    objClienteEstatistica.dtDataPrimeiraCompra = objCliFilEmp.dtDataPrimeiraCompra
    objClienteEstatistica.dtDataUltimaCompra = objCliFilEmp.dtDataUltimaCompra
    objClienteEstatistica.dValorAcumuladoCompras = objCliFilEmp.dValorAcumuladoCompras
      
    Cliente_Le_Estatistica_Compras = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_Estatistica_Compras:

    Cliente_Le_Estatistica_Compras = Err
        
    Select Case Err
    
        Case 52581
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152707)

    End Select
    
    Exit Function
        
End Function

'Maristela(Inicio) Alteração: Não seleciona mais as Datas da tabela FilialClienteFilEmp
Function CliFilEmp_Le(ByVal objFilCliFilEmp As ClassFilCliFilEmp) As Long
'Le em FilialClienteFilEmp dados estatísticos de compras do Cliente.

Dim lErro As Long
Dim tFilCliFilEmp As typeFilCliFilEmp
Dim lComando  As Long
Dim lNumFiliaisCli As Long
Dim sFiltro As String

On Error GoTo Erro_CliFilEmp_Le
    
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 52583
    
    '#######################################################
    'Alterado por Wagner
    If objFilCliFilEmp.iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        'Lê a tabela FilCliFilEmp para trazer os dados
        lErro = Comando_Executar(lComando, "SELECT SUM(NumeroCompras), SUM(MediaCompra), COUNT(*), SUM(ValorAcumuladoCompras) FROM FilialClienteFilEmp WHERE CodCliente =? " & sFiltro, tFilCliFilEmp.lNumeroCompras, tFilCliFilEmp.dMediaCompra, lNumFiliaisCli, tFilCliFilEmp.dValorAcumuladoCompras, objFilCliFilEmp.lCodCliente)
    Else
        'Lê a tabela FilCliFilEmp para trazer os dados
        lErro = Comando_Executar(lComando, "SELECT SUM(NumeroCompras), SUM(MediaCompra), COUNT(*), SUM(ValorAcumuladoCompras) FROM FilialClienteFilEmp WHERE FilialEmpresa = ? AND CodCliente =? ", tFilCliFilEmp.lNumeroCompras, tFilCliFilEmp.dMediaCompra, lNumFiliaisCli, tFilCliFilEmp.dValorAcumuladoCompras, objFilCliFilEmp.iFilialEmpresa, objFilCliFilEmp.lCodCliente)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 52584
    '#######################################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 52585

    objFilCliFilEmp.lNumeroCompras = tFilCliFilEmp.lNumeroCompras
    objFilCliFilEmp.dValorAcumuladoCompras = tFilCliFilEmp.dValorAcumuladoCompras
    
    'pesquisa a primeira e última compra
    lErro = ClienteFilEmp_Pesquisa_Compra(objFilCliFilEmp)
    If lErro <> SUCESSO Then gError 90524
    
    If lNumFiliaisCli > 0 Then
    
        objFilCliFilEmp.dMediaCompra = tFilCliFilEmp.dMediaCompra / lNumFiliaisCli
    
    Else
    
        objFilCliFilEmp.dMediaCompra = 0
    
    End If
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    CliFilEmp_Le = SUCESSO

    Exit Function

    CliFilEmp_Le = SUCESSO

Erro_CliFilEmp_Le:

    CliFilEmp_Le = gErr

    Select Case gErr

        Case 52583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 52584, 52585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTEFILEMP", gErr)
     
        Case 90524, 177597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152708)
    
    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
'Maristela(Fim)

'Maristela(Inicio)
Function ClienteFilEmp_Pesquisa_Compra(ByVal objFilCliFilEmp As ClassFilCliFilEmp) As Long
'Pesquisa a primeira e a ultima nota fiscal da FilialEmpresa (data de emissao mais antiga e a mais nova) para o Cliente em questão.

Dim lErro As Long
Dim lComando As Long
Dim dtDataEmissaoPri As Date
Dim dtDataEmissaoUlt As Date
Dim tFilCliFilEmp As typeFilCliFilEmp
Dim sFiltro As String

On Error GoTo Erro_ClienteFilEmp_Pesquisa_Compra

     'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90521
    
    'Inicializa as Datas
    dtDataEmissaoPri = DATA_NULA
    dtDataEmissaoUlt = DATA_NULA
    
    '#######################################################
    'Alterado por Wagner
    If objFilCliFilEmp.iFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
        
        lErro = Comando_Executar(lComando, "SELECT MIN(NFiscal.DataEmissao), MAX(NFiscal.DataEmissao) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturamento = ? AND TiposDocInfo.Complementar = ? AND NFiscal.DataEmissao <> ? AND NFiscal.Cliente = ? " & sFiltro, dtDataEmissaoPri, dtDataEmissaoUlt, TIPODOCINFO_FATURAVEL, DOCINFO_NORMAL, DATA_NULA, objFilCliFilEmp.lCodCliente)
    Else
        lErro = Comando_Executar(lComando, "SELECT MIN(NFiscal.DataEmissao), MAX(NFiscal.DataEmissao) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturamento = ? AND TiposDocInfo.Complementar = ? AND NFiscal.DataEmissao <> ? AND NFiscal.Cliente = ? AND NFiscal.FilialEmpresa = ? ", dtDataEmissaoPri, dtDataEmissaoUlt, TIPODOCINFO_FATURAVEL, DOCINFO_NORMAL, DATA_NULA, objFilCliFilEmp.lCodCliente, objFilCliFilEmp.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 90522
    '#######################################################
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90523
    
    If lErro = AD_SQL_SUCESSO Then
        objFilCliFilEmp.dtDataPrimeiraCompra = dtDataEmissaoPri
        objFilCliFilEmp.dtDataUltimaCompra = dtDataEmissaoUlt
    End If
    
    Call Comando_Fechar(lComando)
       
    ClienteFilEmp_Pesquisa_Compra = SUCESSO
    
    Exit Function
    
Erro_ClienteFilEmp_Pesquisa_Compra:

    ClienteFilEmp_Pesquisa_Compra = gErr
    
    Select Case gErr
        
        Case 90521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90522, 90523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
            
        Case 177597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152709)
        
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
       
    Exit Function

End Function
'Maristela(Fim)

Function Cliente_Le_NomeRed_Estendida(ByVal objCliente As ClassCliente, ByVal objClienteEstatistica As ClassFilialClienteEst) As Long
'Le todos os Dados do cliente inclusive os dados de Estatistica com o NomeReduzido sendo passado como parametro

Dim lErro As Long

On Error GoTo Erro_Cliente_Le_NomeRed_Estendida
    
    'Faz a Leitura na Tabela de cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 52692
    
    'Se não encontrou  - - -> Erro
    If lErro = 12348 Then Error 52693
    
    objClienteEstatistica.lCodCliente = objCliente.lCodigo
    objClienteEstatistica.iFilialEmpresa = giFilialEmpresa
    
    'Faz a Leitura dos Dados de Estatistica
    lErro = CF("Cliente_Le_Estatistica", objClienteEstatistica)
    If lErro <> SUCESSO Then Error 52694
    
    Cliente_Le_NomeRed_Estendida = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_NomeRed_Estendida:
    
    Cliente_Le_NomeRed_Estendida = Err
    
    Select Case Err
        
        Case 52692, 52694  'Tratado na rotina chamada
        
        Case 52693 'Tratado na Rotina Chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152710)

    End Select

    Exit Function
        
End Function

'Maristela(Inicio) Alteração: Não seleciona mais as Datas da tabela FilialClienteFilEmp
Function FilCliFilEmp_Le(ByVal objFilCliFilEmp As ClassFilCliFilEmp) As Long
'Le em FilialClienteFilEmp os dados estatísticos de compras da filial cliente

Dim lErro As Long
Dim tFilCliFilEmp As typeFilCliFilEmp
Dim lComando  As Long

On Error GoTo Erro_FilCliFilEmp_Le

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90509
    
    'Lê a tabela FilCliFilEmp para trazer os dados
    lErro = Comando_Executar(lComando, "SELECT NumeroCompras, MediaCompra, ValorAcumuladoCompras FROM FilialClienteFilEmp WHERE FilialEmpresa = ? AND CodCliente =? AND CodFilial=?", tFilCliFilEmp.lNumeroCompras, tFilCliFilEmp.dMediaCompra, tFilCliFilEmp.dValorAcumuladoCompras, objFilCliFilEmp.iFilialEmpresa, objFilCliFilEmp.lCodCliente, objFilCliFilEmp.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 90510

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90511

    objFilCliFilEmp.lNumeroCompras = tFilCliFilEmp.lNumeroCompras
    objFilCliFilEmp.dValorAcumuladoCompras = tFilCliFilEmp.dValorAcumuladoCompras
    objFilCliFilEmp.dMediaCompra = tFilCliFilEmp.dMediaCompra
    
    'pesquisa a primeira e última compra
    lErro = FilialClienteFilEmp_Pesquisa_Compra(objFilCliFilEmp)
    If lErro <> SUCESSO Then gError 90512
            
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    FilCliFilEmp_Le = SUCESSO

    Exit Function

    FilCliFilEmp_Le = SUCESSO

Erro_FilCliFilEmp_Le:

    FilCliFilEmp_Le = gErr

    Select Case gErr

        Case 90509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90510, 90511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTEFILEMP", gErr)

        Case 90512
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152711)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
'Maristela(Fim)

'Maristela (Inicio)
Function FilialClienteFilEmp_Pesquisa_Compra(ByVal objFilCliFilEmp As ClassFilCliFilEmp) As Long
'Pesquisa a primeira e a ultima nota fiscal da FilialEmpresa (data de emissao mais antiga e a mais nova) para a Filia do Cliente em questão.

Dim lErro As Long
Dim lComando As Long
Dim dtDataEmissaoPri As Date
Dim dtDataEmissaoUlt As Date
Dim tFilCliFilEmp As typeFilCliFilEmp

On Error GoTo Erro_FilialClienteFilEmp_Pesquisa_Compra

     'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90512
    
    'Inicializa as Datas
    dtDataEmissaoPri = DATA_NULA
    dtDataEmissaoUlt = DATA_NULA
    
    lErro = Comando_Executar(lComando, "SELECT MIN(NFiscal.DataEmissao), MAX(NFiscal.DataEmissao) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturamento = ? AND TiposDocInfo.Complementar = ? AND NFiscal.DataEmissao <> ? AND NFiscal.Cliente =? AND NFiscal.FilialEmpresa = ? AND NFiscal.FilialCli = ? ", dtDataEmissaoPri, dtDataEmissaoUlt, TIPODOCINFO_FATURAVEL, DOCINFO_NORMAL, DATA_NULA, objFilCliFilEmp.lCodCliente, objFilCliFilEmp.iFilialEmpresa, objFilCliFilEmp.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 90513
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90514
    
    If lErro = AD_SQL_SUCESSO Then
        objFilCliFilEmp.dtDataPrimeiraCompra = dtDataEmissaoPri
        objFilCliFilEmp.dtDataUltimaCompra = dtDataEmissaoUlt
    End If
    
    Call Comando_Fechar(lComando)
       
    FilialClienteFilEmp_Pesquisa_Compra = SUCESSO
    
    Exit Function
    
Erro_FilialClienteFilEmp_Pesquisa_Compra:

    FilialClienteFilEmp_Pesquisa_Compra = gErr
    
    Select Case gErr
        
        Case 90512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90513, 90514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152712)
        
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
       
    Exit Function

End Function
'Maristela(Fim)

Function FilialCliente_Le_Estendida(ByVal objFilialCliente As ClassFilialCliente, ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le todos os Dados da Filial Cliente inclusive os dados de Estatística

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Le_Estendida

    'Faz a Leitura na Tabela de Filial cliente
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then Error 52590

    'Se não encontrou  - - -> Erro
    If lErro = 12567 Then Error 52591

    objFilialClienteEstatistica.lCodCliente = objFilialCliente.lCodCliente
    objFilialClienteEstatistica.iFilialEmpresa = giFilialEmpresa
    objFilialClienteEstatistica.iCodFilial = objFilialCliente.iCodFilial

    'Faz a Leitura dos Dados de Estatistica
    lErro = CF("FilialClienteEstatistica_Le", objFilialClienteEstatistica)
    If lErro <> SUCESSO Then Error 52592

    FilialCliente_Le_Estendida = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Estendida:

    FilialCliente_Le_Estendida = Err

    Select Case Err

        Case 52590, 52592 'Tratados nas rotinas chamadas

        Case 52591 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152713)

    End Select

    Exit Function

End Function

Function FilialClienteEstatistica_Le(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le os Dados de Estatística da Filial Cliente

Dim lErro As Long

On Error GoTo Erro_FilialClienteEstatistica_Le
    
    'Le a soma dos Saldos de TitulosReceber
    lErro = FilialCliente_Le_Estatistica_TitRec(objFilialClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52698 Then Error 52699
    
    'Le a soma de NotasFiscais de Venda (faturáveis) que não estão associadas a Fatura
    lErro = FilialCliente_Le_Estatistica_NFsNaoFaturadas(objFilialClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 25655 Then Error 25651
    
    'Le o Saldo de Duplicata da Filial Cliente em questão
    lErro = FilialCliente_Le_Estatistica_Duplicatas(objFilialClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52601 Then Error 52593

    'Le os dados de estatistica com relação ao PV
    lErro = FilialCliente_Le_Estatistica_PV(objFilialClienteEstatistica)
    If lErro <> SUCESSO Then Error 52594

    'Le os dados do cliente Filial com relação à atraso
    lErro = FilialCliente_Le_Estatistica_Atraso(objFilialClienteEstatistica)
    If lErro <> SUCESSO Then Error 52595

    'Le os dados com relação a cheques devolvidos
    lErro = FilialCliente_Le_Estatistica_ChequeDev(objFilialClienteEstatistica)
    If lErro <> SUCESSO Then Error 52596

    'Le os dados com relação a compras
    lErro = FilialCliente_Le_Estatistica_Compras(objFilialClienteEstatistica)
    If lErro <> SUCESSO Then Error 52597

    FilialClienteEstatistica_Le = SUCESSO

    Exit Function

Erro_FilialClienteEstatistica_Le:

    FilialClienteEstatistica_Le = Err

    Select Case Err

        Case 25651, 52593, 52594, 52595, 52596, 52597, 52699 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152714)

    End Select

    Exit Function

End Function

Private Function FilialCliente_Le_Estatistica_TitRec(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le o Saldo dos Titulos da Filial Cliente em questão
'Filtro: lCodCliente , objFilialClienteEstatistica.iFilialEmpresa , iCodFilial

Dim lErro As Long
Dim dSaldoTitulos As Double
Dim lComando As Long

On Error GoTo Erro_FilialCliente_Le_Estatistica_TitRec
            
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52695
               
    'le o saldo dos Titulos em TitulosRec
    lErro = Comando_Executar(lComando, "SELECT SUM(Saldo) FROM TitulosRec WHERE Cliente =? AND FilialEmpresa= ? AND Filial=? ", dSaldoTitulos, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 52696

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52697

    If lErro = AD_SQL_SEM_DADOS Then Error 52698
    
    objFilialClienteEstatistica.dSaldoTitulos = dSaldoTitulos
    
    'Libera comandos
    Call Comando_Fechar(lComando)
        
    FilialCliente_Le_Estatistica_TitRec = SUCESSO
    
    Exit Function
    
Erro_FilialCliente_Le_Estatistica_TitRec:
    
    FilialCliente_Le_Estatistica_TitRec = Err

    Select Case Err
            
        Case 52695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52696, 52697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC3", Err, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iFilialEmpresa)
        
        Case 52698 'tratado na rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152715)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Private Function FilialCliente_Le_Estatistica_NFsNaoFaturadas(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le a soma dos valores das NFs faturáveis não faturadas da FilialCliente em questão
'Filtro: objFilialClienteEstatistica.lCodCliente , objFilialClienteEstatistica.iFilialEmpresa, objClienteEstistica.iCodFilial

Dim lErro As Long
Dim dValorNFsNaoFaturadas As Double
Dim lComando As Long

On Error GoTo Erro_FilialCliente_Le_Estatistica_NFsNaoFaturadas
            
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25652
               
    'Le o saldo das NFs (faturáveis) não faturadas em NFiscal
    lErro = Comando_Executar(lComando, "SELECT SUM(ValorTotal) FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Codigo = ? AND NumIntDocCPR = 0 AND Cliente = ? AND FilialEmpresa = ? AND FilialCli = ?", dValorNFsNaoFaturadas, TIPODOCINFO_FATURAVEL, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 25653

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25654
    If lErro = AD_SQL_SEM_DADOS Then Error 25655
    
    objFilialClienteEstatistica.dValorNFsNaoFaturadas = dValorNFsNaoFaturadas
    
    'Libera comandos
    Call Comando_Fechar(lComando)
        
    FilialCliente_Le_Estatistica_NFsNaoFaturadas = SUCESSO
    
    Exit Function
    
Erro_FilialCliente_Le_Estatistica_NFsNaoFaturadas:
    
    FilialCliente_Le_Estatistica_NFsNaoFaturadas = Err

    Select Case Err
            
        Case 25652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 25653, 25654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)
        
        Case 25655 'tratado na rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152716)

    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function



Private Function FilialCliente_Le_Estatistica_Duplicatas(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le o Saldo de Duplicata da filial Cliente em questão
'Filtro: lCodCliente , objFilialClienteEstatistica.iFilialEmpresa

Dim lErro As Long
Dim dSaldoDuplicatas As Double
Dim lComando As Long

On Error GoTo Erro_FilialCliente_Le_Estatistica_Duplicatas

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52598

    'le o saldo
    lErro = Comando_Executar(lComando, "SELECT SUM(ParcelasRec.Saldo) FROM TitulosRec, ParcelasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND TitulosRec.Cliente =? AND TitulosRec.Filial=? AND TitulosRec.FilialEmpresa= ?  AND ParcelasRec.CarteiraCobranca = ?", dSaldoDuplicatas, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial, objFilialClienteEstatistica.iFilialEmpresa, CARTEIRA_DESCONTADA)
    If lErro <> AD_SQL_SUCESSO Then Error 52599

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52600

    If lErro = AD_SQL_SEM_DADOS Then Error 52601

    objFilialClienteEstatistica.dSaldoDuplicatas = dSaldoDuplicatas

    'Libera comandos
    Call Comando_Fechar(lComando)

    FilialCliente_Le_Estatistica_Duplicatas = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Estatistica_Duplicatas:

    FilialCliente_Le_Estatistica_Duplicatas = Err

    Select Case Err

        Case 52598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52599, 52600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC_TITULOSREC_SALDO", Err, objFilialClienteEstatistica.lCodCliente)

        Case 52601 'tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152717)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function FilialCliente_Le_Estatistica_PV(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Esta Função lê estatísticas com relação aos pedidos de venda

Dim lErro As Long
Dim tItemPedido As typeItemPedido
Dim colItemPedido As New colItemPedido
Dim objItemPedido As New ClassItemPedido
Dim colCodPedido As New Collection
Dim lCodPedBloqueio As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sNumIntItemPV As String
Dim lCodPedido As Long
Dim vCodPedido As Variant

On Error GoTo Erro_FilialCliente_Le_Estatistica_PV

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52602

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 52603

    'Seleciona todos os Pedidos de Venda da FilialCliente e da objFilialClienteEstatistica.iFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT CodPedido,PrecoTotal,Quantidade,QuantCancelada,QuantFaturada , ItensPedidoDeVenda.NumIntDoc FROM  PedidosDeVenda,ItensPedidoDeVenda WHERE PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.Cliente = ? AND PedidosDeVenda.Filial = ? AND PedidosDeVenda.FilialEmpresa = ?", tItemPedido.lCodPedido, tItemPedido.dPrecoTotal, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.lNumIntDoc, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial, objFilialClienteEstatistica.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 52604

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52605

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Key para a coleção
        sNumIntItemPV = CStr(tItemPedido.lNumIntDoc)
        
        'Joga na coleção
        colItemPedido.Add objFilialClienteEstatistica.iFilialEmpresa, tItemPedido.lCodPedido, 0, "", tItemPedido.dQuantidade, 0, tItemPedido.dPrecoTotal, 0, DATA_NULA, "", 0, tItemPedido.dQuantCancelada, 0, Nothing, "", "", "", "", 0, tItemPedido.dQuantFaturada, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, tItemPedido.lNumIntDoc, sNumIntItemPV

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52606

    Loop

    'Seleciona todos os pedidos com bloqueios de crédito
    lErro = Comando_Executar(lComando1, "SELECT Codigo FROM  PedidosDeVenda,BloqueiosPV WHERE PedidosDeVenda.FilialEmpresa = BloqueiosPV.FilialEmpresa AND PedidosDeVenda.Codigo = BloqueiosPV.PedidoDeVenda AND PedidosDeVenda.Cliente = ? AND PedidosDeVenda.Filial = ? AND PedidosDeVenda.FilialEmpresa = ? AND BloqueiosPV.TipoDeBloqueio =? AND BloqueiosPV.DataLib =?", lCodPedBloqueio, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial, objFilialClienteEstatistica.iFilialEmpresa, BLOQUEIO_CREDITO, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then Error 52607

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52608

    Do While lErro <> AD_SQL_SEM_DADOS

        'preenche na coleção
        colCodPedido.Add lCodPedBloqueio

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52609
    Loop

    'Retira da coleção de pedidos os pedidos com bloqueio
    For Each vCodPedido In colCodPedido

        'Codigo do pedido bloqueado por crédito
        lCodPedido = CLng(vCodPedido)
        
        For Each objItemPedido In colItemPedido
            
            If objItemPedido.lCodPedido = lCodPedido Then
            
                sNumIntItemPV = CStr(objItemPedido.lNumIntDoc)
                colItemPedido.Remove (sNumIntItemPV)
                
            End If
            
        Next
    Next

    objFilialClienteEstatistica.dSaldoPedidosLiberados = 0

    'Calcula o saldo dos pedidos liberados
    For Each objItemPedido In colItemPedido
        If (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada) <> 0 Then
            objFilialClienteEstatistica.dSaldoPedidosLiberados = objItemPedido.dPrecoTotal * (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada - objItemPedido.dQuantFaturada) / (objItemPedido.dQuantidade - objItemPedido.dQuantCancelada)
        End If
    Next

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    FilialCliente_Le_Estatistica_PV = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Estatistica_PV:

    FilialCliente_Le_Estatistica_PV = Err

    Select Case Err

        Case 52602, 52603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52604, 52605, 52606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDO_PEDIDOVENDA", Err)

        Case 52607, 52608, 52609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOS_VENDA_BLOQUEIOSPV", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152718)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function FilialCliente_Le_Estatistica_Atraso(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Função que calcula a Estatistica relacionada com atraso da Filial cliente

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim dValPagtosAtraso1 As Double
Dim dValPagtosAtraso2 As Double
Dim dDiasAtrasoValor As Double, dDiasAtrasoValor1 As Double, dDiasAtrasoValor2    As Double
Dim dValPagtosEmDia As Double, dValPagtosEmDia1 As Double, dValPagtosEmDia2 As Double
Dim lMaiorAtraso1 As Long, lMaiorAtraso2 As Long
Dim dSaldoAtrasados As Double
Dim iIndice As Integer, bSuportaTimeStampDiff As Boolean

On Error GoTo Erro_FilialCliente_Le_Estatistica_Atraso

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52610
    Next

    lErro = Conexao_Suporta_TimeStampDiff(bSuportaTimeStampDiff, GL_lConexao)
    If lErro <> SUCESSO Then Error 59293
    
    'Seleciona nas baixas de Títulos Receber
    'Soma(Valores baixados), maior atraso(em dias), soma(dias de atraso * valortes atrasados)
    If bSuportaTimeStampDiff Then
        lErro = Comando_Executar(alComando(0), "SELECT SUM(ValorBaixado), MAX( {fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) } ), SUM( {fn TIMESTAMPDIFF(SQL_TSI_DAY, DataVencimentoReal, Data ) } * ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND TitulosRec.Filial = ? AND ParcelasRec.DataVencimento < BaixasRec.Data", dValPagtosAtraso1, lMaiorAtraso1, dDiasAtrasoValor1, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial)
    Else
        lErro = Comando_Executar(alComando(0), "SELECT SUM(ValorBaixado), MAX( Data - DataVencimentoReal), SUM((Data - DataVencimentoReal)* ValorBaixado) FROM TitulosRec, ParcelasRec,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND TitulosRec.Filial = ? AND ParcelasRec.DataVencimento < BaixasRec.Data", dValPagtosAtraso1, lMaiorAtraso1, dDiasAtrasoValor1, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 52611

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52612
    
    'Seleciona nas baixas de Títulos Receber Baixados
    'Soma(Valores baixados), maior atraso(em dias), soma(dias de atraso * valortes atrasados)
    If bSuportaTimeStampDiff Then
        lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorBaixado), MAX( {fn TIMESTAMPDIFF(SQL_TSI_DAY , DataVencimentoReal, Data) } ), SUM( {fn TIMESTAMPDIFF(SQL_TSI_DAY, DataVencimentoReal, Data ) } * ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRecBaixados.Cliente = ? AND TitulosRecBaixados.Filial = ? AND ParcelasRecBaixadas.DataVencimento < BaixasRec.Data", dValPagtosAtraso2, lMaiorAtraso2, dDiasAtrasoValor2, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial)
    Else
        lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorBaixado), MAX( Data - DataVencimentoReal), SUM((Data - DataVencimentoReal)* ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRecBaixados.Cliente = ? AND TitulosRecBaixados.Filial = ? AND ParcelasRecBaixadas.DataVencimento < BaixasRec.Data", dValPagtosAtraso2, lMaiorAtraso2, dDiasAtrasoValor2, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 52613

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52614

    objFilialClienteEstatistica.dValPagtosAtraso = dValPagtosAtraso1 + dValPagtosAtraso2

    dDiasAtrasoValor = dDiasAtrasoValor1 + dDiasAtrasoValor2

    If lMaiorAtraso1 > lMaiorAtraso2 Then
        objFilialClienteEstatistica.lMaiorAtraso = lMaiorAtraso1
    Else
        objFilialClienteEstatistica.lMaiorAtraso = lMaiorAtraso2
    End If

    'seleciona soma (valores baixados), de Titulos Receber que foram pagos sem atraso
    lErro = Comando_Executar(alComando(2), "SELECT SUM(ValorBaixado) FROM TitulosRec, ParcelasREc,BaixasParcRec,BaixasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND ParcelasRec.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND TitulosRec.Cliente = ? AND TitulosRec.Filial = ? AND DataVencimento >= Data", dValPagtosEmDia1, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 52615

    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52616
    
    'seleciona soma (valores baixados), de Titulos Receber Baixados que foram pagos sem atraso
    lErro = Comando_Executar(alComando(3), "SELECT SUM(ValorBaixado) FROM TitulosRecBaixados, ParcelasRecBaixadas,BaixasParcRec,BaixasRec WHERE TitulosRecBaixados.NumIntDoc = ParcelasRecBaixadas.NumIntTitulo AND ParcelasRecBaixadas.NumIntDoc = BaixasParcRec.NumIntParcela AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND FilialEmpresa =? AND Cliente = ? AND Filial = ? AND DataVencimento >= Data", dValPagtosEmDia1, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 52617

    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52618

    dValPagtosEmDia = dValPagtosEmDia1 + dValPagtosEmDia2

    If dDiasAtrasoValor <> 0 Then
        objFilialClienteEstatistica.lMediaAtraso = dDiasAtrasoValor / (objFilialClienteEstatistica.dValPagtosAtraso + dValPagtosEmDia)
    Else
        objFilialClienteEstatistica.lMediaAtraso = 0
    End If

    'Seleciona o saldo de títulos receber atrasados
    lErro = Comando_Executar(alComando(4), "SELECT SUM(ParcelasRec.Saldo) FROM TitulosRec, ParcelasRec WHERE TitulosRec.NumIntDoc = ParcelasRec.NumIntTitulo AND FilialEmpresa =? AND TitulosRec.Cliente = ?  AND TitulosRec.Filial =? AND ParcelasRec.Saldo > 0  AND ParcelasRec.DataVencimento < ?", dSaldoAtrasados, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial, gdtDataAtual)
    If lErro <> AD_SQL_SUCESSO Then Error 52619

    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52620

    objFilialClienteEstatistica.dSaldoAtrasados = dSaldoAtrasados

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    FilialCliente_Le_Estatistica_Atraso = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Estatistica_Atraso:

    FilialCliente_Le_Estatistica_Atraso = Err

    Select Case Err

        Case 52610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52611, 52612, 52615, 52616, 52619, 52620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC3", Err, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iFilialEmpresa)

        Case 52613, 52614, 52617, 52618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS1", Err, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iFilialEmpresa)

        Case 59293
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152719)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function FilialCliente_Le_Estatistica_ChequeDev(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Função que calcula a estatística com relação a cheque devolvido

Dim lErro As Long
Dim dtDataUltChequeDevolvido As Date
Dim lNumChequesDevolvidos As Long
Dim lComando  As Long
Dim lComando2  As Long

On Error GoTo Erro_FilialCliente_Le_Estatistica_ChequeDev

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52621

    'Inicializa comandos
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 52622

    'Seleciona para BaixasTitRec data do último cheque devolvido e a quantidade de cheque devolvidos
    'Seleciona a partir MovCC_BaixaPR_TitRec
    lErro = Comando_Executar(lComando, "SELECT MAX(DataMovimento), COUNT(*) FROM MovCC_BaixaPR_TitRec WHERE FilialEmpresa =? AND Cliente = ? AND Filial = ? AND Status = ?", dtDataUltChequeDevolvido, lNumChequesDevolvidos, objFilialClienteEstatistica.iFilialEmpresa, objFilialClienteEstatistica.lCodCliente, objFilialClienteEstatistica.iCodFilial, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 52623
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52624
    
    objFilialClienteEstatistica.dtDataUltChequeDevolvido = dtDataUltChequeDevolvido
    objFilialClienteEstatistica.iNumChequesDevolvidos = CInt(lNumChequesDevolvidos)

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    FilialCliente_Le_Estatistica_ChequeDev = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Estatistica_ChequeDev:

    FilialCliente_Le_Estatistica_ChequeDev = Err

    Select Case Err

        Case 52621, 52622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52623, 52624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVCC_BAIXAPR_TITREC", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152720)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function FilialCliente_Le_Estatistica_Compras(ByVal objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Função que calcula a estatística de compras

Dim lErro As Long
Dim objFilCliFilEmp As New ClassFilCliFilEmp

On Error GoTo Erro_FilialCliente_Le_Estatistica_Compras

    objFilCliFilEmp.lCodCliente = objFilialClienteEstatistica.lCodCliente
    objFilCliFilEmp.iFilialEmpresa = objFilialClienteEstatistica.iFilialEmpresa
    objFilCliFilEmp.iCodFilial = objFilialClienteEstatistica.iCodFilial

    'le na tabela de FilialClienteFilEmp
    lErro = CF("FilCliFilEmp_Le", objFilCliFilEmp)
    If lErro <> SUCESSO And lErro <> 52632 Then Error 52627

    objFilialClienteEstatistica.lNumeroCompras = objFilCliFilEmp.lNumeroCompras
    objFilialClienteEstatistica.dMediaCompra = objFilCliFilEmp.dMediaCompra
    objFilialClienteEstatistica.dtDataPrimeiraCompra = objFilCliFilEmp.dtDataPrimeiraCompra
    objFilialClienteEstatistica.dtDataUltimaCompra = objFilCliFilEmp.dtDataUltimaCompra
    objFilialClienteEstatistica.dValorAcumuladoCompras = objFilCliFilEmp.dValorAcumuladoCompras

    FilialCliente_Le_Estatistica_Compras = SUCESSO

    Exit Function

Erro_FilialCliente_Le_Estatistica_Compras:

    FilialCliente_Le_Estatistica_Compras = Err

    Select Case Err

        Case 52627

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152721)

    End Select

    Exit Function

End Function

Function TipoBloqueioPV_Lock(ByVal iTipoDeBloqueioPV As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoBloqueioPV_Lock

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51411
    
    'Busca no BD o Tipo de bloqueio passado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM TiposDeBloqueio WHERE Codigo =? ", iCodigo, iTipoDeBloqueioPV)
    If lErro <> AD_SQL_SUCESSO Then Error 51412
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51413
    If lErro <> AD_SQL_SUCESSO Then Error 51414 'Não encontrou
    
    'Faz lock shared no tipo de bloqueio
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 51415
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TipoBloqueioPV_Lock = SUCESSO
    
    Exit Function
    
Erro_TipoBloqueioPV_Lock:
    
    TipoBloqueioPV_Lock = Err
    
    Select Case Err
    
        Case 51411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 51412, 51413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEBLOQUEIO", Err)
        
        Case 51414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", Err, iTipoDeBloqueioPV)
        
        Case 51415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPODEBLOQUEIO", Err, iTipoDeBloqueioPV)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152722)
    
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoDeVendaBaixado_LeMod(ByVal lComando As Long, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado.

Dim lErro As Long
Dim tPedidoVenda As typePedidoVenda

On Error GoTo Erro_PedidoDeVendaBaixado_LeMod

    tPedidoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tPedidoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tPedidoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tPedidoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    tPedidoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
    tPedidoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
    tPedidoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
    tPedidoVenda.sOBS = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPedidoVenda.sEmitente = String(STRING_USUARIO_CODIGO, 0)
    tPedidoVenda.sUsuarioUltAlteracao = String(STRING_USUARIO_CODIGO, 0)
    
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresaFaturamento, Cliente, Filial, FilialEntrega, CodTransportadora, CondicaoPagto, PercAcrescFinanceiro, DataEmissao, MensagemNota, NaturezaOp, ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, PedidoCliente, CanalVenda, TabelaPreco, ProxSeqBloqueio, FaturaIntegral, ComissaoAutomatica, CobrancaAutomatica, DataReferencia, NumIntDoc, ValorProdutos, CodTranspRedesp, DetPagFrete, DataEntrega, FreteRespons,Moeda,PesoBruto,PesoLiq,Placa,PlacaUF,TaxaMoeda,ValorOutrasDespesas,VolumeEspecie,VolumeMarca,VolumeNumero,VolumeQuant,VolumeTotal,PedidoRepresentante, OBS, Andamento, " & _
        "ValorDescontoTit,ValorDescontoItens,ValorItens, CodigoBase, Parc, Emitente, UsuarioUltAlteracao, DataInclusao, DataAlteracao, HoraInclusao, HoraAlteracao FROM PedidosDeVendaBaixados WHERE FilialEmpresa = ? AND Codigo = ? ", 0, tPedidoVenda.iFilialEmpresaFaturamento, tPedidoVenda.lCliente, tPedidoVenda.iFilial, tPedidoVenda.iFilialEntrega, tPedidoVenda.iCodTransportadora, tPedidoVenda.iCondicaoPagto, _
        tPedidoVenda.dPercAcrescFinanceiro, tPedidoVenda.dtDataEmissao, tPedidoVenda.sMensagemNota, tPedidoVenda.sNaturezaOp, tPedidoVenda.dValorTotal, tPedidoVenda.dValorFrete, tPedidoVenda.dValorDesconto, tPedidoVenda.dValorSeguro, tPedidoVenda.sPedidoCliente, tPedidoVenda.iCanalVenda, tPedidoVenda.iTabelaPreco, tPedidoVenda.iProxSeqBloqueio, tPedidoVenda.iFaturaIntegral, tPedidoVenda.iComissaoAutomatica, tPedidoVenda.iCobrancaAutomatica, tPedidoVenda.dtDataReferencia, tPedidoVenda.lNumIntDoc, tPedidoVenda.dValorProdutos, tPedidoVenda.iCodTranspRedesp, tPedidoVenda.iDetPagFrete, tPedidoVenda.dtDataEntrega, tPedidoVenda.iFreteRespons, tPedidoVenda.iMoeda, tPedidoVenda.dPesoBruto, tPedidoVenda.dPesoLiq, tPedidoVenda.sPlaca, tPedidoVenda.sPlacaUF, tPedidoVenda.dTaxaMoeda, tPedidoVenda.dValorOutrasDespesas, tPedidoVenda.lVolumeEspecie, tPedidoVenda.lVolumeMarca, _
        tPedidoVenda.sVolumeNumero, tPedidoVenda.lVolumeQuant, tPedidoVenda.dVolumeTotal, tPedidoVenda.sPedidoRepresentante, tPedidoVenda.sOBS, tPedidoVenda.iAndamento, tPedidoVenda.dValorDescontoTit, tPedidoVenda.dValorDescontoItens, tPedidoVenda.dValorItens, tPedidoVenda.lCodigoBase, tPedidoVenda.iParc, tPedidoVenda.sEmitente, tPedidoVenda.sUsuarioUltAlteracao, tPedidoVenda.dtDataInclusao, tPedidoVenda.dtDataAlteracao, tPedidoVenda.dHoraInclusao, tPedidoVenda.dHoraAlteracao, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 62203

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62204

    If lErro = AD_SQL_SEM_DADOS Then Error 62205

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 62206

    objPedidoVenda.iFilialEmpresaFaturamento = tPedidoVenda.iFilialEmpresaFaturamento
    objPedidoVenda.lCliente = tPedidoVenda.lCliente
    objPedidoVenda.iFilial = tPedidoVenda.iFilial
    objPedidoVenda.iFilialEntrega = tPedidoVenda.iFilialEntrega
    objPedidoVenda.iCodTransportadora = tPedidoVenda.iCodTransportadora
    objPedidoVenda.iCodTranspRedesp = tPedidoVenda.iCodTranspRedesp
    objPedidoVenda.iDetPagFrete = tPedidoVenda.iDetPagFrete
    objPedidoVenda.dtDataEntrega = tPedidoVenda.dtDataEntrega
    objPedidoVenda.iCondicaoPagto = tPedidoVenda.iCondicaoPagto
    objPedidoVenda.dPercAcrescFinanceiro = tPedidoVenda.dPercAcrescFinanceiro
    objPedidoVenda.dtDataEmissao = tPedidoVenda.dtDataEmissao
    objPedidoVenda.sMensagemNota = tPedidoVenda.sMensagemNota
    objPedidoVenda.sNaturezaOp = tPedidoVenda.sNaturezaOp
    objPedidoVenda.dValorTotal = tPedidoVenda.dValorTotal
    objPedidoVenda.dValorFrete = tPedidoVenda.dValorFrete
    objPedidoVenda.dValorDesconto = tPedidoVenda.dValorDesconto
    objPedidoVenda.dValorSeguro = tPedidoVenda.dValorSeguro
    objPedidoVenda.sPedidoCliente = tPedidoVenda.sPedidoCliente
    objPedidoVenda.iCanalVenda = tPedidoVenda.iCanalVenda
    objPedidoVenda.iTabelaPreco = tPedidoVenda.iTabelaPreco
    objPedidoVenda.iProxSeqBloqueio = tPedidoVenda.iProxSeqBloqueio
    objPedidoVenda.iFaturaIntegral = tPedidoVenda.iFaturaIntegral
    objPedidoVenda.iComissaoAutomatica = tPedidoVenda.iComissaoAutomatica
    objPedidoVenda.iCobrancaAutomatica = tPedidoVenda.iCobrancaAutomatica
    objPedidoVenda.dtDataReferencia = tPedidoVenda.dtDataReferencia
    objPedidoVenda.lNumIntDoc = tPedidoVenda.lNumIntDoc
    objPedidoVenda.dValorProdutos = tPedidoVenda.dValorProdutos
    objPedidoVenda.sPedidoRepresentante = tPedidoVenda.sPedidoRepresentante
    objPedidoVenda.iFreteRespons = tPedidoVenda.iFreteRespons
    objPedidoVenda.iMoeda = tPedidoVenda.iMoeda
    objPedidoVenda.dPesoBruto = tPedidoVenda.dPesoBruto
    objPedidoVenda.dPesoLiq = tPedidoVenda.dPesoLiq
    objPedidoVenda.sPlaca = tPedidoVenda.sPlaca
    objPedidoVenda.sPlacaUF = tPedidoVenda.sPlacaUF
    objPedidoVenda.dTaxaMoeda = tPedidoVenda.dTaxaMoeda
    objPedidoVenda.dValorOutrasDespesas = tPedidoVenda.dValorOutrasDespesas
    objPedidoVenda.lVolumeEspecie = tPedidoVenda.lVolumeEspecie
    objPedidoVenda.lVolumeMarca = tPedidoVenda.lVolumeMarca
    objPedidoVenda.sVolumeNumero = tPedidoVenda.sVolumeNumero
    objPedidoVenda.lVolumeQuant = tPedidoVenda.lVolumeQuant
    objPedidoVenda.dVolumeTotal = tPedidoVenda.dVolumeTotal
    objPedidoVenda.sOBS = tPedidoVenda.sOBS
    objPedidoVenda.iAndamento = tPedidoVenda.iAndamento
    objPedidoVenda.dValorDescontoTit = tPedidoVenda.dValorDescontoTit
    objPedidoVenda.dValorItens = tPedidoVenda.dValorItens
    objPedidoVenda.dValorDescontoItens = tPedidoVenda.dValorDescontoItens
    objPedidoVenda.lCodigoBase = tPedidoVenda.lCodigoBase
    objPedidoVenda.iParc = tPedidoVenda.iParc
    objPedidoVenda.sEmitente = tPedidoVenda.sEmitente
    objPedidoVenda.sUsuarioUltAlteracao = tPedidoVenda.sUsuarioUltAlteracao
    objPedidoVenda.dtDataInclusao = tPedidoVenda.dtDataInclusao
    objPedidoVenda.dtDataAlteracao = tPedidoVenda.dtDataAlteracao
    objPedidoVenda.dHoraInclusao = tPedidoVenda.dHoraInclusao
    objPedidoVenda.dHoraAlteracao = tPedidoVenda.dHoraAlteracao
        
    lErro = CF("PV_Le_Customizado", objPedidoVenda)
    If lErro <> SUCESSO Then gError 62204
    
    PedidoDeVendaBaixado_LeMod = SUCESSO

    Exit Function

Erro_PedidoDeVendaBaixado_LeMod:

    PedidoDeVendaBaixado_LeMod = Err

    Select Case Err

        Case 62203, 62204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedidoVenda.lCodigo)

        Case 62205 'Tratado na rotina chamadora

        Case 62206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152723)

    End Select

    Exit Function

End Function

Function ReservasItemPV_Le_NumIntOrigem(ByVal objItemPV As ClassItemPedido, ByVal colReservaBD As colReservaItem) As Long
'Lê nas tabelas Reserva e Almoxarifado as reservas do item de pedido de venda com o NumintDoc passado

Dim lErro As Long
Dim lComando As Long
Dim tReserva As typeReserva
Dim sAlmoxarifado As String
Dim iTipoOrigem As Integer

On Error GoTo Erro_ReservasItemPV_Le_NumIntOrigem

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51597

    If objItemPV.iPossuiGrade = MARCADO Then
        iTipoOrigem = TIPO_PEDIDO_GRADE
    Else
        iTipoOrigem = TIPO_PEDIDO_COD
    End If
        
    With tReserva
    
        sAlmoxarifado = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sResponsavel = String(STRING_NOME, 0)  'Responsável pela reserva
    
        'Verifica Reserva do Item da própria filial
        lErro = Comando_Executar(lComando, "SELECT Almoxarifado, Quantidade, NomeReduzido, Responsavel, DataValidade FROM Almoxarifado, Reserva WHERE Almoxarifado.Codigo = Reserva.Almoxarifado AND Reserva.Produto = ? AND Reserva.TipoDoc = ? AND NumIntOrigem =?", _
                .iAlmoxarifado, .dQuantidade, sAlmoxarifado, .sResponsavel, .dtDataValidade, objItemPV.sProduto, iTipoOrigem, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 51598
    
        'Tenta ler Reserva da própria filial
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51599
    
        'Coloca na coleção caso não esteja vazio
        Do While lErro <> AD_SQL_SEM_DADOS
    
            colReservaBD.Add .iAlmoxarifado, .dQuantidade, .sResponsavel, sAlmoxarifado, .dtDataValidade, 0
    
            'le o proximo registro da tabela
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51600
    
        Loop
    
    End With
    
    If colReservaBD.Count = 0 Then Error 51601

    'Libera comandos
    Call Comando_Fechar(lComando)

    ReservasItemPV_Le_NumIntOrigem = SUCESSO

    Exit Function

Erro_ReservasItemPV_Le_NumIntOrigem:

    ReservasItemPV_Le_NumIntOrigem = Err

    Select Case Err

        Case 51597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 51598, 51599, 51600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVAITEMBD", Err)

        Case 51601
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152724)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MovEstoque_Le_QuantItemNF(ByVal objNFiscal As ClassNFiscal, dQuantidade As Double, ByVal iTipoMov As Integer, ByVal iItemNF As Integer, ByVal iAlmoxarifado As Integer) As Long
'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque
Dim objItemMovEstoque1 As ClassItemMovEstoque

On Error GoTo Erro_MovEstoque_Le_QuantItemNF
    
    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83342
    
    'Lê NumIntDoc de MovimentoEstoque
    lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.Quantidade FROM TipoDocInfoNFiscal, ItensNFiscal, MovimentoEstoque WHERE TipoDocInfoNFiscal.FilialEmpresa = ? AND TipoDocInfoNFiscal.Serie = ? AND TipoDocInfoNFiscal.NumNotaFiscal = ? AND TipoDocInfoNFiscal.DataEmissao >= ? AND TipoDocInfoNFiscal.DataEmissao <= ? AND TipoDocInfoNFiscal.Status <> ? AND TipoDocInfoNFiscal.Tipo = ? AND TipoDocInfoNFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.Item =? AND MovimentoEstoque.NumIntDocOrigem = ItensNFiscal.NumIntDoc AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND MovimentoEstoque.TipoMov = ? AND MovimentoEstoque.Almoxarifado = ?", _
                                dQuantidade, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, DOCINFO_NF_INT_SAIDA, iItemNF, TIPO_ORIGEM_ITEMNF, iTipoMov, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 83343

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83344

    'se não encontrou o movimento de estoque referente ao item da nf e escaninho
    If lErro = AD_SQL_SEM_DADOS Then gError 83345
    
    'Fecha comandos
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_QuantItemNF = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_QuantItemNF:

    MovEstoque_Le_QuantItemNF = gErr
    
    Select Case gErr
    
        Case 83342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 83343, 83344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL_ITENSNF_MOVESTOQUE", gErr, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_SAIDA, iItemNF, TIPO_ORIGEM_ITEMNF, iTipoMov)
        
        Case 83345
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152725)

    End Select
    
    'Fecha comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function EmbalagensProdItem_Le(ByVal lNumIntDoc As Long, ByVal iTipoDocOrigem As Integer, ByVal colItens As Collection, ByVal colEmbProdItem As Collection) As Long
'Preenche em colEmbProdItem os dados referentes ao pedido de venda ou nota fiscal identificados por lNumIntDoc e iTipoDocOrigem
'colItens serve para se obter o item

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iEmbalagem As Integer, iItem As Integer
Dim sProduto As String, sUMEmbalagem As String
Dim dCapacidade As Double, lNumIntDocItem As Long
Dim dQuantEmb As Double, dQuantProduto As Double
Dim dPesoLiqTotal As Double, dPesoBruto As Double
Dim objEmbProdItem As ClassEmbProdItem, sUMPeso As String

On Error GoTo Erro_EmbalagensProdItem_Le

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 100001
    Next
   
    sUMEmbalagem = String(STRING_UM_SIGLA, 0)
    sUMPeso = String(STRING_UM_SIGLA, 0)
    sProduto = String(STRING_PRODUTO, 0)
    
    If iTipoDocOrigem = DOCORIGEM_PV Then
        
        'obtem embalagens de Pedido de Venda aberto
        lErro = Comando_Executar(alComando(0), "SELECT ItensPedidoDeVenda.NumIntDoc, EmbalagensItem.QuantProduto, EmbalagensItem.QuantEmb, ItensPedidoDeVenda.Produto, EmbalagensItem.Embalagem, ProdutoEmbalagem.UMEmbalagem, ProdutoEmbalagem.UMPeso, ProdutoEmbalagem.Capacidade, ProdutoEmbalagem.PesoLiqTotal,ProdutoEmbalagem.PesoBruto FROM PedidosDeVenda, ItensPedidoDeVenda, EmbalagensItem, ProdutoEmbalagem WHERE PedidosDeVenda.NumIntDoc = ? AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND EmbalagensItem.NumIntDocItem = ItensPedidoDeVenda.NumIntDoc AND EmbalagensItem.TipoDocOrigem = ? AND EmbalagensItem.Embalagem = ProdutoEmbalagem.Embalagem AND ProdutoEmbalagem.Produto = ItensPedidoDeVenda.Produto ORDER BY ItensPedidoDeVenda.NumIntDoc", lNumIntDocItem, dQuantProduto, dQuantEmb, sProduto, iEmbalagem, sUMEmbalagem, sUMPeso, dCapacidade, dPesoLiqTotal, dPesoBruto, lNumIntDoc, iTipoDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 100002

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100003

        Do While lErro <> AD_SQL_SEM_DADOS

            For iIndice = 1 To colItens.Count
       
               If colItens(iIndice) = lNumIntDocItem Then
                    
                   Set objEmbProdItem = New ClassEmbProdItem
    
                    With objEmbProdItem
                        .iTipoDocOrigem = iTipoDocOrigem
                        .iItem = iIndice
                        .sProduto = sProduto
                        .iEmbalagem = iEmbalagem
                        .sUMEmbalagem = sUMEmbalagem
                        .sUMPeso = sUMPeso
                        .dCapacidade = dCapacidade
                        .dQuantEmb = dQuantEmb
                        .dQuantProduto = dQuantProduto
                        .dPesoLiqTotal = dPesoLiqTotal
                        .dPesoBruto = dPesoBruto
                    End With
        
                   colEmbProdItem.Add objEmbProdItem
                
                   Exit For
                   
               End If
            
            Next
            
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100004

        Loop
    
        sUMEmbalagem = String(STRING_UM_SIGLA, 0)
        sUMPeso = String(STRING_UM_SIGLA, 0)
        sProduto = String(STRING_PRODUTO, 0)
    
        'obtem embalagens de Pedido de Venda baixado
        lErro = Comando_Executar(alComando(1), "SELECT ItensPedidoDeVendaBaixados.NumIntDoc, EmbalagensItem.QuantProduto, EmbalagensItem.QuantEmb, ItensPedidoDeVendaBaixados.Produto, EmbalagensItem.Embalagem, ProdutoEmbalagem.UMEmbalagem, ProdutoEmbalagem.UMPeso, ProdutoEmbalagem.Capacidade, ProdutoEmbalagem.PesoLiqTotal,ProdutoEmbalagem.PesoBruto FROM PedidosDeVendaBaixados, ItensPedidoDeVendaBaixados, EmbalagensItem, ProdutoEmbalagem WHERE " & _
            "PedidosDeVendaBaixados.NumIntDoc = ? AND PedidosDeVendaBaixados.Codigo = ItensPedidoDeVendaBaixados.CodPedido AND PedidosDeVendaBaixados.FilialEmpresa = ItensPedidoDeVendaBaixados.FilialEmpresa AND EmbalagensItem.NumIntDocItem = ItensPedidoDeVendaBaixados.NumIntDoc AND EmbalagensItem.TipoDocOrigem = ? AND EmbalagensItem.Embalagem = ProdutoEmbalagem.Embalagem AND ProdutoEmbalagem.Produto = ItensPedidoDeVendaBaixados.Produto ORDER BY ItensPedidoDeVendaBaixados.NumIntDoc", lNumIntDocItem, dQuantProduto, dQuantEmb, sProduto, iEmbalagem, sUMEmbalagem, sUMPeso, dCapacidade, dPesoLiqTotal, dPesoBruto, lNumIntDoc, iTipoDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 100005
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100006
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            For iIndice = 1 To colItens.Count
       
               If colItens(iIndice) = lNumIntDocItem Then

                    Set objEmbProdItem = New ClassEmbProdItem
                    
                    With objEmbProdItem
                        .iTipoDocOrigem = iTipoDocOrigem
                        .iItem = iIndice
                        .sProduto = sProduto
                        .iEmbalagem = iEmbalagem
                        .sUMEmbalagem = sUMEmbalagem
                        .sUMPeso = sUMPeso
                        .dCapacidade = dCapacidade
                        .dQuantEmb = dQuantEmb
                        .dQuantProduto = dQuantProduto
                        .dPesoLiqTotal = dPesoLiqTotal
                        .dPesoBruto = dPesoBruto
                    End With
                    
                    colEmbProdItem.Add objEmbProdItem
                
               End If
               
            Next

            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100007

        Loop
    Else
        '??? Tratar NF
        
        'obtem embalagens de Pedido de Venda aberto
        lErro = Comando_Executar(alComando(0), "SELECT ItensNFiscal.NumIntDoc, EmbalagensItem.QuantProduto, EmbalagensItem.QuantEmb, ItensNFiscal.Produto, EmbalagensItem.Embalagem, ProdutoEmbalagem.UMEmbalagem, ProdutoEmbalagem.UMPeso, ProdutoEmbalagem.Capacidade, ProdutoEmbalagem.PesoLiqTotal,ProdutoEmbalagem.PesoBruto FROM ItensNFiscal, EmbalagensItem, ProdutoEmbalagem WHERE ItensNFiscal.NumIntNF = ? AND EmbalagensItem.NumIntDocItem = ItensNFiscal.NumIntDoc AND EmbalagensItem.TipoDocOrigem = ? AND EmbalagensItem.Embalagem = ProdutoEmbalagem.Embalagem AND ProdutoEmbalagem.Produto = ItensNFiscal.Produto ORDER BY ItensNFiscal.NumIntDoc", lNumIntDocItem, dQuantProduto, dQuantEmb, sProduto, iEmbalagem, sUMEmbalagem, sUMPeso, dCapacidade, dPesoLiqTotal, dPesoBruto, lNumIntDoc, iTipoDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 100002

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100003

        Do While lErro <> AD_SQL_SEM_DADOS

            For iIndice = 1 To colItens.Count
       
               If colItens(iIndice).lNumIntDoc = lNumIntDocItem Then
                    
                   Set objEmbProdItem = New ClassEmbProdItem
    
                    With objEmbProdItem
                        .iTipoDocOrigem = iTipoDocOrigem
                        .iItem = iIndice
                        .sProduto = sProduto
                        .iEmbalagem = iEmbalagem
                        .sUMEmbalagem = sUMEmbalagem
                        .sUMPeso = sUMPeso
                        .dCapacidade = dCapacidade
                        .dQuantEmb = dQuantEmb
                        .dQuantProduto = dQuantProduto
                        .dPesoLiqTotal = dPesoLiqTotal
                        .dPesoBruto = dPesoBruto
                    End With
        
                   colEmbProdItem.Add objEmbProdItem
                
                   Exit For
                   
               End If
            
            Next
            
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 100004

        Loop
    End If

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    EmbalagensProdItem_Le = SUCESSO

    Exit Function

Erro_EmbalagensProdItem_Le:

    EmbalagensProdItem_Le = gErr

    Select Case gErr

        Case 100001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 100002 To 100007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGENS_PV", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152726)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


''Function EmbalagensProdItem_Le(objPedidoVenda As ClassPedidoDeVenda, objEmbProdItem As ClassEmbProdItem) As Long
'''Lê as Embalagens dos itens do Pedido de Venda passada
''
''Dim lErro As Long
''Dim lComando As Long
''Dim objItemPedido As ClassItemPedido
''Dim iTipoDocOrigem As Integer
''
''On Error GoTo Erro_EmbalagensProdItem_Le
''
''    'Abre o Comando
''    lComando = Comando_Abrir()
''    If lComando = 0 Then gError 35706
''
''    For Each objItemPedido In objPedidoVenda.colItensPedido
''
''        'Lê as embalagens do Produto
''        lErro = Comando_Executar(lComando, "SELECT EmbalagensItem.TipoDocOrigem FROM ItensPedidoDeVenda, EmbalagensItem, ProdutoEmbalagem WHERE NumIntDocItem = ? AND ItensPedidoDeVenda.NumIntDoc = EmbalagensItem.NumIntDocItem AND EmbalagensItem.Embalagem = ProdutoEmbalagem.Embalagem AND ItensPedidoDeVenda.Produto = ProdutoEmbalagem.Produto", iTipoDocOrigem, objItemPedido.lNumIntDoc)
''        If lErro <> AD_SQL_SUCESSO Then gError 35710
''
''        lErro = Comando_BuscarPrimeiro(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 35711
''
''        Do While lErro <> AD_SQL_SEM_DADOS
''
''            Set objEmbProdItem = New ClassEmbProdItem
''
''            objEmbProdItem.iTipoDocOrigem = iTipoDocOrigem
''
''            lErro = Comando_BuscarProximo(lComando)
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 35712
''
''        Loop
''
''    Next
''
''    'Fecha o Comando
''    Call Comando_Fechar(lComando)
''
''    EmbalagensProdItem_Le = SUCESSO
''
''    Exit Function
''
''Erro_EmbalagensProdItem_Le:
''
''    EmbalagensProdItem_Le = gErr
''
''    Select Case gErr
''
'''        Case 35706
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'''
'''        Case 35707, 35708
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objItemNF.sProduto)
'''
'''        Case 35709
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemNF.sProduto)
'''
'''        Case 35710, 35711, 35712
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr, objItemNF.lNumIntDoc)
'''
''        Case 83341
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152727)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
'''Janaina
''End Function

Function ProdutoEmbalagem_Le_Padrao(ByVal objProdutoEmbalagem As ClassProdutoEmbalagem) As Long
'Obtem dados da embalagem padrao de um produto
'Retorna 100000 se nao houver embalagem padrao

Dim lErro As Long
Dim lComando As Long
Dim tProdEmb As typeProdutoEmbalagem

On Error GoTo Erro_ProdutoEmbalagem_Le_Padrao

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96112

    tProdEmb.sUMEmbalagem = String(STRING_UM_SIGLA, 0)
    tProdEmb.sUMPeso = String(STRING_UM_SIGLA, 0)
    tProdEmb.sNomeProdEmb = String(STRING_PRODUTOEMBALAGEM_NOMEPRODEMB, 0)
    
    'Lê as embalagens do Produto
    lErro = Comando_Executar(lComando, "SELECT Embalagem,Capacidade,NomeProdEmb,UMEMbalagem,UMPeso,PesoLiqTotal,PesoBruto FROM ProdutoEmbalagem WHERE Produto = ? AND Padrao = ?", tProdEmb.iEmbalagem, tProdEmb.dCapacidade, tProdEmb.sNomeProdEmb, tProdEmb.sUMEmbalagem, tProdEmb.sUMPeso, tProdEmb.dPesoLiqTotal, tProdEmb.dPesoBruto, objProdutoEmbalagem.sProduto, 1)
    If lErro <> AD_SQL_SUCESSO Then gError 96113

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96114

    If lErro <> AD_SQL_SUCESSO Then gError 100000
    
    With tProdEmb
        
        objProdutoEmbalagem.iEmbalagem = .iEmbalagem
        objProdutoEmbalagem.dCapacidade = .dCapacidade
        objProdutoEmbalagem.sUMEmbalagem = .sUMEmbalagem
        objProdutoEmbalagem.dPesoLiqTotal = .dPesoLiqTotal
        objProdutoEmbalagem.dPesoBruto = .dPesoBruto
        objProdutoEmbalagem.sUMPeso = .sUMPeso
        objProdutoEmbalagem.sNomeProdEmb = .sNomeProdEmb
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdutoEmbalagem_Le_Padrao = SUCESSO

    Exit Function

Erro_ProdutoEmbalagem_Le_Padrao:

    ProdutoEmbalagem_Le_Padrao = gErr

    Select Case gErr

        Case 96112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96113, 96114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM", gErr, objProdutoEmbalagem.sProduto)
        
        Case 100000
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152728)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
Function ProdutoEmbalagem_Le_CapacMaxima(ByVal objProdutoEmbalagem As ClassProdutoEmbalagem, ByVal dQuantAlocar As Double, ByVal sUMItem As String, ByVal objProduto As ClassProduto, dFatorEmb As Double) As Long
'Preenche ProdutoEmbalagem com os dados da embalagem do produto com a maior capacidade maxima que seja menor ou igual a dQuantAlocar

Dim lErro As Long
Dim lComando As Long
Dim sUMEmbalagem As String
Dim dCapacidade As Double
Dim iEmbalagem As Integer
Dim dPesoLiqTotal As Double, dFator As Double
Dim dPesoBruto As Double, dCapacConv As Double, dCapacMax As Double

On Error GoTo Erro_ProdutoEmbalagem_Le_CapacMaxima

    dCapacMax = -1
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96106

    sUMEmbalagem = String(STRING_UMEMBALAGEM, 0)
        
    'Lê as embalagens do Produto candidatas
    lErro = Comando_Executar(lComando, "SELECT Embalagem, UMEmbalagem, Capacidade, PesoLiqTotal, PesoBruto FROM ProdutoEmbalagem WHERE Produto = ?", iEmbalagem, sUMEmbalagem, dCapacidade, dPesoLiqTotal, dPesoBruto, objProdutoEmbalagem.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 96107

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96108

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItem, sUMEmbalagem, dFator)
        If lErro <> SUCESSO Then gError 81719
    
        dCapacConv = dCapacidade / dFator
        
        If dCapacConv > dCapacMax And (dCapacConv <= (dQuantAlocar + QTDE_ESTOQUE_DELTA)) Then
        
            objProdutoEmbalagem.iEmbalagem = iEmbalagem
            objProdutoEmbalagem.sUMEmbalagem = sUMEmbalagem
            objProdutoEmbalagem.dCapacidade = dCapacidade
            objProdutoEmbalagem.dPesoLiqTotal = dPesoLiqTotal
            objProdutoEmbalagem.dPesoBruto = dPesoBruto
        
            dCapacMax = dCapacConv
            dFatorEmb = dFator
            
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81720

    Loop
    
    If dCapacMax = -1 Then gError 100008
        
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdutoEmbalagem_Le_CapacMaxima = SUCESSO

    Exit Function

Erro_ProdutoEmbalagem_Le_CapacMaxima:

    ProdutoEmbalagem_Le_CapacMaxima = gErr

    Select Case gErr

        Case 81719
        
        Case 96106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96107, 96108, 81720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM", gErr, objProdutoEmbalagem.sProduto)

        Case 100008 'nao encontrou embalagem
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152729)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoEmbalagem_Le_CapacMinima(ByVal objProdutoEmbalagem As ClassProdutoEmbalagem, ByVal dQuantAlocar As Double, ByVal sUMItem As String, ByVal objProduto As ClassProduto, dFatorEmb As Double) As Long
'Preenche ProdutoEmbalagem com os dados da embalagem do produto com a menor capacidade que seja maior ou igual a dQuantAlocar

Dim lErro As Long
Dim lComando As Long
Dim sUMEmbalagem As String
Dim dCapacidade As Double
Dim iEmbalagem As Integer
Dim dPesoLiqTotal As Double, dFator As Double
Dim dPesoBruto As Double, dCapacConv As Double, dCapacMax As Double

On Error GoTo Erro_ProdutoEmbalagem_Le_CapacMinima

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96109

    sUMEmbalagem = String(STRING_UM_SIGLA, 0)
        
    dCapacMax = -1
    
    lErro = Comando_Executar(lComando, "SELECT Embalagem, UMEmbalagem, Capacidade, PesoLiqTotal, PesoBruto FROM ProdutoEmbalagem WHERE Produto = ? ORDER BY Padrao DESC", iEmbalagem, sUMEmbalagem, dCapacidade, dPesoLiqTotal, dPesoBruto, objProdutoEmbalagem.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 96110

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96111

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItem, sUMEmbalagem, dFator)
        If lErro <> SUCESSO Then gError 81719
    
        dCapacConv = dCapacidade / dFator
        
        If (dCapacConv < dCapacMax Or dCapacMax = -1) And ((dCapacConv + QTDE_ESTOQUE_DELTA) >= dQuantAlocar) Then
        
            objProdutoEmbalagem.iEmbalagem = iEmbalagem
            objProdutoEmbalagem.sUMEmbalagem = sUMEmbalagem
            objProdutoEmbalagem.dCapacidade = dCapacidade
            objProdutoEmbalagem.dPesoLiqTotal = dPesoLiqTotal
            objProdutoEmbalagem.dPesoBruto = dPesoBruto
        
            dCapacMax = dCapacConv
            dFatorEmb = dFator
            
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81720

    Loop
    
    If dCapacMax = -1 Then gError 100009
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdutoEmbalagem_Le_CapacMinima = SUCESSO

    Exit Function

Erro_ProdutoEmbalagem_Le_CapacMinima:

    ProdutoEmbalagem_Le_CapacMinima = gErr

    Select Case gErr

        Case 96109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96110, 96111, 81720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM", gErr, objProdutoEmbalagem.sProduto)

        Case 100009 'nao encontrou embalagem
        
        Case 81719
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152730)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluida por leo em 25/02/02
Public Function BloqueiosPV_Atraso_Testa(ByVal objPedidoVenda As ClassPedidoDeVenda, iAtraso As Integer) As Long
'Verifica se tem bloqueio de atraso

Dim lErro As Long
Dim lComando As Long
Dim iTipoDeBloqueio As Integer

On Error GoTo Erro_BloqueiosPV_Atraso_Testa

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94379

    'Busca Bloqueios de Crédito liberado
    lErro = Comando_Executar(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND TipoDeBloqueio = ? AND DataLib <> ?", iTipoDeBloqueio, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, BLOQUEIO_DIAS_ATRASO, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 94380

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94381

    iAtraso = 0

    'se o bloqueio de credito estiver aprovado ==> retona a flag indicativa
    If lErro = AD_SQL_SUCESSO Then iAtraso = BLOQUEIO_POR_ATRASO_LIBERADO

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    BloqueiosPV_Atraso_Testa = SUCESSO

    Exit Function

Erro_BloqueiosPV_Atraso_Testa:

    BloqueiosPV_Atraso_Testa = gErr

    Select Case gErr

        Case 94379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 94380, 94381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152731)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Incluido por Leo em 22/02/02
Public Function BloqueiosPV_Aprovado_Testa(ByVal objPedidoVenda As ClassPedidoDeVenda, iStatus As Integer, ByVal iTipoBloqueio As Integer) As Long
'Verifica se existe o Bloqueio p/ o Tipo de Bloqueio recebido como parâmetro
'Função genérica

Dim lErro As Long
Dim lComando As Long
Dim iTipoBloqAux As Integer

On Error GoTo Erro_BloqueiosPV_Aprovado_Testa

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94352

    'Busca o tipo de Bloqueio recebido como parâmetro
    lErro = Comando_Executar(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND TipoDeBloqueio = ? AND DataLib <> ?", iTipoBloqAux, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, iTipoBloqueio, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 94353

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94354

    iStatus = 0

    'se o tipo de bloqueio estiver aprovado ==> retona a flag indicativa
    If lErro = AD_SQL_SUCESSO Then iStatus = BLOQUEIOPV_LIBERADO

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    BloqueiosPV_Aprovado_Testa = SUCESSO

    Exit Function

Erro_BloqueiosPV_Aprovado_Testa:

    BloqueiosPV_Aprovado_Testa = gErr

    Select Case gErr

        Case 94352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 94353, 94354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152732)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Rotinas Inseridas por Daniel em 15/05/02

Public Function ItensOV_Le(ByVal objOrcamentoVenda As ClassOrcamentoVenda, Optional ByVal bLeTributacao As Boolean = False) As Long
'Le os itens de OrcamentoVenda - SEM TRIBUTACAO !!!

Dim lErro As Long
Dim lComando As Long
Dim tItensOV As typeItemOV
Dim objItensOV As ClassItemOV

On Error GoTo Erro_ItensOV_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101143

    tItensOV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItensOV.sProduto = String(STRING_PRODUTO, 0)
    tItensOV.sUnidadeMed = String(STRING_UM_SIGLA, 0)
    tItensOV.sVersaoKit = String(STRING_KIT_VERSAO, 0)
    tItensOV.sVersaoKitBase = String(STRING_KIT_VERSAO, 0)
    tItensOV.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0) 'Inserido por Wagner
    
    'Lê os itens do orçamento passado
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Status, MotivoPerda, Observacao, PCSituacao, PCPrecoUnitCalc, PrazoEntrega, Marcado FROM ItensOrcamentoVenda WHERE FilialEmpresa = ? AND CodOrcamento = ? ORDER BY NumIntDoc", tItensOV.lNumIntDoc, tItensOV.sProduto, tItensOV.dQuantidade, tItensOV.dPrecoUnitario, tItensOV.dPrecoTotal, tItensOV.iClasseUM, tItensOV.sUnidadeMed, tItensOV.dValorDesconto, tItensOV.dtDataEntrega, tItensOV.sDescricao, tItensOV.sVersaoKit, tItensOV.sVersaoKitBase, tItensOV.lStatus, tItensOV.lMotivoPerda, tItensOV.sObservacao, tItensOV.iPCSituacao, tItensOV.dPCPrecoUnitCalc, tItensOV.iPrazoEntrega, tItensOV.iMarcado, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase, Status, MotivoPerda, Observacao, PCSituacao, PCPrecoUnitCalc, PrazoEntrega, Marcado FROM ItensOrcamentoVendaHist WHERE FilialEmpresa = ? AND CodOrcamento = ? AND Versao = ? ORDER BY NumIntDoc", tItensOV.lNumIntDoc, tItensOV.sProduto, tItensOV.dQuantidade, tItensOV.dPrecoUnitario, tItensOV.dPrecoTotal, tItensOV.iClasseUM, tItensOV.sUnidadeMed, tItensOV.dValorDesconto, tItensOV.dtDataEntrega, tItensOV.sDescricao, tItensOV.sVersaoKit, tItensOV.sVersaoKitBase, tItensOV.lStatus, tItensOV.lMotivoPerda, tItensOV.sObservacao, tItensOV.iPCSituacao, tItensOV.dPCPrecoUnitCalc, tItensOV.iPrazoEntrega, tItensOV.iMarcado, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iVersao)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 101144

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 101145
    
    If lErro = AD_SQL_SEM_DADOS Then gError 101146

    Do While lErro = AD_SQL_SUCESSO

        With tItensOV
            Set objItensOV = New ClassItemOV

            .iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
            .lCodPedidoOV = objOrcamentoVenda.lCodigo

            objItensOV.iFilialEmpresa = .iFilialEmpresa
            objItensOV.lCodOrcamento = .lCodPedidoOV
            objItensOV.sProduto = .sProduto
            objItensOV.dQuantidade = .dQuantidade
            objItensOV.dPrecoUnitario = .dPrecoUnitario
            objItensOV.dPrecoTotal = .dPrecoTotal
            objItensOV.dValorDesconto = .dValorDesconto
            objItensOV.dtDataEntrega = .dtDataEntrega
            objItensOV.sUnidadeMed = .sUnidadeMed
            objItensOV.iClasseUM = .iClasseUM
            objItensOV.sDescricao = .sDescricao
            objItensOV.sProdutoDescricao = .sDescricao
            objItensOV.iStatus = .iStatus
            objItensOV.lNumIntDoc = .lNumIntDoc
            objItensOV.sVersaoKit = .sVersaoKit
            objItensOV.sVersaoKitBase = .sVersaoKitBase
            objItensOV.lStatus = .lStatus
            objItensOV.lMotivoPerda = .lMotivoPerda
            objItensOV.sObservacao = .sObservacao
            objItensOV.iPCSituacao = .iPCSituacao
            objItensOV.dPCPrecoUnitCalc = .dPCPrecoUnitCalc
            objItensOV.iPrazoEntrega = .iPrazoEntrega
            objItensOV.iMarcado = .iMarcado
            
            'Adiciona o item a colecao de itens do orçamento de Venda
            objOrcamentoVenda.colItens.Add objItensOV
        End With
        
        objItensOV.iHistorico = objOrcamentoVenda.iHistorico
        
        If bLeTributacao Then
        
            If objOrcamentoVenda.iHistorico <> MARCADO Then
                objItensOV.objTributacao.iTipoDoc = TIPODOC_TRIB_OV
            Else
                objItensOV.objTributacao.iTipoDoc = TIPODOC_TRIB_OVHIST
            End If
            
            objItensOV.objTributacao.lNumIntDocItem = objItensOV.lNumIntDoc
            lErro = CF("TributacaoDocItem_Le", objItensOV.objTributacao)
            If lErro <> SUCESSO Then gError 141370
        End If
        
        If objOrcamentoVenda.iHistorico <> MARCADO Then
            objItensOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOV
        Else
            objItensOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOVHIST
        End If
            
        objItensOV.objInfoAdicDocItem.lNumIntDocItem = objItensOV.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Le", objItensOV.objInfoAdicDocItem)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 141370
        
        lErro = ItensOVGrade_Le(objItensOV)
        If lErro <> SUCESSO Then gError 141370
        
        'Busca o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 101147

    Loop

    '#####################################
    'Inserido por Wagner 09/12/05
    lErro = CF("ItensOV_Le_Customizado", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 141370
    '#####################################
    
    'Tratamento feito somente para exibir o antigo campo de desconto global na tela em forma de desconto nos itens
    If objOrcamentoVenda.dValorDesconto > 0 Then
        objOrcamentoVenda.dValorDesconto = 0
        For Each objItensOV In objOrcamentoVenda.colItens
            objItensOV.dValorDesconto = objItensOV.dValorDesconto + objItensOV.objTributacao.dValorDescontoItem
            objItensOV.dPrecoTotal = objItensOV.dPrecoTotal - objItensOV.objTributacao.dValorDescontoItem
            objItensOV.objTributacao.dValorDescontoItem = 0
        Next
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensOV_Le = SUCESSO

    Exit Function

Erro_ItensOV_Le:

    ItensOV_Le = gErr

    Select Case gErr

        Case 101143
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101144, 101145, 101147
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case 101146
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_VENDA_SEM_ITENS", gErr, objOrcamentoVenda.lCodigo)
            
        Case 141370 'Inserido por Wagner
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152733)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItensOV_Le_Lock(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Faz lock exclusive nos itens de OrcamentoVenda e lê os NumIntDoc desses itens

Dim lErro As Long
Dim lComando As Long
Dim tItensOV As typeItemOV
Dim objItensOV As ClassItemOV

On Error GoTo Erro_ItensOV_Le_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101148

    tItensOV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItensOV.sProduto = String(STRING_PRODUTO, 0)
    tItensOV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    '###########################
    'INSERIDO POR WAGNER
    tItensOV.sVersaoKit = String(STRING_KIT_VERSAO, 0)
    tItensOV.sVersaoKitBase = String(STRING_KIT_VERSAO, 0)
    '##########################

    '###########################
    'ALTERADO POR WAGNER
    'Lê os itens do orçamento passado
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao, VersaoKit, VersaoKitBase FROM ItensOrcamentoVenda WHERE FilialEmpresa = ? AND CodOrcamento = ?", 0, tItensOV.lNumIntDoc, tItensOV.sProduto, tItensOV.dQuantidade, tItensOV.dPrecoUnitario, tItensOV.dPrecoTotal, tItensOV.iClasseUM, tItensOV.sUnidadeMed, tItensOV.dValorDesconto, tItensOV.dtDataEntrega, tItensOV.sDescricao, tItensOV.sVersaoKit, tItensOV.sVersaoKitBase, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101149
    '###########################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 101150
    If lErro = AD_SQL_SEM_DADOS Then gError 101151

    Do While lErro = AD_SQL_SUCESSO
        'faz o lock do item
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 101152

        With tItensOV
            Set objItensOV = New ClassItemOV

            .iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
            .lCodPedidoOV = objOrcamentoVenda.lCodigo

            objItensOV.iFilialEmpresa = .iFilialEmpresa
            objItensOV.lCodOrcamento = .lCodPedidoOV
            objItensOV.sProduto = .sProduto
            objItensOV.dQuantidade = .dQuantidade
            objItensOV.dPrecoUnitario = .dPrecoUnitario
            objItensOV.dPrecoTotal = .dPrecoTotal
            objItensOV.dValorDesconto = .dValorDesconto
            objItensOV.dtDataEntrega = .dtDataEntrega
            objItensOV.sUnidadeMed = .sUnidadeMed
            objItensOV.iClasseUM = .iClasseUM
            objItensOV.sDescricao = .sDescricao
            objItensOV.iStatus = .iStatus
            objItensOV.lNumIntDoc = .lNumIntDoc

            '###########################
            'INSERIDO POR WAGNER
            objItensOV.sVersaoKit = .sVersaoKit
            objItensOV.sVersaoKitBase = .sVersaoKitBase
            '###########################

            'Adiciona o item a colecao de itens do orçamento de Venda
            objOrcamentoVenda.colItens.Add objItensOV
        End With

        'Busca o próximo item
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 101153

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensOV_Le_Lock = SUCESSO

    Exit Function

Erro_ItensOV_Le_Lock:

    ItensOV_Le_Lock = gErr

    Select Case gErr

        Case 101148
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101149, 101150, 101153
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case 101151
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_VENDA_SEM_ITENS", gErr, objOrcamentoVenda.lCodigo)

        Case 101152
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSOV", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152734)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Public Function OrcamentoDeVenda_Le_Complementos(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'
'Dim lComando As Long, objTributacaoOV As ClassTributacaoOV
'Dim tTributacaoComplOV As typeTributacaoComplPV, objTributacaoItemOV As ClassTributacaoItemPV
'Dim lErro As Long
'
'On Error GoTo Erro_OrcamentoDeVenda_Le_Complementos
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 101154
'
'    Set objTributacaoOV = objOrcamentoVenda.objTributacaoOV
'
'    tTributacaoComplOV.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'    'ler dados de frete, seguro, despesas, descontos da tabela TributacaoComplOV
'    lErro = Comando_Executar(lComando, "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, " & _
'        "IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, " & _
'        "ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoComplOV WHERE FilialEmpresa = ? AND CodOrcamento = ?", _
'        tTributacaoComplOV.iItem, tTributacaoComplOV.sNaturezaOp, tTributacaoComplOV.iNaturezaOpManual, tTributacaoComplOV.iTipoTributacao, tTributacaoComplOV.iTipoTributacaoManual, tTributacaoComplOV.iIPITipo, tTributacaoComplOV.iIPITipoManual, tTributacaoComplOV.dIPIBaseCalculo, tTributacaoComplOV.iIPIBaseManual, tTributacaoComplOV.dIPIPercRedBase, _
'        tTributacaoComplOV.iIPIPercRedBaseManual, tTributacaoComplOV.dIPIAliquota, tTributacaoComplOV.iIPIAliquotaManual, tTributacaoComplOV.dIPIValor, tTributacaoComplOV.iIPIValorManual, tTributacaoComplOV.iICMSTipo, tTributacaoComplOV.iICMSTipoManual, tTributacaoComplOV.dICMSBase, tTributacaoComplOV.iICMSBaseManual, tTributacaoComplOV.dICMSPercRedBase, _
'        tTributacaoComplOV.iICMSPercRedBaseManual, tTributacaoComplOV.dICMSAliquota, tTributacaoComplOV.iICMSAliquotaManual, tTributacaoComplOV.dICMSValor, tTributacaoComplOV.iICMSValorManual, tTributacaoComplOV.dICMSSubstBase, tTributacaoComplOV.iICMSSubstBaseManual, _
'        tTributacaoComplOV.dICMSSubstAliquota, tTributacaoComplOV.iICMSSubstAliquotaManual, tTributacaoComplOV.dICMSSubstValor, tTributacaoComplOV.iICMSSubstValorManual, tTributacaoComplOV.dPISCredito, tTributacaoComplOV.dCOFINSCredito, tTributacaoComplOV.dIPICredito, tTributacaoComplOV.dICMSCredito, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 101155
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101156
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 101272
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemOV = New ClassTributacaoItemPV
'
'        With objTributacaoItemOV
'
'            .lNumIntDoc = tTributacaoComplOV.iItem
'            .sNaturezaOp = tTributacaoComplOV.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoComplOV.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoComplOV.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoComplOV.iTipoTributacaoManual
'            .iIPITipo = tTributacaoComplOV.iIPITipo
'            .iIPITipoManual = tTributacaoComplOV.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoComplOV.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoComplOV.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoComplOV.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoComplOV.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoComplOV.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoComplOV.iIPIAliquotaManual
'            .dIPIValor = tTributacaoComplOV.dIPIValor
'            .iIPIValorManual = tTributacaoComplOV.iIPIValorManual
'            .iICMSTipo = tTributacaoComplOV.iICMSTipo
'            .iICMSTipoManual = tTributacaoComplOV.iICMSTipoManual
'            .dICMSBase = tTributacaoComplOV.dICMSBase
'            .iICMSBaseManual = tTributacaoComplOV.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoComplOV.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoComplOV.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoComplOV.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoComplOV.iICMSAliquotaManual
'            .dICMSValor = tTributacaoComplOV.dICMSValor
'            .iICMSValorManual = tTributacaoComplOV.iICMSValorManual
'            .dICMSSubstBase = tTributacaoComplOV.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoComplOV.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoComplOV.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoComplOV.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoComplOV.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoComplOV.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoComplOV.dICMSCredito
'            .dIPICredito = tTributacaoComplOV.dIPICredito
'            .dCOFINSCredito = tTributacaoComplOV.dCOFINSCredito
'            .dPISCredito = tTributacaoComplOV.dPISCredito
'
'        End With
'
'        Select Case tTributacaoComplOV.iItem
'
'            Case ITEM_TIPO_FRETE
'                Set objTributacaoOV.objTributacaoFrete = objTributacaoItemOV
'
'            Case ITEM_TIPO_SEGURO
'                Set objTributacaoOV.objTributacaoSeguro = objTributacaoItemOV
'
'            Case ITEM_TIPO_DESCONTO
'                Set objTributacaoOV.objTributacaoDesconto = objTributacaoItemOV
'
'            Case ITEM_TIPO_OUTRAS_DESP
'                Set objTributacaoOV.objTributacaoOutras = objTributacaoItemOV
'
'            Case Else 'item invalido
'                gError 101157
'
'        End Select
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101158
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    OrcamentoDeVenda_Le_Complementos = SUCESSO
'
'    Exit Function
'
'Erro_OrcamentoDeVenda_Le_Complementos:
'
'    OrcamentoDeVenda_Le_Complementos = gErr
'
'    Select Case gErr
'
'        Case 101154
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 101155, 101156, 101158
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOSDEVENDA", gErr, objOrcamentoVenda.lCodigo)
'
'        Case 101157
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPL_OV_TIPO_INVALIDO", gErr)
'
'        Case 101272
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152735)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Public Function OrcamentoVenda_Le_Tributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Le a Tributacao do Orcamento de Venda Passado como parametro.

Dim lErro As Long
'Dim lComando As Long
'Dim tTributacaoOV As typeTributacaoOV
'Dim objTributacaoOV As New ClassTributacaoOV
'Dim objTribComplOV As ClassTributacaoItemPV

On Error GoTo Erro_OrcamentoVenda_Le_Tributacao

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 101159
'
'    With tTributacaoOV
'        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, CodOrcamento, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoOV WHERE FilialEmpresa = ? AND CodOrcamento = ?", _
'            .iFilialEmpresa, .lCodOrcamento, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual, _
'            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 101160
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101161
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 101162
'
'    objOrcamentoVenda.iNaturezaOpManual = VAR_PREENCH_MANUAL
'    objTributacaoOV.lCodOrcamento = objOrcamentoVenda.lCodigo
'    objTributacaoOV.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
'
'    With objTributacaoOV
'        .dICMSBase = tTributacaoOV.dICMSBase
'        .dICMSSubstBase = tTributacaoOV.dICMSSubstBase
'        .dICMSSubstValor = tTributacaoOV.dICMSSubstValor
'        .dICMSValor = tTributacaoOV.dICMSValor
'        .dIPIBase = tTributacaoOV.dIPIBase
'        .dIPIValor = tTributacaoOV.dIPIValor
'        .dIRRFAliquota = tTributacaoOV.dIRRFAliquota
'        .dIRRFBase = tTributacaoOV.dIRRFBase
'        .dIRRFValor = tTributacaoOV.dIRRFValor
'        .dISSAliquota = tTributacaoOV.dISSAliquota
'        .dISSBase = tTributacaoOV.dISSBase
'        .dISSValor = tTributacaoOV.dISSValor
'        .iFilialEmpresa = tTributacaoOV.iFilialEmpresa
'        .iICMSBaseManual = tTributacaoOV.iICMSBaseManual
'        .iICMSSubstBaseManual = tTributacaoOV.iICMSSubstBaseManual
'        .iICMSSubstValorManual = tTributacaoOV.iICMSSubstValorManual
'        .iICMSValorManual = tTributacaoOV.iICMSValorManual
'        .iIPIBaseManual = tTributacaoOV.iIPIBaseManual
'        .iIPIValorManual = tTributacaoOV.iIPIValorManual
'        .iIRRFAliquotaManual = tTributacaoOV.iIRRFAliquotaManual
'        .iIRRFValorManual = tTributacaoOV.iIRRFValorManual
'        .iISSIncluso = tTributacaoOV.iISSIncluso
'        .iISSInclusoManual = tTributacaoOV.iISSValorManual
'        .iISSValorManual = tTributacaoOV.iISSValorManual
'        .iTaxacaoAutomatica = tTributacaoOV.iTaxacaoAutomatica
'        .iTipoTributacao = tTributacaoOV.iTipoTributacao
'        .iTipoTributacaoManual = tTributacaoOV.iTipoTributacaoManual
'        .lCodOrcamento = tTributacaoOV.lCodOrcamento
'        .iPISRetidoManual = tTributacaoOV.iPISRetidoManual
'        .iISSRetidoManual = tTributacaoOV.iISSRetidoManual
'        .iCOFINSRetidoManual = tTributacaoOV.iCOFINSRetidoManual
'        .iCSLLRetidoManual = tTributacaoOV.iCSLLRetidoManual
'        .dPISRetido = tTributacaoOV.dPISRetido
'        .dISSRetido = tTributacaoOV.dISSRetido
'        .dCOFINSRetido = tTributacaoOV.dCOFINSRetido
'        .dCSLLRetido = tTributacaoOV.dCSLLRetido
'
'        .dPISCredito = tTributacaoOV.dPISCredito
'        .iPISCreditoManual = tTributacaoOV.iPISCreditoManual
'        .dCOFINSCredito = tTributacaoOV.dCOFINSCredito
'        .iCOFINSCreditoManual = tTributacaoOV.iCOFINSCreditoManual
'        .dIPICredito = tTributacaoOV.dIPICredito
'        .iIPICreditoManual = tTributacaoOV.iIPICreditoManual
'        .dICMSCredito = tTributacaoOV.dICMSCredito
'        .iICMSCreditoManual = tTributacaoOV.iICMSCreditoManual
'        .dValorINSS = tTributacaoOV.dValorINSS
'        .iINSSValorManual = tTributacaoOV.iINSSValorManual
'        .iINSSRetido = tTributacaoOV.iINSSRetido
'        .iINSSRetidoManual = tTributacaoOV.iINSSRetidoManual
'        .dINSSBase = tTributacaoOV.dINSSBase
'        .iINSSBaseManual = tTributacaoOV.iINSSBaseManual
'        .dINSSDeducoes = tTributacaoOV.dINSSDeducoes
'        .iINSSDeducoesManual = tTributacaoOV.iINSSDeducoesManual
'
'    End With
'
'    Set objOrcamentoVenda.objTributacaoOV = objTributacaoOV
'
'    lErro = CF("OrcamentoDeVenda_Le_Complementos", objOrcamentoVenda)
'    If lErro <> SUCESSO And lErro <> 101272 Then gError 101163
'    If lErro = 101272 Then gError 101273
'
'    Call Comando_Fechar(lComando)

    If objOrcamentoVenda.iHistorico <> MARCADO Then
        objOrcamentoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_OV
    Else
        objOrcamentoVenda.objTributacao.iTipoDoc = TIPODOC_TRIB_OVHIST
    End If
    objOrcamentoVenda.objTributacao.lNumIntDoc = objOrcamentoVenda.lNumIntDoc

    lErro = CF("TributacaoDoc_Le", objOrcamentoVenda.objTributacao)
    If lErro <> SUCESSO Then gError 101163
    
    OrcamentoVenda_Le_Tributacao = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Le_Tributacao:

    OrcamentoVenda_Le_Tributacao = gErr

    Select Case gErr
'
'        Case 101159
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 101160, 101161
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBOV", gErr)
'
'        Case 101162, 84420 To 84423, 101163
'
'        Case 101273
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_OV_NAO_CADASTRADO", gErr)

        Case 101163

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152736)

    End Select

'    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcamentoVenda_Le(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do Código e Filial Empresa passados.

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim iPrazoValidade As Integer

On Error GoTo Erro_OrcamentoVenda_Le
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101229

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 101229

    tOrcamentoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tOrcamentoVenda.sNomeCli = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tOrcamentoVenda.sNomeFilialCli = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tOrcamentoVenda.sUsuario = String(STRING_USUARIO, 0)
    tOrcamentoVenda.sContato = String(STRING_CONTATO, 0)
    tOrcamentoVenda.sEmail = String(STRING_EMAIL, 0)
    tOrcamentoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tOrcamentoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tOrcamentoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    tOrcamentoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
    tOrcamentoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
    tOrcamentoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
    tOrcamentoVenda.sPrazoTexto = String(STRING_MAXIMO, 0)
      
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp, NumIntPedVenda, NumIntNFiscal, NomeCli, NomeFilialCli, Vendedor, Vendedor2, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao,  ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, TabelaPreco, DataReferencia, NumIntDoc,ValorOutrasDespesas,ValorProdutos, PrazoValidade, Usuario, CodigoBase, Status, MotivoPerda, StatusComercial, Versao, DataUltAlt, HoraUltAlt,ValorDescontoTit,ValorDescontoItens,ValorItens, Contato, Email, FilialEntrega,CodTransportadora,MensagemNota,PedidoCliente,CanalVenda,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,VolumeNumero,PesoLiq,PesoBruto,FreteRespons,CodTranspRedesp,DetPagFrete,DataEntrega,VolumeTotal,PedidoRepresentante,PrazoEntrega,DataEnvio,DataEnt, Moeda, DataPerda, PrazoTexto, Cotacao, DataPrevReceb, DataProxCobr, Idioma,PercParticVend2 FROM OrcamentoVenda WHERE FilialEmpresa = ? AND Codigo = ? ", _
        tOrcamentoVenda.sNaturezaOp, tOrcamentoVenda.lNumIntPedVenda, tOrcamentoVenda.lNumIntNFiscal, tOrcamentoVenda.sNomeCli, tOrcamentoVenda.sNomeFilialCli, tOrcamentoVenda.iVendedor, tOrcamentoVenda.iVendedor2, tOrcamentoVenda.lCliente, tOrcamentoVenda.iFilial, tOrcamentoVenda.iCondicaoPagto, tOrcamentoVenda.dPercAcrescFinanceiro, tOrcamentoVenda.dtDataEmissao, tOrcamentoVenda.dValorTotal, tOrcamentoVenda.dValorFrete, tOrcamentoVenda.dValorDesconto, tOrcamentoVenda.dValorSeguro, tOrcamentoVenda.iTabelaPreco, tOrcamentoVenda.dtDataReferencia, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorOutrasDespesas, tOrcamentoVenda.dValorProdutos, iPrazoValidade, tOrcamentoVenda.sUsuario, tOrcamentoVenda.lCodigoBase, tOrcamentoVenda.lStatus, tOrcamentoVenda.lMotivoPerda, tOrcamentoVenda.iStatusComercial, tOrcamentoVenda.iVersao, tOrcamentoVenda.dtDataUltAlt, tOrcamentoVenda.dHoraUltAlt, _
        tOrcamentoVenda.dValorDescontoTit, tOrcamentoVenda.dValorDescontoItens, tOrcamentoVenda.dValorItens, tOrcamentoVenda.sContato, tOrcamentoVenda.sEmail, _
        tOrcamentoVenda.iFilialEntrega, tOrcamentoVenda.iCodTransportadora, tOrcamentoVenda.sMensagemNota, tOrcamentoVenda.sPedidoCliente, tOrcamentoVenda.iCanalVenda, tOrcamentoVenda.sPlaca, tOrcamentoVenda.sPlacaUF, tOrcamentoVenda.lVolumeQuant, tOrcamentoVenda.lVolumeEspecie, tOrcamentoVenda.lVolumeMarca, tOrcamentoVenda.sVolumeNumero, tOrcamentoVenda.dPesoLiq, tOrcamentoVenda.dPesoBruto, tOrcamentoVenda.iFreteRespons, tOrcamentoVenda.iCodTranspRedesp, tOrcamentoVenda.iDetPagFrete, tOrcamentoVenda.dtDataEntrega, tOrcamentoVenda.dVolumeTotal, tOrcamentoVenda.sPedidoRepresentante, tOrcamentoVenda.iPrazoEntrega, tOrcamentoVenda.dtDataEnvio, tOrcamentoVenda.iDataEnt, tOrcamentoVenda.iMoeda, tOrcamentoVenda.dtDataPerda, tOrcamentoVenda.sPrazoTexto, tOrcamentoVenda.dCotacao, tOrcamentoVenda.dtDataPrevReceb, tOrcamentoVenda.dtDataProxCobr, tOrcamentoVenda.iIdioma, tOrcamentoVenda.dPercParticVend2, _
        objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101230

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101231
    
    objOrcamentoVenda.iHistorico = DESMARCADO
    
    'Se tiver passado a versão e não for a versão ativa lê o histórico
    If objOrcamentoVenda.iVersao <> 0 And objOrcamentoVenda.iVersao <> tOrcamentoVenda.iVersao Then
    
        tOrcamentoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        tOrcamentoVenda.sNomeCli = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        tOrcamentoVenda.sNomeFilialCli = String(STRING_FILIAL_CLIENTE_NOME, 0)
        tOrcamentoVenda.sUsuario = String(STRING_USUARIO, 0)
        tOrcamentoVenda.sContato = String(STRING_CONTATO, 0)
        tOrcamentoVenda.sEmail = String(STRING_EMAIL, 0)
        tOrcamentoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
        tOrcamentoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
        tOrcamentoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
        tOrcamentoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
        tOrcamentoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
        tOrcamentoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
        tOrcamentoVenda.sPrazoTexto = String(STRING_MAXIMO, 0)
    
        lErro = Comando_Executar(lComando2, "SELECT NaturezaOp, NumIntPedVenda, NumIntNFiscal, NomeCli, NomeFilialCli, Vendedor, Vendedor2, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao,  ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, TabelaPreco, DataReferencia, NumIntDoc,ValorOutrasDespesas,ValorProdutos, PrazoValidade, Usuario, CodigoBase, Status, MotivoPerda, StatusComercial, Versao, DataUltAlt, HoraUltAlt,ValorDescontoTit,ValorDescontoItens,ValorItens, Contato, Email, FilialEntrega,CodTransportadora,MensagemNota,PedidoCliente,CanalVenda,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,VolumeNumero,PesoLiq,PesoBruto,FreteRespons,CodTranspRedesp,DetPagFrete,DataEntrega,VolumeTotal,PedidoRepresentante,PrazoEntrega,DataEnvio,DataEnt, Moeda, DataPerda, PrazoTexto, Cotacao, DataPrevReceb, DataProxCobr, Idioma, PercParticVend2 FROM OrcamentoVendaHist WHERE FilialEmpresa = ? AND Codigo = ? AND Versao = ? ", _
            tOrcamentoVenda.sNaturezaOp, tOrcamentoVenda.lNumIntPedVenda, tOrcamentoVenda.lNumIntNFiscal, tOrcamentoVenda.sNomeCli, tOrcamentoVenda.sNomeFilialCli, tOrcamentoVenda.iVendedor, tOrcamentoVenda.iVendedor2, tOrcamentoVenda.lCliente, tOrcamentoVenda.iFilial, tOrcamentoVenda.iCondicaoPagto, tOrcamentoVenda.dPercAcrescFinanceiro, tOrcamentoVenda.dtDataEmissao, tOrcamentoVenda.dValorTotal, tOrcamentoVenda.dValorFrete, tOrcamentoVenda.dValorDesconto, tOrcamentoVenda.dValorSeguro, tOrcamentoVenda.iTabelaPreco, tOrcamentoVenda.dtDataReferencia, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorOutrasDespesas, tOrcamentoVenda.dValorProdutos, iPrazoValidade, tOrcamentoVenda.sUsuario, tOrcamentoVenda.lCodigoBase, tOrcamentoVenda.lStatus, tOrcamentoVenda.lMotivoPerda, tOrcamentoVenda.iStatusComercial, tOrcamentoVenda.iVersao, tOrcamentoVenda.dtDataUltAlt, tOrcamentoVenda.dHoraUltAlt, _
            tOrcamentoVenda.dValorDescontoTit, tOrcamentoVenda.dValorDescontoItens, tOrcamentoVenda.dValorItens, tOrcamentoVenda.sContato, tOrcamentoVenda.sEmail, _
            tOrcamentoVenda.iFilialEntrega, tOrcamentoVenda.iCodTransportadora, tOrcamentoVenda.sMensagemNota, tOrcamentoVenda.sPedidoCliente, tOrcamentoVenda.iCanalVenda, tOrcamentoVenda.sPlaca, tOrcamentoVenda.sPlacaUF, tOrcamentoVenda.lVolumeQuant, tOrcamentoVenda.lVolumeEspecie, tOrcamentoVenda.lVolumeMarca, tOrcamentoVenda.sVolumeNumero, tOrcamentoVenda.dPesoLiq, tOrcamentoVenda.dPesoBruto, tOrcamentoVenda.iFreteRespons, tOrcamentoVenda.iCodTranspRedesp, tOrcamentoVenda.iDetPagFrete, tOrcamentoVenda.dtDataEntrega, tOrcamentoVenda.dVolumeTotal, tOrcamentoVenda.sPedidoRepresentante, tOrcamentoVenda.iPrazoEntrega, tOrcamentoVenda.dtDataEnvio, tOrcamentoVenda.iDataEnt, tOrcamentoVenda.iMoeda, tOrcamentoVenda.dtDataPerda, tOrcamentoVenda.sPrazoTexto, tOrcamentoVenda.dCotacao, tOrcamentoVenda.dtDataPrevReceb, tOrcamentoVenda.dtDataProxCobr, tOrcamentoVenda.iIdioma, tOrcamentoVenda.dPercParticVend2, _
            objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iVersao)
        If lErro <> AD_SQL_SUCESSO Then gError 101230
    
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101231
        
        If lErro = AD_SQL_SUCESSO Then objOrcamentoVenda.iHistorico = MARCADO
    
    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 101232

    objOrcamentoVenda.lNumIntPedVenda = tOrcamentoVenda.lNumIntPedVenda
    objOrcamentoVenda.lNumIntNFiscal = tOrcamentoVenda.lNumIntNFiscal
    objOrcamentoVenda.sNomeFilialCli = tOrcamentoVenda.sNomeFilialCli
    objOrcamentoVenda.sNomeCli = tOrcamentoVenda.sNomeCli
    objOrcamentoVenda.iPrazoValidade = iPrazoValidade
    objOrcamentoVenda.lCliente = tOrcamentoVenda.lCliente
    objOrcamentoVenda.iFilial = tOrcamentoVenda.iFilial
    objOrcamentoVenda.iCondicaoPagto = tOrcamentoVenda.iCondicaoPagto
    objOrcamentoVenda.dPercAcrescFinanceiro = tOrcamentoVenda.dPercAcrescFinanceiro
    objOrcamentoVenda.dtDataEmissao = tOrcamentoVenda.dtDataEmissao
    objOrcamentoVenda.sNaturezaOp = tOrcamentoVenda.sNaturezaOp
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal
    objOrcamentoVenda.dValorFrete = tOrcamentoVenda.dValorFrete
    objOrcamentoVenda.dValorDesconto = tOrcamentoVenda.dValorDesconto
    objOrcamentoVenda.dValorOutrasDespesas = tOrcamentoVenda.dValorOutrasDespesas
    objOrcamentoVenda.dValorSeguro = tOrcamentoVenda.dValorSeguro
    objOrcamentoVenda.iTabelaPreco = tOrcamentoVenda.iTabelaPreco
    objOrcamentoVenda.dtDataReferencia = tOrcamentoVenda.dtDataReferencia
    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorProdutos = tOrcamentoVenda.dValorProdutos
    objOrcamentoVenda.iVendedor = tOrcamentoVenda.iVendedor
    objOrcamentoVenda.iVendedor2 = tOrcamentoVenda.iVendedor2
    objOrcamentoVenda.sUsuario = tOrcamentoVenda.sUsuario
    objOrcamentoVenda.lCodigoBase = tOrcamentoVenda.lCodigoBase
    objOrcamentoVenda.lStatus = tOrcamentoVenda.lStatus
    objOrcamentoVenda.lMotivoPerda = tOrcamentoVenda.lMotivoPerda
    objOrcamentoVenda.iStatusComercial = tOrcamentoVenda.iStatusComercial
    objOrcamentoVenda.iVersao = tOrcamentoVenda.iVersao
    objOrcamentoVenda.dtDataUltAlt = tOrcamentoVenda.dtDataUltAlt
    objOrcamentoVenda.dHoraUltAlt = tOrcamentoVenda.dHoraUltAlt
    objOrcamentoVenda.dValorDescontoTit = tOrcamentoVenda.dValorDescontoTit
    objOrcamentoVenda.dValorDescontoItens = tOrcamentoVenda.dValorDescontoItens
    objOrcamentoVenda.dValorItens = tOrcamentoVenda.dValorItens
    objOrcamentoVenda.sContato = tOrcamentoVenda.sContato
    objOrcamentoVenda.sEmail = tOrcamentoVenda.sEmail
        
    objOrcamentoVenda.iFilialEntrega = tOrcamentoVenda.iFilialEntrega
    objOrcamentoVenda.iPrazoEntrega = tOrcamentoVenda.iPrazoEntrega
    objOrcamentoVenda.iCodTransportadora = tOrcamentoVenda.iCodTransportadora
    objOrcamentoVenda.iCodTranspRedesp = tOrcamentoVenda.iCodTranspRedesp
    objOrcamentoVenda.sMensagemNota = tOrcamentoVenda.sMensagemNota
    objOrcamentoVenda.sPedidoCliente = tOrcamentoVenda.sPedidoCliente
    objOrcamentoVenda.sPedidoRepresentante = tOrcamentoVenda.sPedidoRepresentante
    objOrcamentoVenda.iCanalVenda = tOrcamentoVenda.iCanalVenda
    objOrcamentoVenda.dPesoBruto = tOrcamentoVenda.dPesoBruto
    objOrcamentoVenda.dPesoLiq = tOrcamentoVenda.dPesoLiq
    objOrcamentoVenda.sPlaca = tOrcamentoVenda.sPlaca
    objOrcamentoVenda.sPlacaUF = tOrcamentoVenda.sPlacaUF
    objOrcamentoVenda.lVolumeQuant = tOrcamentoVenda.lVolumeQuant
    objOrcamentoVenda.lVolumeEspecie = tOrcamentoVenda.lVolumeEspecie
    objOrcamentoVenda.lVolumeMarca = tOrcamentoVenda.lVolumeMarca
    objOrcamentoVenda.sVolumeNumero = tOrcamentoVenda.sVolumeNumero
    objOrcamentoVenda.dVolumeTotal = tOrcamentoVenda.dVolumeTotal
    objOrcamentoVenda.dtDataEntrega = tOrcamentoVenda.dtDataEntrega
    objOrcamentoVenda.iDetPagFrete = tOrcamentoVenda.iDetPagFrete
    objOrcamentoVenda.iFreteRespons = tOrcamentoVenda.iFreteRespons
    objOrcamentoVenda.iDataEnt = tOrcamentoVenda.iDataEnt
    objOrcamentoVenda.dtDataEnvio = tOrcamentoVenda.dtDataEnvio
    
    objOrcamentoVenda.iMoeda = tOrcamentoVenda.iMoeda
    objOrcamentoVenda.dtDataPerda = tOrcamentoVenda.dtDataPerda
    objOrcamentoVenda.sPrazoTexto = tOrcamentoVenda.sPrazoTexto
    objOrcamentoVenda.dCotacao = tOrcamentoVenda.dCotacao
    
    objOrcamentoVenda.dtDataPrevReceb = tOrcamentoVenda.dtDataPrevReceb
    objOrcamentoVenda.dtDataProxCobr = tOrcamentoVenda.dtDataProxCobr
    objOrcamentoVenda.iIdioma = tOrcamentoVenda.iIdioma
    
    objOrcamentoVenda.dPercParticVend2 = tOrcamentoVenda.dPercParticVend2
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    OrcamentoVenda_Le = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Le:

    OrcamentoVenda_Le = gErr

    Select Case gErr

        Case 101232

        Case 101230, 101231
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)
        
        Case 101229
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152737)
    
    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Public Function OrcamentoVenda_Le_Completo(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do Código passado (incluindo tributacao, itens de orcamento, parcelas)

Dim lErro As Long

On Error GoTo Erro_OrcamentoVenda_Le_Completo

    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 101233
    If lErro = 101232 Then gError 101234

    lErro = CF("OrcamentoVenda_Le_Tributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101162 Then gError 101235
    If lErro = 101162 Then gError 101281

    'ler parcelas
    lErro = CF("ParcelasOV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101284 Then gError 101236
'    If lErro = 101284 Then gError 101285

    'ler itens c/tributacao
    lErro = CF("OrcamentoVenda_Le_Itens_ComTributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101278 Then gError 101237
    If lErro = 101278 Then gError 101279
    
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OV
    Else
        objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OVHIST
    End If
    objOrcamentoVenda.objInfoAdic.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objOrcamentoVenda.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 101237

    OrcamentoVenda_Le_Completo = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Le_Completo:

    OrcamentoVenda_Le_Completo = gErr

    Select Case gErr

        Case 101233 To 101237, 101279, 101281, 101285

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152738)

    End Select

    Exit Function

End Function

Public Function OrcamentoVenda_Le_Itens_ComTributacao(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

'Dim lComando As Long, tItemPedido As typeItemPedido, objItemPedido As ClassItemOV
'Dim tTributacaoItemOV As typeTributacaoItemPV, objTributacaoItemOV As ClassTributacaoItemPV
Dim lErro As Long
'
''#####################################
''INSERIDO POR WAGNER
'Dim sVersaoKit As String
'Dim sVersaoKitBase As String
''#####################################
'Dim lStatus As Long
'Dim lMotivoPerda As Long
'Dim sObservacao As String
'Dim iPCSituacao As Integer, dPCPrecoUnitCalc As Double

On Error GoTo Erro_OrcamentoVenda_Le_Itens_ComTributacao
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 101238
'
'    tItemPedido.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
'    tItemPedido.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
'    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
'    tItemPedido.sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
'    tItemPedido.sProdutoNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
'    tItemPedido.sUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA, 0)
'
'    '############################################
'    'INSERIDO POR WAGNER
'    sVersaoKit = String(STRING_KIT_VERSAO, 0)
'    sVersaoKitBase = String(STRING_KIT_VERSAO, 0)
'    '############################################
'
'    sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0) 'Inserido por Wagner
'
'    With tTributacaoItemOV
'
'        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        '###########################################
'        'ALTERADO POR WAGNER
'        lErro = Comando_Executar(lComando, "SELECT ItensOrcamentoVenda.NumIntDoc,ItensOrcamentoVenda.Produto, Quantidade, PrecoUnitario, PrecoTotal, ItensOrcamentoVenda.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensOrcamentoVenda.Descricao, " & _
'            "ItensOrcamentoVenda.Descricao, NomeReduzido, SiglaUMEstoque, ControleEstoque, " & _
'            "NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, TributacaoItemOV.IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, " & _
'            "ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, ItensOrcamentoVenda.VersaoKit, ItensOrcamentoVenda.VersaoKitBase, ItensOrcamentoVenda.Status, ItensOrcamentoVenda.MotivoPerda, ItensOrcamentoVenda.Observacao, ItensOrcamentoVenda.PCSituacao, ItensOrcamentoVenda.PCPrecoUnitCalc " & _
'            "FROM ItensOrcamentoVenda, Produtos, TributacaoItemOV WHERE ItensOrcamentoVenda.FilialEmpresa = ? AND ItensOrcamentoVenda.CodOrcamento = ? AND ItensOrcamentoVenda.Produto = Produtos.Codigo AND ItensOrcamentoVenda.NumIntDoc = TributacaoItemOV.NumIntDoc ORDER BY ItensOrcamentoVenda.NumIntDoc", _
'            tItemPedido.lNumIntDoc, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
'            tItemPedido.sDescricao, tItemPedido.sProdutoDescricao, _
'            tItemPedido.sProdutoNomeReduzido, tItemPedido.sUMEstoque, tItemPedido.iControleEstoque, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, _
'            .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, sVersaoKit, sVersaoKitBase, lStatus, lMotivoPerda, sObservacao, iPCSituacao, dPCPrecoUnitCalc, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 101239
'        '###########################################
'
'    End With
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101240
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 101278
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemOV = New ClassTributacaoItemPV
'
'        With objTributacaoItemOV
'
'            .lNumIntDoc = tTributacaoItemOV.lNumIntDoc
'            .sNaturezaOp = tTributacaoItemOV.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoItemOV.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoItemOV.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoItemOV.iTipoTributacaoManual
'            .iIPITipo = tTributacaoItemOV.iIPITipo
'            .iIPITipoManual = tTributacaoItemOV.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoItemOV.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoItemOV.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoItemOV.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoItemOV.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoItemOV.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoItemOV.iIPIAliquotaManual
'            .dIPIValor = tTributacaoItemOV.dIPIValor
'            .iIPIValorManual = tTributacaoItemOV.iIPIValorManual
'            .iICMSTipo = tTributacaoItemOV.iICMSTipo
'            .iICMSTipoManual = tTributacaoItemOV.iICMSTipoManual
'            .dICMSBase = tTributacaoItemOV.dICMSBase
'            .iICMSBaseManual = tTributacaoItemOV.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoItemOV.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoItemOV.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoItemOV.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoItemOV.iICMSAliquotaManual
'            .dICMSValor = tTributacaoItemOV.dICMSValor
'            .iICMSValorManual = tTributacaoItemOV.iICMSValorManual
'            .dICMSSubstBase = tTributacaoItemOV.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoItemOV.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoItemOV.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoItemOV.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoItemOV.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoItemOV.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoItemOV.dICMSCredito
'            .dIPICredito = tTributacaoItemOV.dIPICredito
'            .dCOFINSCredito = tTributacaoItemOV.dCOFINSCredito
'            .dPISCredito = tTributacaoItemOV.dPISCredito
'
'        End With
'
'        With tItemPedido
'            '#######################################
'            'ALTERADO POR WAGNER
'            Set objItemPedido = objOrcamentoVenda.IncluirItem(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .sUnidadeMed, .iClasseUM, .sDescricao, .iStatus, .lNumIntDoc, objTributacaoItemOV)
'            objItemPedido.sVersaoKit = sVersaoKit
'            objItemPedido.sVersaoKitBase = sVersaoKitBase
'            '#######################################
'        End With
'
'        lErro = ItensOVGrade_Le(objItemPedido)
'        If lErro <> SUCESSO Then gError 86364
'
'        '#####################################
'        'Inserido por Wagner 09/12/05
'        objItemPedido.lStatus = lStatus
'        objItemPedido.lMotivoPerda = lMotivoPerda
'        objItemPedido.sObservacao = sObservacao
'
'        objItemPedido.iPCSituacao = iPCSituacao
'        objItemPedido.dPCPrecoUnitCalc = dPCPrecoUnitCalc
'
'        lErro = CF("ItensOV_Le_Customizado", objOrcamentoVenda)
'        If lErro <> SUCESSO Then gError 141391
'        '#####################################
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101241
'
'    Loop
'
'    '############################################
'    'Inserido por Wagner 22/03/2006
'    lErro = CF("OVItens_Le_Customizado", objOrcamentoVenda)
'    If lErro <> SUCESSO Then gError 177044
'    '############################################
'
'    Call Comando_Fechar(lComando)

    lErro = CF("ItensOV_Le", objOrcamentoVenda, True)
    If lErro <> SUCESSO Then gError 177044
    
    OrcamentoVenda_Le_Itens_ComTributacao = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Le_Itens_ComTributacao:

    OrcamentoVenda_Le_Itens_ComTributacao = gErr

    Select Case gErr

        Case 101238
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 101239, 101240, 101241, 101278, 177044
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152739)

    End Select

'    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcamentoVenda_Lock(ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal lComando As Long) As Long
'Lê e faz o Lock no Orcamento de Venda

Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim lCodigo As Long

On Error GoTo Erro_OrcamentoVenda_Lock

    'Lê o Orcamento com o Código e a filialempresa passados
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, ValorTotal FROM OrcamentoVenda WHERE FilialEmpresa = ? and Codigo = ?", 0, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorTotal, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101242

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101243
    If lErro = AD_SQL_SEM_DADOS Then gError 101244

    'Faz o "Lock" no Orcamento de Venda
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 101245

    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal

    OrcamentoVenda_Lock = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Lock:

    OrcamentoVenda_Lock = gErr

    Select Case gErr

        Case 101242, 101243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)

        Case 101244

        Case 101245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORCAMENTO_DE_VENDA", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152740)

    End Select

    Exit Function

End Function

Public Function OVObjetosVinculados_Lock(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Faz "lock" nos objetos associados ao Orcamento de Venda

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim objProduto As New ClassProduto
Dim objItemOV As ClassItemOV
Dim iIndice As Integer

On Error GoTo Erro_OVObjetosVinculados_Lock

    'Abe os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 101246
    Next

    If objOrcamentoVenda.lCliente <> 0 Then 'Se o Cliente estiver cadastrado
        
        'Loca a FilialCliente
        lErro = CF("FilialCliente_Lock", objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial)
        If lErro <> SUCESSO And lErro <> 43050 Then gError 101247
        If lErro <> SUCESSO Then gError 101248
    
    End If
    
    'Loca a NaturezaOp
    lErro = CF("NaturezaOperacao_Lock", objOrcamentoVenda.sNaturezaOp, alComando(0))
    If lErro <> SUCESSO And lErro <> 40551 Then gError 101249
    If lErro <> SUCESSO Then gError 101250


    If objOrcamentoVenda.iTabelaPreco <> 0 Then
        'Locka a TabelaPreco
        lErro = CF("TabelaDePreco_Lock", objOrcamentoVenda.iTabelaPreco)
        If lErro <> SUCESSO And lErro <> 43128 Then gError 101251
        If lErro <> SUCESSO Then gError 101252
    End If

    If objOrcamentoVenda.iCondicaoPagto <> 0 Then
        'Locka a Condição de Pagamento
        lErro = CF("CondicaoPagto_Lock", alComando(0), objOrcamentoVenda.iCondicaoPagto)
        If lErro <> SUCESSO Then gError 101253
    End If


    'Para cada item de Pedido Venda
    For Each objItemOV In objOrcamentoVenda.colItens

        objProduto.sCodigo = objItemOV.sProduto
        'Locka  o produto
        lErro = CF("Produto_Lock_Shared", alComando(5), objProduto)
        If lErro <> SUCESSO Then gError 101254

        'Armazena os dados do produto para que possam ser usados depois
        objItemOV.iClasseUM = objProduto.iClasseUM

    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OVObjetosVinculados_Lock = SUCESSO

    Exit Function

Erro_OVObjetosVinculados_Lock:

    OVObjetosVinculados_Lock = gErr

    Select Case gErr

        Case 101247, 101249, 101251, 101253, 101254

        Case 101246
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101248
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objOrcamentoVenda.lCliente, objOrcamentoVenda.iFilial)

        Case 101250
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, objOrcamentoVenda.sNaturezaOp)

        Case 101252
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", gErr, objOrcamentoVenda.iTabelaPreco)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152741)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ParcelasOV_Le(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê as parcelas associadas ao objOrcamentoVenda

Dim lErro As Long
Dim lComando As Long
Dim tParcelaOV As typeParcelaOV
Dim objParcelaOV As ClassParcelaOV

On Error GoTo Erro_ParcelasOV_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101263

    'Seleciona Parcelas vinculadas ao Pedido de Venda
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        lErro = Comando_Executar(lComando, "SELECT NumParcela, DataVencimento, Valor, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor FROM ParcelasOV WHERE OrcamentoVenda = ? AND FilialEmpresa = ? ORDER BY DataVencimento", tParcelaOV.iNumParcela, tParcelaOV.dtDataVencimento, tParcelaOV.dValor, tParcelaOV.iDesconto1Codigo, tParcelaOV.dtDesconto1Ate, tParcelaOV.dDesconto1Valor, tParcelaOV.iDesconto2Codigo, tParcelaOV.dtDesconto2Ate, tParcelaOV.dDesconto2Valor, tParcelaOV.iDesconto3Codigo, tParcelaOV.dtDesconto3Ate, tParcelaOV.dDesconto3Valor, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumParcela, DataVencimento, Valor, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor FROM ParcelasOVHist WHERE OrcamentoVenda = ? AND FilialEmpresa = ? AND Versao = ? ORDER BY DataVencimento", tParcelaOV.iNumParcela, tParcelaOV.dtDataVencimento, tParcelaOV.dValor, tParcelaOV.iDesconto1Codigo, tParcelaOV.dtDesconto1Ate, tParcelaOV.dDesconto1Valor, tParcelaOV.iDesconto2Codigo, tParcelaOV.dtDesconto2Ate, tParcelaOV.dDesconto2Valor, tParcelaOV.iDesconto3Codigo, tParcelaOV.dtDesconto3Ate, tParcelaOV.dDesconto3Valor, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.iVersao)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 101264

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101265
    
    If lErro = AD_SQL_SEM_DADOS Then gError 101284

    Do While lErro <> AD_SQL_SEM_DADOS

        With tParcelaOV
            Set objParcelaOV = New ClassParcelaOV

            objParcelaOV.dValor = .dValor
            objParcelaOV.dtDataVencimento = .dtDataVencimento
            objParcelaOV.iNumParcela = .iNumParcela
            objParcelaOV.iDesconto1Codigo = .iDesconto1Codigo
            objParcelaOV.dtDesconto1Ate = .dtDesconto1Ate
            objParcelaOV.dDesconto1Valor = .dDesconto1Valor
            objParcelaOV.iDesconto2Codigo = .iDesconto2Codigo
            objParcelaOV.dtDesconto2Ate = .dtDesconto2Ate
            objParcelaOV.dDesconto2Valor = .dDesconto2Valor
            objParcelaOV.dtDesconto3Ate = .dtDesconto3Ate
            objParcelaOV.dDesconto3Valor = .dDesconto3Valor
            objParcelaOV.iDesconto3Codigo = .iDesconto3Codigo

            objOrcamentoVenda.colParcela.Add objParcelaOV
        End With

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101266

    Loop

    Call Comando_Fechar(lComando)

    ParcelasOV_Le = SUCESSO

    Exit Function

Erro_ParcelasOV_Le:

    ParcelasOV_Le = gErr

    Select Case gErr

        Case 101284
        
        Case 101263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101264, 101265, 101266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASPEDIDODEVENDA", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152742)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcamentoVenda_Le_NumIntDoc(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do NumIntPedVenda passado.

Dim lComando As Long
Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim iPrazoValidade As Integer

On Error GoTo Erro_OrcamentoVenda_Le_NumIntDoc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94459

    tOrcamentoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tOrcamentoVenda.sNomeCli = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tOrcamentoVenda.sNomeFilialCli = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tOrcamentoVenda.sUsuario = String(STRING_USUARIO, 0)
    tOrcamentoVenda.sContato = String(STRING_CONTATO, 0)
    tOrcamentoVenda.sEmail = String(STRING_EMAIL, 0)
    tOrcamentoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tOrcamentoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tOrcamentoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    tOrcamentoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
    tOrcamentoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
    tOrcamentoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
    tOrcamentoVenda.sPrazoTexto = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT NaturezaOp, NumIntPedVenda, NumIntNFiscal, Codigo, FilialEmpresa, NomeCli, NomeFilialCli, Vendedor, Vendedor2, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao,  ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, TabelaPreco, DataReferencia, NumIntDoc,ValorOutrasDespesas,ValorProdutos, PrazoValidade, Usuario,CodigoBase, Status, MotivoPerda, StatusComercial, Versao, DataUltAlt, HoraUltAlt,ValorDescontoTit,ValorDescontoItens,ValorItens, Contato, Email, FilialEntrega,CodTransportadora,MensagemNota,PedidoCliente,CanalVenda,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,VolumeNumero,PesoLiq,PesoBruto,FreteRespons,CodTranspRedesp,DetPagFrete,DataEntrega,VolumeTotal,PedidoRepresentante,PrazoEntrega,DataEnvio,DataEnt, Moeda, DataPerda, PrazoTexto, Cotacao, DataPrevReceb, DataProxCobr, Idioma, PercParticVend2 FROM OrcamentoVenda WHERE NumIntDoc = ? ", _
        tOrcamentoVenda.sNaturezaOp, tOrcamentoVenda.lNumIntPedVenda, tOrcamentoVenda.lNumIntNFiscal, tOrcamentoVenda.lCodigo, tOrcamentoVenda.iFilialEmpresa, tOrcamentoVenda.sNomeCli, tOrcamentoVenda.sNomeFilialCli, tOrcamentoVenda.iVendedor, tOrcamentoVenda.iVendedor2, tOrcamentoVenda.lCliente, tOrcamentoVenda.iFilial, tOrcamentoVenda.iCondicaoPagto, tOrcamentoVenda.dPercAcrescFinanceiro, tOrcamentoVenda.dtDataEmissao, tOrcamentoVenda.dValorTotal, tOrcamentoVenda.dValorFrete, tOrcamentoVenda.dValorDesconto, tOrcamentoVenda.dValorSeguro, tOrcamentoVenda.iTabelaPreco, tOrcamentoVenda.dtDataReferencia, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorOutrasDespesas, tOrcamentoVenda.dValorProdutos, iPrazoValidade, tOrcamentoVenda.sUsuario, tOrcamentoVenda.lCodigoBase, tOrcamentoVenda.lStatus, tOrcamentoVenda.lMotivoPerda, tOrcamentoVenda.iStatusComercial, tOrcamentoVenda.iVersao, tOrcamentoVenda.dtDataUltAlt, tOrcamentoVenda.dHoraUltAlt, _
        tOrcamentoVenda.dValorDescontoTit, tOrcamentoVenda.dValorDescontoItens, tOrcamentoVenda.dValorItens, tOrcamentoVenda.sContato, tOrcamentoVenda.sEmail, _
        tOrcamentoVenda.iFilialEntrega, tOrcamentoVenda.iCodTransportadora, tOrcamentoVenda.sMensagemNota, tOrcamentoVenda.sPedidoCliente, tOrcamentoVenda.iCanalVenda, tOrcamentoVenda.sPlaca, tOrcamentoVenda.sPlacaUF, tOrcamentoVenda.lVolumeQuant, tOrcamentoVenda.lVolumeEspecie, tOrcamentoVenda.lVolumeMarca, tOrcamentoVenda.sVolumeNumero, tOrcamentoVenda.dPesoLiq, tOrcamentoVenda.dPesoBruto, tOrcamentoVenda.iFreteRespons, tOrcamentoVenda.iCodTranspRedesp, tOrcamentoVenda.iDetPagFrete, tOrcamentoVenda.dtDataEntrega, tOrcamentoVenda.dVolumeTotal, tOrcamentoVenda.sPedidoRepresentante, tOrcamentoVenda.iPrazoEntrega, tOrcamentoVenda.dtDataEnvio, tOrcamentoVenda.iDataEnt, tOrcamentoVenda.iMoeda, tOrcamentoVenda.dtDataPerda, tOrcamentoVenda.sPrazoTexto, tOrcamentoVenda.dCotacao, tOrcamentoVenda.dtDataPrevReceb, tOrcamentoVenda.dtDataProxCobr, tOrcamentoVenda.iIdioma, tOrcamentoVenda.dPercParticVend2, _
        objOrcamentoVenda.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 94460

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94461
    
    'Se não encontrou ---> Erro tratado na rotina Chamadora
    If lErro = AD_SQL_SEM_DADOS Then gError 94462
    
    objOrcamentoVenda.lNumIntPedVenda = tOrcamentoVenda.lNumIntPedVenda
    objOrcamentoVenda.lCodigo = tOrcamentoVenda.lCodigo
    objOrcamentoVenda.iFilialEmpresa = tOrcamentoVenda.iFilialEmpresa
    objOrcamentoVenda.lNumIntNFiscal = tOrcamentoVenda.lNumIntNFiscal
    objOrcamentoVenda.sNomeFilialCli = tOrcamentoVenda.sNomeFilialCli
    objOrcamentoVenda.sNomeCli = tOrcamentoVenda.sNomeCli
    objOrcamentoVenda.iPrazoValidade = iPrazoValidade
    objOrcamentoVenda.lCliente = tOrcamentoVenda.lCliente
    objOrcamentoVenda.iFilial = tOrcamentoVenda.iFilial
    objOrcamentoVenda.iCondicaoPagto = tOrcamentoVenda.iCondicaoPagto
    objOrcamentoVenda.dPercAcrescFinanceiro = tOrcamentoVenda.dPercAcrescFinanceiro
    objOrcamentoVenda.dtDataEmissao = tOrcamentoVenda.dtDataEmissao
    objOrcamentoVenda.sNaturezaOp = tOrcamentoVenda.sNaturezaOp
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal
    objOrcamentoVenda.dValorFrete = tOrcamentoVenda.dValorFrete
    objOrcamentoVenda.dValorDesconto = tOrcamentoVenda.dValorDesconto
    objOrcamentoVenda.dValorOutrasDespesas = tOrcamentoVenda.dValorOutrasDespesas
    objOrcamentoVenda.dValorSeguro = tOrcamentoVenda.dValorSeguro
    objOrcamentoVenda.iTabelaPreco = tOrcamentoVenda.iTabelaPreco
    objOrcamentoVenda.dtDataReferencia = tOrcamentoVenda.dtDataReferencia
    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorProdutos = tOrcamentoVenda.dValorProdutos
    objOrcamentoVenda.iVendedor = tOrcamentoVenda.iVendedor
    objOrcamentoVenda.iVendedor2 = tOrcamentoVenda.iVendedor2
    objOrcamentoVenda.sUsuario = tOrcamentoVenda.sUsuario
    objOrcamentoVenda.lCodigoBase = tOrcamentoVenda.lCodigoBase
    objOrcamentoVenda.lStatus = tOrcamentoVenda.lStatus
    objOrcamentoVenda.lMotivoPerda = tOrcamentoVenda.lMotivoPerda
    objOrcamentoVenda.iStatusComercial = tOrcamentoVenda.iStatusComercial
    objOrcamentoVenda.iVersao = tOrcamentoVenda.iVersao
    objOrcamentoVenda.dtDataUltAlt = tOrcamentoVenda.dtDataUltAlt
    objOrcamentoVenda.dHoraUltAlt = tOrcamentoVenda.dHoraUltAlt
    objOrcamentoVenda.dValorDescontoTit = tOrcamentoVenda.dValorDescontoTit
    objOrcamentoVenda.dValorDescontoItens = tOrcamentoVenda.dValorDescontoItens
    objOrcamentoVenda.dValorItens = tOrcamentoVenda.dValorItens
    objOrcamentoVenda.sContato = tOrcamentoVenda.sContato
    objOrcamentoVenda.sEmail = tOrcamentoVenda.sEmail
    
    objOrcamentoVenda.iFilialEntrega = tOrcamentoVenda.iFilialEntrega
    objOrcamentoVenda.iPrazoEntrega = tOrcamentoVenda.iPrazoEntrega
    objOrcamentoVenda.iCodTransportadora = tOrcamentoVenda.iCodTransportadora
    objOrcamentoVenda.iCodTranspRedesp = tOrcamentoVenda.iCodTranspRedesp
    objOrcamentoVenda.sMensagemNota = tOrcamentoVenda.sMensagemNota
    objOrcamentoVenda.sPedidoCliente = tOrcamentoVenda.sPedidoCliente
    objOrcamentoVenda.sPedidoRepresentante = tOrcamentoVenda.sPedidoRepresentante
    objOrcamentoVenda.iCanalVenda = tOrcamentoVenda.iCanalVenda
    objOrcamentoVenda.dPesoBruto = tOrcamentoVenda.dPesoBruto
    objOrcamentoVenda.dPesoLiq = tOrcamentoVenda.dPesoLiq
    objOrcamentoVenda.sPlaca = tOrcamentoVenda.sPlaca
    objOrcamentoVenda.sPlacaUF = tOrcamentoVenda.sPlacaUF
    objOrcamentoVenda.lVolumeQuant = tOrcamentoVenda.lVolumeQuant
    objOrcamentoVenda.lVolumeEspecie = tOrcamentoVenda.lVolumeEspecie
    objOrcamentoVenda.lVolumeMarca = tOrcamentoVenda.lVolumeMarca
    objOrcamentoVenda.sVolumeNumero = tOrcamentoVenda.sVolumeNumero
    objOrcamentoVenda.dVolumeTotal = tOrcamentoVenda.dVolumeTotal
    objOrcamentoVenda.dtDataEntrega = tOrcamentoVenda.dtDataEntrega
    objOrcamentoVenda.iDetPagFrete = tOrcamentoVenda.iDetPagFrete
    objOrcamentoVenda.iFreteRespons = tOrcamentoVenda.iFreteRespons
    objOrcamentoVenda.iDataEnt = tOrcamentoVenda.iDataEnt
    objOrcamentoVenda.dtDataEnvio = tOrcamentoVenda.dtDataEnvio
    
    objOrcamentoVenda.iMoeda = tOrcamentoVenda.iMoeda
    objOrcamentoVenda.dtDataPerda = tOrcamentoVenda.dtDataPerda
    objOrcamentoVenda.sPrazoTexto = tOrcamentoVenda.sPrazoTexto
    objOrcamentoVenda.dCotacao = tOrcamentoVenda.dCotacao
    
    objOrcamentoVenda.dtDataPrevReceb = tOrcamentoVenda.dtDataPrevReceb
    objOrcamentoVenda.dtDataProxCobr = tOrcamentoVenda.dtDataProxCobr
    objOrcamentoVenda.iIdioma = tOrcamentoVenda.iIdioma
    
    objOrcamentoVenda.dPercParticVend2 = tOrcamentoVenda.dPercParticVend2
    
    Call Comando_Fechar(lComando)

    OrcamentoVenda_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Le_NumIntDoc:

    OrcamentoVenda_Le_NumIntDoc = gErr

    Select Case gErr

        Case 94462 'Tratado na rotina chamadora

        Case 94460, 94461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)
        
        Case 94459
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152743)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrcamentoVenda_Le_NumIntPedVenda(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Lê o Orcamento de Venda a partir do NumIntPedVenda passado.

Dim lComando As Long
Dim lErro As Long
Dim tOrcamentoVenda As typeOrcamentoVenda
Dim iPrazoValidade As Integer

On Error GoTo Erro_OrcamentoVenda_Le_NumIntPedVenda

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94453

    tOrcamentoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tOrcamentoVenda.sNomeCli = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tOrcamentoVenda.sNomeFilialCli = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tOrcamentoVenda.sUsuario = String(STRING_USUARIO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NaturezaOp, NumIntNFiscal, Codigo, FilialEmpresa, NomeCli, NomeFilialCli, Vendedor, Vendedor2, Cliente, Filial, CondicaoPagto, PercAcrescFinanceiro, DataEmissao,  ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, TabelaPreco, DataReferencia, NumIntDoc,ValorOutrasDespesas,ValorProdutos, PrazoValidade, Usuario, CodigoBase FROM OrcamentoVenda WHERE NumIntPedVenda = ? ", _
    tOrcamentoVenda.sNaturezaOp, tOrcamentoVenda.lNumIntNFiscal, tOrcamentoVenda.lCodigo, tOrcamentoVenda.iFilialEmpresa, tOrcamentoVenda.sNomeCli, tOrcamentoVenda.sNomeFilialCli, tOrcamentoVenda.iVendedor, tOrcamentoVenda.iVendedor2, tOrcamentoVenda.lCliente, tOrcamentoVenda.iFilial, tOrcamentoVenda.iCondicaoPagto, tOrcamentoVenda.dPercAcrescFinanceiro, tOrcamentoVenda.dtDataEmissao, tOrcamentoVenda.dValorTotal, tOrcamentoVenda.dValorFrete, tOrcamentoVenda.dValorDesconto, tOrcamentoVenda.dValorSeguro, tOrcamentoVenda.iTabelaPreco, tOrcamentoVenda.dtDataReferencia, tOrcamentoVenda.lNumIntDoc, tOrcamentoVenda.dValorOutrasDespesas, tOrcamentoVenda.dValorProdutos, iPrazoValidade, tOrcamentoVenda.sUsuario, tOrcamentoVenda.lCodigoBase, objOrcamentoVenda.lNumIntPedVenda)
    If lErro <> AD_SQL_SUCESSO Then gError 94454

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94455
    
    'Se não encontrou ---> Erro tratado na rotina Chamadora
    If lErro = AD_SQL_SEM_DADOS Then gError 94456

    objOrcamentoVenda.lCodigo = tOrcamentoVenda.lCodigo
    objOrcamentoVenda.iFilialEmpresa = tOrcamentoVenda.iFilialEmpresa
    objOrcamentoVenda.lNumIntNFiscal = tOrcamentoVenda.lNumIntNFiscal
    objOrcamentoVenda.sNomeFilialCli = tOrcamentoVenda.sNomeFilialCli
    objOrcamentoVenda.sNomeCli = tOrcamentoVenda.sNomeCli
    objOrcamentoVenda.iPrazoValidade = iPrazoValidade
    objOrcamentoVenda.lCliente = tOrcamentoVenda.lCliente
    objOrcamentoVenda.iFilial = tOrcamentoVenda.iFilial
    objOrcamentoVenda.iCondicaoPagto = tOrcamentoVenda.iCondicaoPagto
    objOrcamentoVenda.dPercAcrescFinanceiro = tOrcamentoVenda.dPercAcrescFinanceiro
    objOrcamentoVenda.dtDataEmissao = tOrcamentoVenda.dtDataEmissao
    objOrcamentoVenda.sNaturezaOp = tOrcamentoVenda.sNaturezaOp
    objOrcamentoVenda.dValorTotal = tOrcamentoVenda.dValorTotal
    objOrcamentoVenda.dValorFrete = tOrcamentoVenda.dValorFrete
    objOrcamentoVenda.dValorDesconto = tOrcamentoVenda.dValorDesconto
    objOrcamentoVenda.dValorOutrasDespesas = tOrcamentoVenda.dValorOutrasDespesas
    objOrcamentoVenda.dValorSeguro = tOrcamentoVenda.dValorSeguro
    objOrcamentoVenda.iTabelaPreco = tOrcamentoVenda.iTabelaPreco
    objOrcamentoVenda.dtDataReferencia = tOrcamentoVenda.dtDataReferencia
    objOrcamentoVenda.lNumIntDoc = tOrcamentoVenda.lNumIntDoc
    objOrcamentoVenda.dValorProdutos = tOrcamentoVenda.dValorProdutos
    objOrcamentoVenda.iVendedor = tOrcamentoVenda.iVendedor
    objOrcamentoVenda.iVendedor2 = tOrcamentoVenda.iVendedor2
    objOrcamentoVenda.sUsuario = tOrcamentoVenda.sUsuario
    objOrcamentoVenda.lCodigoBase = tOrcamentoVenda.lCodigoBase
    
    Call Comando_Fechar(lComando)

    OrcamentoVenda_Le_NumIntPedVenda = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Le_NumIntPedVenda:

    OrcamentoVenda_Le_NumIntPedVenda = gErr

    Select Case gErr

        Case 94456 'Tratado na rotina chamadora

        Case 94454, 94455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)
        
        Case 94453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152744)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrdemProducao_Le_Col_Empenho2(ByVal objProdutoKitProdutos As ClassProdutoKitProdutos, ByVal colItensOP As Collection, Optional ByVal bIncluiEmbProdRaiz As Boolean = True) As Long
'Le a partir do Produto/Versao todos os componentes do Kit e os retorna na colecao (não há duplicidade na colecao)
'A Embalagem do ProdutoRaiz é adicionada como último elemento a ser empenhado caso o bIncluiEmbProdRaiz = TRUE

Dim lErro As Long
Dim dFator As Double
Dim lComando As Long
Dim bAchou As Boolean
Dim dQtdEmb As Double
Dim dQuantidade As Double
Dim dQuantidadePerda As Double
Dim objProduto As New ClassProduto
Dim objProdutoKit As New ClassProdutoKit
Dim objEmbalagem As New ClassEmbalagem
Dim tProdutoKitProdutos As typeProdutoKitProdutos
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim objProdutoKitProdutosTEMP As New ClassProdutoKitProdutos
Dim objProdutoKitProdutosTEMP2 As New ClassProdutoKitProdutos

On Error GoTo Erro_OrdemProducao_Le_Col_Empenho2

    'Aertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106390
    
    'Faz a leitura do percentual de perda do produto raiz para futuro empenho da embalagem do produto raiz
    objProdutoKit.sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
    objProdutoKit.sVersao = objProdutoKitProdutos.sVersao
    
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 108585
    
    'Se nao encontrou => Erro
    If lErro = 34875 Then gError 108586
    
    'Guarda a quantidade de perda do Produto Raiz
    dQuantidadePerda = objProdutoKitProdutos.dQuantidade * objProdutoKit.dPercentualPerda
    
    'Inicializa as strings do Type
    With tProdutoKitProdutos
    
        .sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)
        .sProduto = String(STRING_PRODUTO, 0)
    
        'para obter apenas os nós do nivel 1
        lErro = Comando_Executar(lComando, "SELECT ProdutoKit.Produto, Quantidade, ProdutoKit.UnidadeMed, Composicao, ControleEstoque, SiglaUMEstoque, ClasseUM, PercentualPerda FROM ProdutoKit, Produtos WHERE ProdutoRaiz = ? AND Versao = ? AND ProdutoKit.Produto = Produtos.Codigo AND ProdutoKit.Nivel = 1", .sProduto, .dQuantidade, .sUnidadeMed, .iComposicao, .iControleEstoque, .sSiglaUMEstoque, .iClasseUM, .dPercentualPerda, objProdutoKitProdutos.sProdutoRaiz, objProdutoKitProdutos.sVersao)
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 106391

    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106392
    
    'Se nao Encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 106393
    
    'Enquanto tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS
        
        bAchou = False
        
        'Transfere os Dados para o obj
        With objProdutoKitProdutosTEMP
        
            .dQuantidade = tProdutoKitProdutos.dQuantidade
            .iClasseUM = tProdutoKitProdutos.iClasseUM
            .iComposicao = tProdutoKitProdutos.iComposicao
            .iControleEstoque = tProdutoKitProdutos.iControleEstoque
            .sProduto = tProdutoKitProdutos.sProduto
            .sSiglaUMEstoque = tProdutoKitProdutos.sSiglaUMEstoque
            .sUnidadeMed = tProdutoKitProdutos.sUnidadeMed
            .sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
            .sVersao = objProdutoKitProdutos.sVersao
            .dPercentualPerda = tProdutoKitProdutos.dPercentualPerda
        
        End With
        
        'Faz a Conversao de unidades de medida
        lErro = CF("UM_Conversao_Trans", objProdutoKitProdutosTEMP.iClasseUM, objProdutoKitProdutosTEMP.sUnidadeMed, objProdutoKitProdutosTEMP.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 106394
        
        If gobjMAT.iReqProdComPerda = MARCADO Then
            objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade / (1 - objProdutoKitProdutosTEMP.dPercentualPerda)
        End If
                
        objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade * dFator
        objProdutoKitProdutosTEMP.sUnidadeMed = objProdutoKitProdutosTEMP.sSiglaUMEstoque
        
        'Percorre na colecao para nao haver duplicidade
        For Each objProdutoKitProdutosTEMP2 In colItensOP
        
            'Verifica se já existe na colecao
            If objProdutoKitProdutosTEMP2.sProduto = objProdutoKitProdutosTEMP.sProduto Then
            
                'Adiciona à quantidade já existente
                objProdutoKitProdutosTEMP2.dQuantidade = objProdutoKitProdutosTEMP2.dQuantidade + objProdutoKitProdutosTEMP.dQuantidade
                bAchou = True
                Exit For
                    
            End If
        
        Next
            
        'Se nao encontrou na colecao => Adiciona
        If bAchou = False Then
            
            'Adiciona na colecao
            colItensOP.Add objProdutoKitProdutosTEMP
            
            'Instancia outro obj
            Set objProdutoKitProdutosTEMP = New ClassProdutoKitProdutos
            
        End If
            
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106395
    
    Loop
    
    'Se for para adicionar a Embalagem...
    If bIncluiEmbProdRaiz = True Then
        
        'Le a Embalagem do Produto Raiz
        objProdutoEmbalagem.sProduto = objProdutoKitProdutos.sProdutoRaiz
        
        lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
        If lErro <> SUCESSO And lErro <> 100000 Then gError 106396
        
        'Se encontrou => Le o Produto Associado a Embalagem Padrao
        If lErro <> 100000 Then
    
            'Le os dados da Embalagem Padrao obtida
            objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
            
            lErro = CF("Embalagem_Le", objEmbalagem)
            If lErro <> SUCESSO And lErro <> 82763 Then gError 106397
    
            'Se encontrou os Dados => Adiciona à Colecao
            If lErro <> 82763 Then
    
                objProduto.sCodigo = objEmbalagem.sProduto
                
                'Tenta ler o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 106421
        
                'Ausência de Produto no BD
                If lErro <> 28030 Then
    
                    'Instancia um novo obj
                    Set objProdutoKitProdutosTEMP = New ClassProdutoKitProdutos
                    
                    'Faz a Conversao de unidades de medida
                    lErro = CF("UM_Conversao_Trans", objProdutoKitProdutos.iClasseUM, objProdutoKitProdutos.sUnidadeMed, objProdutoEmbalagem.sUMEmbalagem, dFator)
                    If lErro <> SUCESSO Then gError 106398
                    
                    'Transfere para o objProdutoKitProdutosTEMP os dados da embalagem
                    dQuantidade = ((objProdutoKitProdutos.dQuantidade - dQuantidadePerda) * dFator) / objProdutoEmbalagem.dCapacidade
                    
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
                        .sUnidadeMed = objProduto.sSiglaUMEstoque
                        .sVersao = objProdutoKitProdutos.sVersao
                        
                    End With
                    
                    'Adiciona na colecao
                    colItensOP.Add objProdutoKitProdutosTEMP
                    
                End If
                
            End If
            
        End If
        
    End If
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

    OrdemProducao_Le_Col_Empenho2 = SUCESSO
    
    Exit Function
    
Erro_OrdemProducao_Le_Col_Empenho2:

    OrdemProducao_Le_Col_Empenho2 = gErr
    
    Select Case gErr
    
        Case 106390
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 106391, 106392, 106395
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT_PRODUTOS", gErr)
            '??? Erro ao tentar ler as tabelas de Produtos e ProdutoKit.
            
        Case 106393 'Sem Dados
        
        Case 106396, 106397, 106421, 108585
        
        Case 106394 'Se está incluindo a Embalagem => Está sendo chamada na geração de OP => Dá msg. Senão está sendo chamada pelo adrelvb e não deve dar msg.
            If bIncluiEmbProdRaiz = True Then Call Rotina_Erro(vbOKOnly, "ERRO_CONVERSAO_UM", gErr, objProdutoKitProdutosTEMP.sUnidadeMed, objProdutoKitProdutosTEMP.sProduto, objProdutoKitProdutosTEMP.sVersao, objProdutoKitProdutosTEMP2.sSiglaUMEstoque)
        
        Case 106398 'Se está incluindo a Embalagem => Está sendo chamada na geração de OP => Dá msg. Senão está sendo chamada pelo adrelvb e não deve dar msg.
            If bIncluiEmbProdRaiz = True Then Call Rotina_Erro(vbOKOnly, "ERRO_CONVERSAO_UM", gErr, objProdutoKitProdutos.sUnidadeMed, objProdutoKitProdutos.sProduto, objProdutoKitProdutos.sVersao, objProdutoEmbalagem.sUMEmbalagem)
            
        Case 108586
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRODUTO_RAIZ", gErr, objProdutoKit.sProdutoRaiz)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152745)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function OrdemProducao_Le_Col_Empenho(ByVal objProdutoKitProdutos As ClassProdutoKitProdutos, ByVal colItensOP As Collection) As Long
'Le a partir do Produto/Versao todos os componentes do Kit e os retorna na colecao (não há duplicidade na colecao)
'A Embalagem do ProdutoRaiz é adicionada como último elemento a ser empenhado caso o bIncluiEmbProdRaiz = TRUE

    If gobjMAT.iEmpenhaEmbalagem = MARCADO Then
        OrdemProducao_Le_Col_Empenho = OrdemProducao_Le_Col_Empenho2(objProdutoKitProdutos, colItensOP, True)
    Else
        OrdemProducao_Le_Col_Empenho = OrdemProducao_Le_Col_Empenho2(objProdutoKitProdutos, colItensOP, False)
    End If
    
End Function



Public Function ReservasItemGrade_Le(objItemRomaneio As ClassItemRomaneioGrade) As Long

Dim objReservaItem As ClassReservaItem
Dim tReserva As typeReserva
Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_ReservasItemGrade_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86346
    
    With tReserva
    
        .sResponsavel = String(STRING_RESERVA_RESPONSAVEL, 0)
        
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa,Almoxarifado,Quantidade,DataValidade,Responsavel,NumIntDoc FROM Reserva WHERE TipoDoc = ? AND NumIntOrigem = ?", _
        tReserva.iFilialEmpresa, tReserva.iAlmoxarifado, tReserva.dQuantidade, tReserva.dtDataValidade, tReserva.sResponsavel, tReserva.lNumIntDoc, PEDIDO_DE_VENDA_GRADE, objItemRomaneio.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 86347
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86349
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objReservaItem = New ClassReservaItem
            
            objReservaItem.dQuantidade = .dQuantidade
            objReservaItem.dtDataValidade = .dtDataValidade
            objReservaItem.iAlmoxarifado = .iAlmoxarifado
            objReservaItem.iFilialEmpresa = .iFilialEmpresa
            objReservaItem.lNumIntDoc = .lNumIntDoc
            objReservaItem.sResponsavel = .sResponsavel
            
            objItemRomaneio.colLocalizacao.Add objReservaItem
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86350
            
        Loop
        
    End With
        
    Call Comando_Fechar(lComando)
    
    ReservasItemGrade_Le = SUCESSO
    
    Exit Function

Erro_ReservasItemGrade_Le:

    ReservasItemGrade_Le = gErr
    
    Select Case gErr
    
        Case 86346
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86347, 86349, 86350
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA1", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORBECIDO_PELO_VB", gErr, Error$)
        
    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function
        
End Function



Public Function ItensOVGrade_Le(objItemOV As ClassItemOV) As Long

Dim lComando As Long
Dim lErro As Long
Dim tItemRomaneio As typeItemRomaneioGrade
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ItensOVGrade_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86353
    
    With tItemRomaneio
    
        .sProduto = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc,NumIntItemOV,Produto,Quantidade, Descricao, SiglaUMEstoque FROM ItensOrcamentoVendaGrade, Produtos WHERE Produtos.Codigo = ItensOrcamentoVendaGrade.Produto AND NumIntItemOV = ?", _
        .lNumIntDoc, .lNumIntItemPV, .sProduto, .dQuantidade, .sDescricao, .sSiglaUMEstoque, objItemOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 86354
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86355
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            objItemOV.iPossuiGrade = MARCADO
            
            Set objItemRomaneio = New ClassItemRomaneioGrade
            
            objItemRomaneio.dQuantidade = .dQuantidade
            objItemRomaneio.lNumIntDoc = .lNumIntDoc
            objItemRomaneio.sDescricao = .sDescricao
            objItemRomaneio.sProduto = .sProduto
            objItemRomaneio.sUMEstoque = .sSiglaUMEstoque
            
            objItemOV.colItensRomaneioGrade.Add objItemRomaneio
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86356
            
        Loop

    End With
    
    Call Comando_Fechar(lComando)
    
    ItensOVGrade_Le = SUCESSO
    
    Exit Function

Erro_ItensOVGrade_Le:
    
    ItensOVGrade_Le = gErr
        
    Select Case gErr
    
        Case 86353
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86354 To 86356
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVGRADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152746)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoDeVenda_Abre_Itens_Grade(objPedidoVenda As ClassPedidoDeVenda) As Long
'verifica se tem itens de grade...se tiver pergunta se quer abrir...se quiser abre os itens de pedido

Dim objItemPV As ClassItemPedido
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemPVGrade As ClassItemPedido
Dim iIndice As Integer
Dim iAchouItemGrade As Integer

On Error GoTo Erro_PedidoDeVenda_Abre_Itens_Grade

    For Each objItemPV In objPedidoVenda.colItensPedido
    
        If objItemPV.colItensRomaneioGrade.Count > 0 Then
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "DESEJA_ABRIR_ITENS_GRADE")
            Exit For
            
        End If
    
    Next

    If vbMsgRes = vbYes Then
    
        For iIndice = objPedidoVenda.colItensPedido.Count To 1 Step -1
        
            Set objItemPV = objPedidoVenda.colItensPedido(iIndice)
            
            iAchouItemGrade = 0
            
            For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
            
                iAchouItemGrade = 1
            
                Set objItemPVGrade = New ClassItemPedido
                
                objItemPVGrade.dPercDesc1 = objItemPV.dPercDesc1
                objItemPVGrade.dPercDesc2 = objItemPV.dPercDesc2
                objItemPVGrade.dPercDesc3 = objItemPV.dPercDesc3
                objItemPVGrade.dPrecoTotal = objItemPV.dPrecoTotal
                objItemPVGrade.dPrecoUnitario = objItemPV.dPrecoUnitario
                objItemPVGrade.dtDataEntrega = objItemPV.dtDataEntrega
                objItemPVGrade.dValorAbatComissao = objItemPV.dValorAbatComissao
                objItemPVGrade.dValorDescGlobal = objItemPV.dValorDescGlobal
                objItemPVGrade.dValorDesconto = objItemPV.dValorDesconto
                objItemPVGrade.iApropriacaoProd = objItemPV.iApropriacaoProd
                objItemPVGrade.iClasseUM = objItemPV.iClasseUM
                objItemPVGrade.iControleEstoque = objItemPV.iControleEstoque
                objItemPVGrade.iFilialEmpresa = objItemPV.iFilialEmpresa
                objItemPVGrade.iItem = objItemPV.iItem
                objItemPVGrade.iOrigemPedido = objItemPV.iOrigemPedido
                objItemPVGrade.iStatus = objItemPV.iStatus
                objItemPVGrade.iTipoDesc1 = objItemPV.iTipoDesc1
                objItemPVGrade.iTipoDesc2 = objItemPV.iTipoDesc2
                objItemPVGrade.iTipoDesc3 = objItemPV.iTipoDesc3
                objItemPVGrade.lCodPedido = objItemPV.lCodPedido
                objItemPVGrade.sUnidadeMed = objItemPV.sUnidadeMed
                
                objItemPVGrade.dQuantAFaturar = objItemRomaneio.dQuantAFaturar
                objItemPVGrade.dQuantCancelada = objItemRomaneio.dQuantCancelada
                objItemPVGrade.dQuantFaturada = objItemRomaneio.dQuantFaturada
                objItemPVGrade.dQuantidade = objItemRomaneio.dQuantidade
                objItemPVGrade.dQuantOP = objItemRomaneio.dQuantOP
                objItemPVGrade.dQuantReservada = objItemRomaneio.dQuantReservada
                objItemPVGrade.dQuantSC = objItemRomaneio.dQuantSC
                objItemPVGrade.lNumIntDoc = objItemRomaneio.lNumIntDoc
                objItemPVGrade.sProduto = objItemRomaneio.sProduto
                objItemPVGrade.sDescricao = objItemRomaneio.sDescricao
                objItemPVGrade.sLote = objItemRomaneio.sLote
                objItemPVGrade.sUMEstoque = objItemRomaneio.sUMEstoque
                objItemPVGrade.dPrecoTotal = objItemPVGrade.dPrecoUnitario * (objItemPV.dQuantidade - objItemPV.dQuantCancelada)
                objItemPVGrade.dPrecoTotal = objItemPVGrade.dPrecoTotal - (objItemPVGrade.dPrecoTotal * objItemPVGrade.dPercDesc1)
                objItemPVGrade.dPrecoTotal = objItemPVGrade.dPrecoTotal - (objItemPVGrade.dPrecoTotal * objItemPVGrade.dPercDesc2)
                objItemPVGrade.dPrecoTotal = objItemPVGrade.dPrecoTotal - (objItemPVGrade.dPrecoTotal * objItemPVGrade.dPercDesc3)
                
                objPedidoVenda.colItensPedido.Add objItemPVGrade.iFilialEmpresa, objItemPVGrade.lCodPedido, _
                objItemPVGrade.iItem, objItemPVGrade.sProduto, objItemPVGrade.dQuantidade, objItemPVGrade.dPrecoUnitario, _
                objItemPVGrade.dPrecoTotal, objItemPVGrade.dValorDesconto, objItemPVGrade.dtDataEntrega, _
                objItemPVGrade.sProdutoDescricao, objItemPVGrade.dValorAbatComissao, objItemPVGrade.dQuantCancelada, _
                objItemPVGrade.dQuantReservada, Nothing, objItemPVGrade.sProdutoNomeReduzido, objItemPVGrade.sUnidadeMed, _
                objItemPVGrade.sLote, objItemPVGrade.sUMEstoque, objItemPVGrade.iClasseUM, objItemPVGrade.dQuantFaturada, _
                objItemPVGrade.dQuantAFaturar, objItemPVGrade.sDescricao, objItemPVGrade.iStatus, objItemPVGrade.iControleEstoque, _
                objItemPVGrade.dQuantOP, objItemPVGrade.dQuantSC, objItemPVGrade.dPercDesc1, objItemPVGrade.iTipoDesc1, _
                objItemPVGrade.dPercDesc2, objItemPVGrade.iTipoDesc1, objItemPVGrade.dPercDesc3, _
                objItemPVGrade.iTipoDesc3, objItemPVGrade.lNumIntDoc
                
            Next
    
            If iAchouItemGrade = 1 Then objPedidoVenda.colItensPedido.Remove (iIndice)
    
        Next
    
    End If

    PedidoDeVenda_Abre_Itens_Grade = SUCESSO
    
    Exit Function
    
Erro_PedidoDeVenda_Abre_Itens_Grade:

    PedidoDeVenda_Abre_Itens_Grade = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152747)
            
    End Select
    
    Exit Function

End Function

'##############################################################
'INSERIDO POR WAGNER

Function OrdemProducao_Le_Filhos(ByVal objOrdemProducaoPai As ClassOrdemDeProducao, ByVal colOP As Collection) As Long
'Lê Todas Ordens de Produções descendentes a Ordem de Produção passada

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tOrdemProducao As typeOrdemProducao
Dim objOrdemProducao As ClassOrdemDeProducao

On Error GoTo Erro_OrdemProducao_Le_Filhos
               
    'Abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 129241
    Next
    
    With tOrdemProducao
    
        .sCodigo = String(STRING_OPCODIGO, 0)

        'Le os Filhos da OrdemDeProducaoPai
        lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, Codigo FROM OrdensDeProducao WHERE FilialEmpresa = ? AND OPGeradora = ?", .iFilialEmpresa, .sCodigo, objOrdemProducaoPai.iFilialEmpresa, objOrdemProducaoPai.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 129242

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129243

        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objOrdemProducao = New ClassOrdemDeProducao
        
            objOrdemProducao.iFilialEmpresa = .iFilialEmpresa
            objOrdemProducao.sCodigo = .sCodigo
    
            'Le os dados do Filho da OP Passada
            lErro = CF("OrdemProducao_Le1", objOrdemProducao)
            If lErro <> SUCESSO Then gError 129244
            
            lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
            If lErro <> SUCESSO Then gError 129245
            
            colOP.Add objOrdemProducao
                                    
            'Busca os Netos da OP Passada
            lErro = OrdemProducao_Le_Filhos(objOrdemProducao, colOP)
            If lErro <> SUCESSO Then gError 129246
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129243
        
        Loop
        
    End With

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemProducao_Le_Filhos = SUCESSO

    Exit Function

Erro_OrdemProducao_Le_Filhos:

    OrdemProducao_Le_Filhos = gErr

    Select Case gErr

        Case 129241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 129242, 129243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", gErr)

        Case 129244, 129245, 129246
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152748)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function OrdemProducao_Le_Versao_OV(ByVal objOP As ClassOrdemDeProducao) As Long
'Lê a versão de um item de OP no item de OV

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sVersao As String
Dim sSQL As String

On Error GoTo Erro_OrdemProducao_Le_Versao_OV
               
    'Abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131316
    Next
    
    For iIndice = 1 To objOP.colItens.Count
    
        sVersao = String(STRING_KIT_VERSAO, 0)
        
        sSQL = "SELECT IOV.VersaoKit " & _
                "FROM ItensPedidoDeVenda AS IPV,  " & _
                "ItensOrcamentoVenda AS IOV,  " & _
                "OrcamentoVenda AS OV,  " & _
                "PedidosDeVenda As PV  " & _
                "WHERE   IPV.CodPedido = ? AND IPV.Produto = ? AND IPV.FilialEmpresa = ? AND  " & _
                "OV.NumIntPedVenda = PV.NumIntDoc AND  " & _
                "PV.Codigo = IPV.CodPedido AND PV.FilialEmpresa = IPV.FilialEmpresa AND " & _
                "OV.Codigo = IOV.CodOrcamento AND OV.FilialEmpresa = IOV.FilialEmpresa AND " & _
                "IPV.Produto = IOV.Produto "
    
        'Le a Versao
        lErro = Comando_Executar(alComando(1), sSQL, sVersao, objOP.colItens.Item(iIndice).lCodPedido, objOP.colItens.Item(iIndice).sProduto, objOP.colItens.Item(iIndice).iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 131317
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131318
    
        If lErro <> AD_SQL_SEM_DADOS Then
    
            objOP.colItens.Item(iIndice).sVersao = sVersao
            
        End If
        
    Next
        
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemProducao_Le_Versao_OV = SUCESSO

    Exit Function

Erro_OrdemProducao_Le_Versao_OV:

    OrdemProducao_Le_Versao_OV = gErr

    Select Case gErr

        Case 131316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131317, 131318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", gErr)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152749)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Cliente_Le_UltimoItemPedido(ByVal objItemPedido As ClassItemPedido, ByVal lCliente As Long) As Long
'Lê Descrição e preço unitário do último Pedido de Vendas relacionado ao Produto e cliente passados

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim sUnidadeMed As String
Dim dPrecoUnitario As Double

On Error GoTo Erro_Cliente_Le_UltimoItemPedido

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 500016
        
    sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    'Procura por pedido de venda de mesmo cliente passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
    lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoUnitario FROM ItensPedidoDeVenda, PedidosDeVenda WHERE ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND PedidosDeVenda.FilialEmpresa = ? AND ItensPedidoDeVenda.Produto = ? AND PedidosDeVenda.Cliente = ? Order By PedidosDeVenda.DataEmissao DESC, PedidosDeVenda.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 500017
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500018
    
    'Se não encontrou ItemPV, erro
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Procura por pedido de venda baixado de mesmo cliente passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
        lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoUnitario FROM ItensPedidoDeVendaBaixados, PedidosDeVendaBaixados WHERE ItensPedidoDeVendaBaixados.CodPedido = PedidosDeVendaBaixados.Codigo AND ItensPedidoDeVendaBaixados.FilialEmpresa = PedidosDeVendaBaixados.FilialEmpresa AND PedidosDeVendaBaixados.FilialEmpresa = ? AND ItensPedidoDeVendaBaixados.Produto = ? AND PedidosDeVendaBaixados.Cliente = ? ORDER BY PedidosDeVendaBaixados.DataEmissao DESC, PedidosDeVendaBaixados.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 500017
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500018
        
        'Se não encontrou ItemPV, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 500019
    
    End If
    
    'Guarda preço unitário e Descrição do produto
    objItemPedido.sUnidadeMed = sUnidadeMed
    objItemPedido.sDescricao = sDescricao
    objItemPedido.dPrecoUnitario = dPrecoUnitario
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Cliente_Le_UltimoItemPedido = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_UltimoItemPedido:

    Cliente_Le_UltimoItemPedido = gErr
    
    Select Case gErr
    
        Case 500016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 500017, 500018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDO_PEDIDOVENDA", gErr)
        
        Case 500019 'ItemPV Não encontrado
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152750)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ClienteFilial_Le_UltimoItemPedido(ByVal objItemPedido As ClassItemPedido, ByVal lCliente As Long, ByVal iFilialCli As Integer) As Long
'Lê Descrição e preço unitário do último Pedido de Vendas relacionado ao Produto, cliente e FilialCliente passados

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String, sUnidadeMed As String
Dim dPrecoUnitario As Double

On Error GoTo Erro_ClienteFilial_Le_UltimoItemPedido

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94407 '500016
        
    sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    'Procura por pedido de venda de mesmo cliente e filial passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
    lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoUnitario FROM ItensPedidoDeVenda, PedidosDeVenda WHERE ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND PedidosDeVenda.FilialEmpresa = ? AND ItensPedidoDeVenda.Produto = ? AND PedidosDeVenda.Cliente = ? AND PedidosDeVenda.Filial = ? Order By PedidosDeVenda.DataEmissao DESC, PedidosDeVenda.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente, iFilialCli)
    If lErro <> AD_SQL_SUCESSO Then gError 94408 '500017
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94409 '500018
    
    'Se não encontrou ItemPV
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Procura por pedido de venda baixado de mesmo cliente e filial passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
        lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoUnitario FROM ItensPedidoDeVendaBaixados, PedidosDeVendaBaixados WHERE ItensPedidoDeVendaBaixados.CodPedido = PedidosDeVendaBaixados.Codigo AND ItensPedidoDeVendaBaixados.FilialEmpresa = PedidosDeVendaBaixados.FilialEmpresa AND PedidosDeVendaBaixados.FilialEmpresa = ? AND ItensPedidoDeVendaBaixados.Produto = ? AND PedidosDeVendaBaixados.Cliente = ? AND PedidosDeVendaBaixados.Filial = ? ORDER BY PedidosDeVendaBaixados.DataEmissao DESC, PedidosDeVendaBaixados.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente, iFilialCli)
        If lErro <> AD_SQL_SUCESSO Then gError 94410 '500017
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94411 '500018
        
        'Se não encontrou ItemPV, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 94412 '500019
    
    End If
    
    'Guarda preço unitário e Descrição do produto
    objItemPedido.sUnidadeMed = sUnidadeMed
    objItemPedido.sDescricao = sDescricao
    objItemPedido.dPrecoUnitario = dPrecoUnitario
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    ClienteFilial_Le_UltimoItemPedido = SUCESSO
    
    Exit Function
    
Erro_ClienteFilial_Le_UltimoItemPedido:

    ClienteFilial_Le_UltimoItemPedido = gErr
    
    Select Case gErr
    
        Case 94407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94408, 94409, 94410, 94411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDIDO_PEDIDOVENDA", gErr)
        
        Case 94412 'ItemPV Não encontrado
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152751)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'#####################################################################
'Inserido por Wagner 09/12/05
Public Function ItensOV_Le_Customizado(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    ItensOV_Le_Customizado = SUCESSO
End Function
'#####################################################################

'#####################################################################
'Inserido por Wagner 09/02/06
Public Function PV_Verifica_Bloqueios(ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, iBloqueiosAtivos As Integer, lComando As Long) As Long
'Verifica se tem bloqueios

Dim lErro As Long
Dim lCount As Long

On Error GoTo Erro_PV_Verifica_Bloqueios

    If lComando = 0 Then
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 141736
    End If

    'Busca Bloqueios de Crédito liberado
    lErro = Comando_Executar(lComando, "SELECT Count(*) FROM BloqueiosPV WHERE FilialEmpresa = ? AND PedidoDeVenda = ? AND DataLib = ?", lCount, iFilialEmpresa, lCodigo, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 141737

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141738

    iBloqueiosAtivos = 0

    'se o bloqueio de credito estiver aprovado ==> retona a flag indicativa
    If lErro = AD_SQL_SUCESSO Then iBloqueiosAtivos = lCount

    PV_Verifica_Bloqueios = SUCESSO

    Exit Function

Erro_PV_Verifica_Bloqueios:

    PV_Verifica_Bloqueios = gErr

    Select Case gErr

        Case 141736
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141737, 141738
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", gErr, lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130461)

    End Select

    Exit Function

End Function
'#####################################################################

'#####################################################################
'Inserido por Wagner 02/06/2006
Function MovEstoque_Le_QuantItemNF_Grade(ByVal objNFiscal As ClassNFiscal, dQuantidade As Double, ByVal iTipoMov As Integer, ByVal iItemNF As Integer, ByVal iAlmoxarifado As Integer, ByVal sProduto As String) As Long
'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoque_Le_QuantItemNF_Grade
    
    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 180058
    
    If objNFiscal.ColItensNF.Item(iItemNF).colItensRomaneioGrade.Count <> 0 Then
    
        'Lê NumIntDoc de MovimentoEstoque
        lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.Quantidade FROM TipoDocInfoNFiscal, ItensNFiscal, ItensNFiscalGrade, MovimentoEstoque WHERE TipoDocInfoNFiscal.FilialEmpresa = ? AND TipoDocInfoNFiscal.Serie = ? AND TipoDocInfoNFiscal.NumNotaFiscal = ? AND TipoDocInfoNFiscal.DataEmissao >= ? AND TipoDocInfoNFiscal.DataEmissao <= ? AND TipoDocInfoNFiscal.Status <> ? AND TipoDocInfoNFiscal.Tipo = ? AND TipoDocInfoNFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItensNFiscalGrade.NumIntItemNF  AND ItensNFiscal.Item =? AND MovimentoEstoque.NumIntDocOrigem = ItensNFiscalGrade.NumIntDoc AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND MovimentoEstoque.TipoMov = ? AND MovimentoEstoque.Almoxarifado = ? AND ItensNFiscalGrade.Produto = ?", _
                                    dQuantidade, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, DOCINFO_NF_INT_SAIDA, iItemNF, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE, iTipoMov, iAlmoxarifado, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 180059
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180060
    
        'se não encontrou o movimento de estoque referente ao item da nf e escaninho
        If lErro = AD_SQL_SEM_DADOS Then gError 180061
    
    End If
    
    'Fecha comandos
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_QuantItemNF_Grade = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_QuantItemNF_Grade:

    MovEstoque_Le_QuantItemNF_Grade = gErr
    
    Select Case gErr
    
        Case 180058
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 180059, 180060
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL_ITENSNF_MOVESTOQUE", gErr, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_SAIDA, iItemNF, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE, iTipoMov)
        
        Case 180061
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180062)

    End Select
    
    'Fecha comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
'#####################################################################

Function PV_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long
    PV_Le_Customizado = SUCESSO
End Function

Function PedidoVenda_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long
    PedidoVenda_Le_Customizado = CF("PV_Le_Customizado", objPV)
End Function

Function ItemPedidoBaixado_Le(ByVal objItemPedido As ClassItemPedido) As Long
'Pesquisa no Banco de Dados a ItemPedido cujo código do Pedido de Venda e o ítem foram passados no objeto

Dim lErro As Long
Dim lComando As Long
Dim tItemPedido As typeItemPedido

On Error GoTo Erro_ItemPedidoBaixado_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178815

    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tItemPedido.sDescricao = String(STRING_DESCRICAO_ITEM, 0)
    tItemPedido.sLote = String(STRING_LOTE_ITEM, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)

    'Pesquisa no Banco de Dados o Item Pedido
    If objItemPedido.lNumIntDoc = 0 Then
        lErro = Comando_Executar(lComando, "SELECT CodPedido , Produto, QuantOP, QuantSC, PercDesc1, TipoDesc1, PercDesc2, TipoDesc2, PercDesc3, TipoDesc3,FilialEmpresa, Quantidade, PrecoUnitario, PrecoTotal, ClasseUM, UnidadeMed, ValorDesconto, DataEntrega , Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantFaturada, QuantReservada, NumIntDoc, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco FROM ItensPedidoDeVendaBaixados WHERE FilialEmpresa = ? and CodPedido = ? AND Produto = ?", _
            tItemPedido.lCodPedido, tItemPedido.sProduto, tItemPedido.dQuantOP, tItemPedido.dQuantSC, tItemPedido.dPercDesc1, tItemPedido.iTipoDesc1, tItemPedido.dPercDesc2, tItemPedido.iTipoDesc2, tItemPedido.dPercDesc3, tItemPedido.iTipoDesc3, tItemPedido.iFilialEmpresa, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus _
            , tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.dQuantReservada, tItemPedido.lNumIntDoc, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRecebForaFaixa, tItemPedido.dQuantFaturadaAMais, tItemPedido.iPrioridade, tItemPedido.dComissao, tItemPedido.iTabelaPreco, objItemPedido.iFilialEmpresa, objItemPedido.lCodPedido, objItemPedido.sProduto)
    Else
        lErro = Comando_Executar(lComando, "SELECT CodPedido , Produto, QuantOP, QuantSC, PercDesc1, TipoDesc1, PercDesc2, TipoDesc2, PercDesc3, TipoDesc3, FilialEmpresa, Quantidade, PrecoUnitario, PrecoTotal, ClasseUM, UnidadeMed, ValorDesconto, DataEntrega , Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantFaturada, QuantReservada, NumIntDoc, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, QuantFaturadaAMais, Prioridade, Comissao, TabelaPreco FROM ItensPedidoDeVendaBaixados WHERE NumIntDoc = ?", _
            tItemPedido.lCodPedido, tItemPedido.sProduto, tItemPedido.dQuantOP, tItemPedido.dQuantSC, tItemPedido.dPercDesc1, tItemPedido.iTipoDesc1, tItemPedido.dPercDesc2, tItemPedido.iTipoDesc2, tItemPedido.dPercDesc3, tItemPedido.iTipoDesc3, tItemPedido.iFilialEmpresa, tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus _
            , tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, tItemPedido.dQuantReservada, tItemPedido.lNumIntDoc, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRecebForaFaixa, tItemPedido.dQuantFaturadaAMais, tItemPedido.iPrioridade, tItemPedido.dComissao, tItemPedido.iTabelaPreco, objItemPedido.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 178816

    'Verifica se há Item Pedido
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178817

    'Não encontrou o Item Pedido
    If lErro = AD_SQL_SEM_DADOS Then gError 178818

    'Preenche objItemPedido
    With objItemPedido

        .iFilialEmpresa = tItemPedido.iFilialEmpresa
        .sUnidadeMed = tItemPedido.sUnidadeMed
        .dPrecoTotal = tItemPedido.dPrecoTotal
        .dPrecoUnitario = tItemPedido.dPrecoUnitario
        .dQuantCancelada = tItemPedido.dQuantCancelada
        .dQuantFaturada = tItemPedido.dQuantFaturada
        .dQuantidade = tItemPedido.dQuantidade
        .dQuantReservada = tItemPedido.dQuantReservada
        .dtDataEntrega = tItemPedido.dtDataEntrega
        .dValorAbatComissao = tItemPedido.dValorAbatComissao
        .dValorDesconto = tItemPedido.dValorDesconto
        .iClasseUM = tItemPedido.iClasseUM
        .iStatus = tItemPedido.iStatus
        .sDescricao = tItemPedido.sDescricao
        .sLote = tItemPedido.sLote
        .lNumIntDoc = tItemPedido.lNumIntDoc
        .dPercentMaisReceb = tItemPedido.dPercentMaisReceb
        .dPercentMenosReceb = tItemPedido.dPercentMenosReceb
        .iRecebForaFaixa = tItemPedido.iRecebForaFaixa
        .dQuantFaturadaAMais = tItemPedido.dQuantFaturadaAMais
        .iPrioridade = tItemPedido.iPrioridade
        .dComissao = tItemPedido.dComissao
        .iTabelaPreco = tItemPedido.iTabelaPreco
        .lCodPedido = tItemPedido.lCodPedido
        .sProduto = tItemPedido.sProduto
        .dQuantOP = tItemPedido.dQuantOP
        .dQuantSC = tItemPedido.dQuantSC
        .dPercDesc1 = tItemPedido.dPercDesc1
        .iTipoDesc1 = tItemPedido.iTipoDesc1
        .dPercDesc2 = tItemPedido.dPercDesc2
        .iTipoDesc2 = tItemPedido.iTipoDesc2
        .dPercDesc3 = tItemPedido.dPercDesc3
        .iTipoDesc3 = tItemPedido.iTipoDesc3
        
    End With

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ItemPedidoBaixado_Le = SUCESSO

    Exit Function

Erro_ItemPedidoBaixado_Le:

    ItemPedidoBaixado_Le = gErr

    Select Case gErr

        Case 178815
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178816, 178817
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDABAIXADOS1", gErr)

        Case 178818

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178819)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PedidoVenda_Busca_PedidoCompra(ByVal objPedidoVenda As ClassPedidoDeVenda, lCodigoPC As Long) As Long
'Verifica se existe pedido de compra, baixada ou não antes de excluir um pedido de venda

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_PedidoVenda_Busca_PedidoCompra

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178832
    
    'Faz leitura na pedido de venda
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM PedidoCompraN WHERE FilialEmpresa=? AND CodigoPV=? AND TipoDocOrigem = ?", lCodigoPC, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, PEDIDOCOMPRA_TIPODOCORIGEM_PV)
    If lErro <> AD_SQL_SUCESSO Then gError 178833
    
    'Posiciona o ponteiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178834
    
    If lErro = AD_SQL_SEM_DADOS Then gError 178835
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    PedidoVenda_Busca_PedidoCompra = SUCESSO

    Exit Function
    
Erro_PedidoVenda_Busca_PedidoCompra:

    PedidoVenda_Busca_PedidoCompra = gErr
    
    Select Case gErr
    
        Case 178832
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178833, 178834
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)
           
        Case 178835
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178836)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function ItemPVItemRC_Le_CodigoPV(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal colobjRC As Collection) As Long
'Lê os itens RC associados ao codigo do pedido de venda

Dim lErro As Long
Dim lComando As Long
Dim lCodigoRC As Long
Dim iFilialEmpresa As Integer
Dim objRC As ClassRequisicaoCompras

On Error GoTo Erro_ItemPVItemRC_Le_CodigoPV

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178838


    'Lê Item RC a partir da Filial/Codigo do PV
    lErro = Comando_Executar(lComando, "SELECT RequisicaoCompraN.FilialEmpresa, RequisicaoCompraN.Codigo  FROM ItemPVItemRC, ItensPedidoDeVenda, ItensReqCompraN, RequisicaoCompraN WHERE ItemPV = ItensPedidoDeVenda.NumIntDoc AND ItemRC = ItensReqCompraN.NumINtDoc AND ItensReqCompraN.ReqCompra = RequisicaoCompraN.NumIntDoc AND ItensPedidoDeVenda.FilialEmpresa = ? AND CodPedido = ? ", iFilialEmpresa, lCodigoRC, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 178839

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178840

    Do While lErro = AD_SQL_SUCESSO
    
        Set objRC = New ClassRequisicaoCompras
    
        objRC.iFilialEmpresa = iFilialEmpresa
        objRC.lCodigo = lCodigoRC
    
        colobjRC.Add objRC
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178841
    
    Loop

    'Lê Item RC a partir da Filial/Codigo do PV Baixados
    lErro = Comando_Executar(lComando, "SELECT RequisicaoCompraN.FilialEmpresa, RequisicaoCompraN.Codigo  FROM ItemPVItemRC, ItensPedidoDeVendaBaixados, ItensReqCompraN, RequisicaoCompraN WHERE ItemPV = ItensPedidoDeVendaBaixados.NumIntDoc AND ItemRC = ItensReqCompraN.NumINtDoc AND ItensReqCompraN.ReqCompra = RequisicaoCompraN.NumIntDoc AND ItensPedidoDeVendaBaixados.FilialEmpresa = ? AND CodPedido = ? ", iFilialEmpresa, lCodigoRC, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 178842

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178843

    Do While lErro = AD_SQL_SUCESSO
    
        Set objRC = New ClassRequisicaoCompras
    
        objRC.iFilialEmpresa = iFilialEmpresa
        objRC.lCodigo = lCodigoRC
    
        colobjRC.Add objRC
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 178844
    
    Loop

    Call Comando_Fechar(lComando)

    ItemPVItemRC_Le_CodigoPV = SUCESSO

    Exit Function

Erro_ItemPVItemRC_Le_CodigoPV:

    ItemPVItemRC_Le_CodigoPV = gErr

    Select Case gErr

        Case 178838
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178839 To 178844
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPVITEMRC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178845)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Valida_TipoPagto(ByVal iLinha As Integer, ByVal iTipoPagto As Integer, ByVal objParcPV As ClassParcelaPedidoVenda) As Long
    Valida_TipoPagto = SUCESSO
End Function

Public Function OrcamentoVenda_Le_ItensFormPreco(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long, objItemFP As ClassItemFormPreco, objItemOV As ClassItemOV
Dim tItemFP As typeItemFormPreco, lComando As Long

On Error GoTo Erro_OrcamentoVenda_Le_ItensFormPreco

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184361
    
    For Each objItemOV In objOrcamentoVenda.colItens
    
        With tItemFP
            .sProduto = String(STRING_PRODUTO, 0)
            .sUnidMed = String(STRING_UM_SIGLA, 0)
            
            lErro = Comando_Executar(lComando, "SELECT NumIntDoc,Sequencial,Produto,UnidMed,Qtde,CustoUnit,PercentMargem,PrecoUnit,PrecoTotal,Situacao FROM ItensformPreco WHERE TipoDocOrigem=? AND NumIntDocOrigem=? ORDER BY Sequencial", _
                .lNumIntDoc, .iSequencial, .sProduto, .sUnidMed, .dQtde, .dCustoUnit, .dPercentMargem, .dPrecoUnit, .dPrecoTotal, .iSituacao, FPORIGEM_ITEMOV, objItemOV.lNumIntDoc)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184362
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184363
    
        Do While lErro = AD_SQL_SUCESSO
        
            Set objItemFP = New ClassItemFormPreco
        
            With objItemFP
                .lNumIntDoc = tItemFP.lNumIntDoc
                .iTipoDocOrigem = FPORIGEM_ITEMOV
                .lNumIntDocOrigem = objItemOV.lNumIntDoc
                .iSequencial = tItemFP.iSequencial
                .sProduto = tItemFP.sProduto
                .sUnidMed = tItemFP.sUnidMed
                .dQtde = tItemFP.dQtde
                .dCustoUnit = tItemFP.dCustoUnit
                .dPercentMargem = tItemFP.dPercentMargem
                .dPrecoUnit = tItemFP.dPrecoUnit
                .dPrecoTotal = tItemFP.dPrecoTotal
                .iSituacao = tItemFP.iSituacao
            End With
            
            objItemOV.colItensFormPreco.Add objItemFP
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184364
        
        Loop
        
    Next
    
    Call Comando_Fechar(lComando)
    
    OrcamentoVenda_Le_ItensFormPreco = SUCESSO
    
    Exit Function
    
Erro_OrcamentoVenda_Le_ItensFormPreco:

    OrcamentoVenda_Le_ItensFormPreco = gErr

    Select Case gErr

        Case 184361
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184362 To 184364
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPVITEMRC", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184360)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItemFormPreco_Obter_CotProd(ByVal lNumIntFormPreco As Long, lNumIntCotacaoProduto As Long) As Long

Dim lErro As Long, lNumIntDoc As Long, lComando As Long

On Error GoTo Erro_ItemFormPreco_Obter_CotProd

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184367
    
    lErro = Comando_Executar(lComando, "SELECT CotacaoProduto FROM CotacaoProdutoItemFP WHERE NumIntItemFormPreco=?", lNumIntDoc, lNumIntFormPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 184368
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184369
    If lErro <> AD_SQL_SUCESSO Then
        lNumIntCotacaoProduto = 0
    Else
        lNumIntCotacaoProduto = lNumIntDoc
    End If
    
    Call Comando_Fechar(lComando)
    
    ItemFormPreco_Obter_CotProd = SUCESSO
    
    Exit Function
    
Erro_ItemFormPreco_Obter_CotProd:

    ItemFormPreco_Obter_CotProd = gErr

    Select Case gErr

        Case 184367
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184368, 184369
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTPRODITEMFP", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184366)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Function Produto_ObtemMargemPadrao(ByVal sProduto As String, dMargemPadrao As Double) As Long
''obtem margem
'Dim lErro As Long, lComando As Long, dvalor1 As Double
'
'On Error GoTo Erro_Produto_ObtemMargemPadrao
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 184373
'
'    lErro = Comando_Executar(lComando, "SELECT Valor1 FROM CategoriaProdutoItem, ProdutoCategoria WHERE ProdutoCategoria.Produto = ? AND ProdutoCategoria.Categoria = ? AND ProdutoCategoria.Item = CategoriaProdutoItem.Item AND CategoriaProdutoItem.Categoria = ProdutoCategoria.Categoria", dvalor1, sProduto, "Margem de Lucro")
'    If lErro <> AD_SQL_SUCESSO Then gError 184374
'
'    lErro = Comando_BuscarProximo(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184375
'    If lErro <> AD_SQL_SUCESSO Then
'        dMargemPadrao = 0
'    Else
'        dMargemPadrao = dvalor1 / 100
'    End If
'
'    Call Comando_Fechar(lComando)
'
'    Produto_ObtemMargemPadrao = SUCESSO
'
'    Exit Function
'
'Erro_Produto_ObtemMargemPadrao:
'
'    Produto_ObtemMargemPadrao = gErr
'
'    Select Case gErr
'
'        Case 184373
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 184374, 184375
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODMARGPADRAO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184372)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function Produto_ObtemMargemPadrao(ByVal sProduto As String, dMargemPadrao As Double) As Long
'obtem margem
Dim lErro As Long, dValor As Double, objContexto As New ClassContextoPlan
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas

On Error GoTo Erro_Produto_ObtemMargemPadrao

    Set objPlanilhas = New ClassPlanilhas
    
    With objContexto
        .iFilialFaturamento = giFilialEmpresa
        .sProduto = sProduto
        .dQuantidade = 1
        .iFilialCli = 0
        .lCliente = 0
        .iVendedor = 0
        .iTabelaPreco = 0
        .iAno = Year(gdtDataAtual)
        .dtDataCustoFixo = gdtDataAtual
        .iRotinaOrigem = 4 'FORMACAO_PRECO_ANALISE_MARGCONTR
        .dTaxaDescPadrao = 0
        .dTaxaValFut = 0
        .dDiasValFut = 0
    End With
    
    With objPlanilhas
        .iTipoPlanilha = PLANILHA_TIPO_TODOS
        .iFilialEmpresa = giFilialEmpresa
        .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
        .sExpressao = "MargemOrcamento"
    End With
    
    colPlanilhas.Add objPlanilhas

    'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
    lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
    If lErro <> SUCESSO Then gError 184409
    
    dMargemPadrao = dValor / 100
    
    Produto_ObtemMargemPadrao = SUCESSO
    
    Exit Function
    
Erro_Produto_ObtemMargemPadrao:

    Produto_ObtemMargemPadrao = gErr

    Select Case gErr

        Case 184409
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184410)

    End Select
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ItemOPItemPV_Le(ByVal objItemOPItemPV As ClassItemOPItemPV) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOPItemPV As typeItemOPItemPV

On Error GoTo Erro_ItemOPItemPV_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202258

    'Alocação de espaço no buffer
    tItemOPItemPV.sCodigoOP = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemOPItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemOPItemPV.sUM = String(STRING_UM_SIGLA, 0)

    'Le a tabelaItemOPItemPV
    lErro = Comando_Executar(lComando, "SELECT NumIntDocItemOP, NumIntDocItemPV, CodigoOP, CodigoPV, FilialEmpresa, Produto, " & _
                "Quantidade, UM, Prioridade, QuantidadeProd FROM ItemOPItemPV WHERE NumIntDocItemOP= ?  AND NumIntDocItemPV= ? ", _
                tItemOPItemPV.lNumIntDocItemOP, tItemOPItemPV.lNumIntDocItemPV, tItemOPItemPV.sCodigoOP, tItemOPItemPV.lCodigoPV, _
                tItemOPItemPV.iFilialEmpresa, tItemOPItemPV.sProduto, tItemOPItemPV.dQuantidade, tItemOPItemPV.sUM, tItemOPItemPV.iPrioridade, tItemOPItemPV.dQuantidadeProd, _
                objItemOPItemPV.lNumIntDocItemOP, objItemOPItemPV.lNumIntDocItemPV)
    If lErro <> AD_SQL_SUCESSO Then gError 202259

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202260

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objItemOPItemPV.lNumIntDocItemOP = tItemOPItemPV.lNumIntDocItemOP
    objItemOPItemPV.lNumIntDocItemPV = tItemOPItemPV.lNumIntDocItemPV
    objItemOPItemPV.sCodigoOP = tItemOPItemPV.sCodigoOP
    objItemOPItemPV.lCodigoPV = tItemOPItemPV.lCodigoPV
    objItemOPItemPV.iFilialEmpresa = tItemOPItemPV.iFilialEmpresa
    objItemOPItemPV.sProduto = tItemOPItemPV.sProduto
    objItemOPItemPV.dQuantidade = tItemOPItemPV.dQuantidade
    objItemOPItemPV.sUM = tItemOPItemPV.sUM
    objItemOPItemPV.iPrioridade = tItemOPItemPV.iPrioridade
    objItemOPItemPV.dQuantidadeProd = tItemOPItemPV.dQuantidadeProd

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOPItemPV_Le = SUCESSO

    Exit Function

Erro_ItemOPItemPV_Le:

    ItemOPItemPV_Le = gErr

    Select Case gErr

        Case 202258
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202259, 202260
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPITEMPV", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202261)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItemOPItemPV_Le2(ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim tItemOPItemPV As typeItemOPItemPV
Dim objItemOPItemPV As ClassItemOPItemPV

On Error GoTo Erro_ItemOPItemPV_Le2

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202262
    
    For iIndice = objItemOP.colItemPV.Count To 1 Step -1
        objItemOP.colItemPV.Remove iIndice
    Next

    'Alocação de espaço no buffer
    tItemOPItemPV.sCodigoOP = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemOPItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemOPItemPV.sUM = String(STRING_UM_SIGLA, 0)

    'Le a tabelaItemOPItemPV
    lErro = Comando_Executar(lComando, "SELECT NumIntDocItemOP, NumIntDocItemPV, CodigoOP, CodigoPV, FilialEmpresa, Produto, " & _
                "Quantidade, UM, Prioridade, QuantidadeProd  FROM ItemOPItemPV WHERE NumIntDocItemOP= ? ORDER BY Seq ", _
                tItemOPItemPV.lNumIntDocItemOP, tItemOPItemPV.lNumIntDocItemPV, tItemOPItemPV.sCodigoOP, tItemOPItemPV.lCodigoPV, _
                tItemOPItemPV.iFilialEmpresa, tItemOPItemPV.sProduto, tItemOPItemPV.dQuantidade, tItemOPItemPV.sUM, tItemOPItemPV.iPrioridade, tItemOPItemPV.dQuantidadeProd, _
                objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 202263

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202264

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItemOPItemPV = New ClassItemOPItemPV

        objItemOPItemPV.lNumIntDocItemOP = tItemOPItemPV.lNumIntDocItemOP
        objItemOPItemPV.lNumIntDocItemPV = tItemOPItemPV.lNumIntDocItemPV
        objItemOPItemPV.sCodigoOP = tItemOPItemPV.sCodigoOP
        objItemOPItemPV.lCodigoPV = tItemOPItemPV.lCodigoPV
        objItemOPItemPV.iFilialEmpresa = tItemOPItemPV.iFilialEmpresa
        objItemOPItemPV.sProduto = tItemOPItemPV.sProduto
        objItemOPItemPV.dQuantidade = tItemOPItemPV.dQuantidade
        objItemOPItemPV.sUM = tItemOPItemPV.sUM
        objItemOPItemPV.iPrioridade = tItemOPItemPV.iPrioridade
        objItemOPItemPV.dQuantidadeProd = tItemOPItemPV.dQuantidadeProd
        
        objItemOP.colItemPV.Add objItemOPItemPV

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202265

    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOPItemPV_Le2 = SUCESSO

    Exit Function

Erro_ItemOPItemPV_Le2:

    ItemOPItemPV_Le2 = gErr

    Select Case gErr

        Case 202262
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202263 To 202265
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPITEMPV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202266)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemPV_ObtemDtPrecoDesde(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPV As ClassItemPedido, dtPrecoDesde As Date, lComando As Long) As Long

Dim lErro As Long
Dim dtDataEmissaoNF As Date, dPrecoUnitario As Double, dQuantidade As Double
Dim dFator As Double, sUM As String, iClasseUM As Integer
Dim dPrecoUnitInicial As Double, sUMItemPV As String

On Error GoTo Erro_ItemPV_ObtemDtPrecoDesde

    dtPrecoDesde = objPedidoVenda.dtDataEmissao
    dPrecoUnitInicial = objItemPV.dPrecoUnitario
    sUMItemPV = objItemPV.sUnidadeMed
    
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NFiscal.DataEmissao, ItensNFiscal.PrecoUnitario, ItensNFiscal.Quantidade, ItensNFiscal.UnidadeMed, Produtos.ClasseUM FROM NFiscal, ItensNFiscal, TiposDocInfo, Produtos WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.Status <> 7 AND TiposDocInfo.Faturamento=1 AND NFiscal.FilialEmpresa = ? AND NFiscal.Cliente = ? AND NFiscal.FilialCli = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao < ? AND (NFiscal.NumPedidoVenda <> ? OR NFiscal.FilialPedido <> ?) AND ItensNFiscal.Produto = Produtos.Codigo ORDER BY DataEmissao DESC", _
        dtDataEmissaoNF, dPrecoUnitario, dQuantidade, sUM, iClasseUM, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCliente, objPedidoVenda.iFilial, objItemPV.sProduto, objPedidoVenda.dtDataEmissao, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 184220
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184221

    Do While lErro = AD_SQL_SUCESSO
    
        'se a nf estava em outra unidade é preciso converter
        If sUM <> sUMItemPV Then
        
            lErro = CF("UM_Conversao", iClasseUM, sUMItemPV, sUM, dFator)
            If lErro <> SUCESSO Then gError 184222
            
            dPrecoUnitario = dPrecoUnitario * dFator
        
        End If
        
        If Abs(dPrecoUnitario - dPrecoUnitInicial) > DELTA_VALORMONETARIO Then
            Exit Do
        Else
            dtPrecoDesde = dtDataEmissaoNF
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184223
    
    Loop
        
    ItemPV_ObtemDtPrecoDesde = SUCESSO
     
    Exit Function
    
Erro_ItemPV_ObtemDtPrecoDesde:

    ItemPV_ObtemDtPrecoDesde = gErr
     
    Select Case gErr
          
        Case 184222
          
        Case 184220, 184221, 184223
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPV_DTPRECODESDE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function PV_Le_Andamento(ByVal objPVSel As ClassPVAndamentoSel) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim tPV As typePedidoVenda
Dim tNF As typeNFiscal
Dim top As typeOrdemProducao
Dim iStatus As Integer
Dim objPV As ClassPedidoDeVenda
Dim objNF As ClassNFiscal
Dim objOP As ClassOrdemDeProducao
Dim colNF As Collection
Dim colOP As Collection

On Error GoTo Erro_PV_Le_Andamento

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205633
    Next
    
    If objPVSel.iSoAbertos = MARCADO Then
        iStatus = STATUS_ABERTO
    Else
        iStatus = 0
    End If
    
    'Alocação de espaço no buffer
    tPV.sOBS = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le o pedido de venda
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, FilialEmpresa, Codigo, Cliente, Filial, DataEmissao, ValorTotal, DataEntrega, OBS, Andamento FROM PedVenTodos WHERE (Codigo = ? OR ? = 0) AND FilialEmpresa = ? AND (Cliente = ? OR ? = 0) AND (DataEmissao>= ? OR ? = ?) AND (DataEmissao <= ? OR ? = ?) AND (DataEntrega >= ? OR ? = ?) AND (DataEntrega <= ? OR ? = ?) AND (Status = ? OR ? = 0) ORDER BY Codigo", _
    tPV.lNumIntDoc, tPV.iFilialEmpresa, tPV.lCodigo, tPV.lCliente, tPV.iFilial, tPV.dtDataEmissao, tPV.dValorTotal, tPV.dtDataEntrega, tPV.sOBS, tPV.iAndamento, objPVSel.lPedido, objPVSel.lPedido, giFilialEmpresa, objPVSel.lCliente, objPVSel.lCliente, objPVSel.dtEmissaoDe, objPVSel.dtEmissaoDe, DATA_NULA, objPVSel.dtEmissaoAte, objPVSel.dtEmissaoAte, DATA_NULA, objPVSel.dtEntregaDe, objPVSel.dtEntregaDe, DATA_NULA, objPVSel.dtEntregaAte, objPVSel.dtEntregaAte, DATA_NULA, iStatus, iStatus)
    If lErro <> AD_SQL_SUCESSO Then gError 205634

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205635

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPV = New ClassPedidoDeVenda
        Set colNF = New Collection
        Set colOP = New Collection

        objPV.lNumIntDoc = tPV.lNumIntDoc
        objPV.iAndamento = tPV.iAndamento
        objPV.iFilialEmpresa = tPV.iFilialEmpresa
        objPV.iFilial = tPV.iFilial
        objPV.lCodigo = tPV.lCodigo
        objPV.lCliente = tPV.lCliente
        objPV.dtDataEmissao = tPV.dtDataEmissao
        objPV.dtDataEntrega = tPV.dtDataEntrega
        objPV.sOBS = tPV.sOBS
        objPV.dValorTotal = tPV.dValorTotal
        
        'Le as NFs
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, FilialEmpresa, NumNotaFiscal FROM NFiscal WHERE NumPedidoVenda = ? AND FilialPedido = ?", tNF.lNumIntDoc, tNF.iFilialEmpresa, tNF.lNumNotaFiscal, tPV.lCodigo, tPV.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 205636
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205637
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objNF = New ClassNFiscal
            colNF.Add objNF
            
            objNF.lNumIntDoc = tNF.lNumIntDoc
            objNF.lNumNotaFiscal = tNF.lNumNotaFiscal
            objNF.iFilialEmpresa = tNF.iFilialEmpresa
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205638
    
        Loop
        
        top.sCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        
        'Le as OPs
        lErro = Comando_Executar(alComando(3), "SELECT DISTINCT FilialEmpresa, Codigo FROM ItensOPTodos WHERE CodPedido = ? AND FilialPedido = ?", top.iFilialEmpresa, top.sCodigo, tPV.lCodigo, tPV.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 205639
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205640
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objOP = New ClassOrdemDeProducao
            colOP.Add objOP
            
            objOP.sCodigo = top.sCodigo
            objOP.iFilialEmpresa = top.iFilialEmpresa
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205641
    
        Loop
        
        objPVSel.colPVs.Add objPV
        objPVSel.colcolNFs.Add colNF
        objPVSel.colcolOPs.Add colOP

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205642

    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PV_Le_Andamento = SUCESSO

    Exit Function

Erro_PV_Le_Andamento:

    PV_Le_Andamento = gErr

    Select Case gErr

        Case 205633
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205634 To 205635, 205642
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", gErr)

        Case 205636 To 205638
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 205639 To 205641
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205643)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PVAndamento_Le(ByVal objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPVAndamento As typePVAndamento

On Error GoTo Erro_PVAndamento_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205704

    'Alocação de espaço no buffer
    tPVAndamento.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Le a tabelaPVAndamento
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Auto, FatorAuto FROM PVAndamento WHERE Codigo= ? ", _
                tPVAndamento.iCodigo, tPVAndamento.sDescricao, tPVAndamento.iAuto, tPVAndamento.iFatorAuto, _
                objPVAndamento.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 205705

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205706

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPVAndamento.iCodigo = tPVAndamento.iCodigo
    objPVAndamento.sDescricao = tPVAndamento.sDescricao
    objPVAndamento.iAuto = tPVAndamento.iAuto
    objPVAndamento.iFatorAuto = tPVAndamento.iFatorAuto

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PVAndamento_Le = SUCESSO

    Exit Function

Erro_PVAndamento_Le:

    PVAndamento_Le = gErr

    Select Case gErr

        Case 205704
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205705, 205706
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PVANDAMENTO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205707)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PVAndamento_Le_Todos(ByVal colPVAnd As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPVAndamento As typePVAndamento
Dim objPVAndamento As ClassPVAndamento

On Error GoTo Erro_PVAndamento_Le_Todos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 205807

    'Alocação de espaço no buffer
    tPVAndamento.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)

    'Le a tabelaPVAndamento
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Auto, FatorAuto FROM PVAndamento ORDER BY Codigo ", _
                tPVAndamento.iCodigo, tPVAndamento.sDescricao, tPVAndamento.iAuto, tPVAndamento.iFatorAuto)
    If lErro <> AD_SQL_SUCESSO Then gError 205808

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205809

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPVAndamento = New ClassPVAndamento

        objPVAndamento.iCodigo = tPVAndamento.iCodigo
        objPVAndamento.sDescricao = tPVAndamento.sDescricao
        objPVAndamento.iAuto = tPVAndamento.iAuto
        objPVAndamento.iFatorAuto = tPVAndamento.iFatorAuto
        
        colPVAnd.Add objPVAndamento

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205810

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PVAndamento_Le_Todos = SUCESSO

    Exit Function

Erro_PVAndamento_Le_Todos:

    PVAndamento_Le_Todos = gErr

    Select Case gErr

        Case 205807
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205808 To 205810
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PVANDAMENTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205811)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OVAcomp_Le(ByVal objOVAcomp As ClassOVAcomp, ByVal colOVs As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim sSQL As String
Dim lNumIntDocOV As Long
Dim dtDataProxCobr As Date
Dim dtDataPrevReceb As Date
Dim objOV As ClassOrcamentoVenda

On Error GoTo Erro_OVAcomp_Le

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206870
    Next
    
    lErro = OVAcomp_PreparaSQL(objOVAcomp, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = OVAcomp_PreparaInt(alComando(1), sSQL, giFilialEmpresa, objOVAcomp.dtDataPrevDe, objOVAcomp.dtDataPrevAte, objOVAcomp.dtDataProxDe, objOVAcomp.dtDataProxAte, objOVAcomp.dtDataEmiDe, objOVAcomp.dtDataEmiAte, objOVAcomp.iVendedor, lNumIntDocOV, dtDataProxCobr, dtDataPrevReceb)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206871

    'Enquanto existirem dados de acordo com o filtro passado
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOV = New ClassOrcamentoVenda
        
        objOV.lNumIntDoc = lNumIntDocOV
        objOV.dtDataProxCobr = dtDataProxCobr
        objOV.dtDataPrevReceb = dtDataPrevReceb
        
        lErro = CF("OrcamentoVenda_Le_NumIntDoc", objOV)
        If lErro <> SUCESSO And lErro <> 94462 Then gError ERRO_SEM_MENSAGEM
        
        colOVs.Add objOV
            
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206872
    
    Loop
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OVAcomp_Le = SUCESSO
     
    Exit Function
    
Erro_OVAcomp_Le:

    OVAcomp_Le = gErr
     
    Select Case gErr
          
        Case 206870
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 206871, 206872
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206873)
     
    End Select
    
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function OVAcomp_PreparaSQL(ByVal objOVAcomp As ClassOVAcomp, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String

On Error GoTo Erro_OVAcomp_PreparaSQL

    sSQL = "SELECT NumIntDoc, DataProxCobr ,DataPrevReceb FROM OrcamentoVenda "
    sWhere = "WHERE FilialEmpresa = ? "
    
    If objOVAcomp.iSoNaoPerdido = MARCADO Then
        sWhere = sWhere & "AND MotivoPerda = 0 "
    End If
    
    If objOVAcomp.dtDataEmiDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataEmissao >= ? "
    End If
    
    If objOVAcomp.dtDataEmiAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataEmissao <= ? "
    End If
       
    If objOVAcomp.iSoNaoFaturado = MARCADO Then
        sWhere = sWhere & "AND NumIntNFiscal = 0 AND (NumIntPedVenda = 0 OR NumIntPedVenda NOT IN (SELECT NumIntDoc FROM PedidosdeVendaBaixados)) "
    End If
       
    If objOVAcomp.iSoEmPV = MARCADO Then
        sWhere = sWhere & "AND NumIntPedVenda <> 0 "
    End If
   
    If objOVAcomp.dtDataPrevDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataPrevReceb >= ? "
    End If
    
    If objOVAcomp.dtDataPrevAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataPrevReceb <= ? "
    End If
    
    If objOVAcomp.dtDataProxDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataProxCobr >= ? "
    End If
    
    If objOVAcomp.dtDataProxAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataProxCobr <= ? "
    End If
       
    If objOVAcomp.iVendedor <> 0 Then
        sWhere = sWhere & "AND (Vendedor = ? OR Vendedor2 = ?) "
    End If
    
    sSQL = sSQL & sWhere & " ORDER BY ValorTotal DESC"
    
    OVAcomp_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_OVAcomp_PreparaSQL:

    OVAcomp_PreparaSQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206874)
     
    End Select
     
    Exit Function

End Function

Private Function OVAcomp_PreparaInt(ByVal lComando As Long, ByVal sSQL As String, ByVal viFilialEmpresa As Variant, ByVal vdtDataPrevDe As Variant, ByVal vdtDataPrevAte As Variant, ByVal vdtDataProxDe As Variant, ByVal vdtDataProxAte As Variant, ByVal vdtDataEmiDe As Variant, ByVal vdtDataEmiAte As Variant, ByVal viVendedor As Variant, vlNumIntOV As Variant, vdtDataProxCobr As Variant, vdtDataPrevReceb As Variant) As Long

Dim lErro As Long
Dim viStatus As Variant

On Error GoTo Erro_OVAcomp_PreparaInt

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 206875

    lErro = Comando_BindVarInt(lComando, vlNumIntOV)
    If (lErro <> AD_SQL_SUCESSO) Then gError 206876
    
    lErro = Comando_BindVarInt(lComando, vdtDataProxCobr)
    If (lErro <> AD_SQL_SUCESSO) Then gError 206877
    
    lErro = Comando_BindVarInt(lComando, vdtDataPrevReceb)
    If (lErro <> AD_SQL_SUCESSO) Then gError 206878
    
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 206879
    
    If vdtDataEmiDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataEmiDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206882
    End If
    
    If vdtDataEmiAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataEmiAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206883
    End If
   
    If vdtDataPrevDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataPrevDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206880
    End If
    
    If vdtDataPrevAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataPrevAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206881
    End If
    
    If vdtDataProxDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataProxDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206882
    End If
    
    If vdtDataProxAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataProxAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206883
    End If
       
    If viVendedor <> 0 Then
        lErro = Comando_BindVarInt(lComando, viVendedor)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206883
    
        lErro = Comando_BindVarInt(lComando, viVendedor)
        If (lErro <> AD_SQL_SUCESSO) Then gError 206883
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 206884
    
    OVAcomp_PreparaInt = SUCESSO
     
    Exit Function
    
Erro_OVAcomp_PreparaInt:

    OVAcomp_PreparaInt = gErr
     
    Select Case gErr

        Case 206875 To 206884
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206885)
     
    End Select
     
    Exit Function

End Function

Function InfoAdicDocItem_Le(ByVal objInfoAdicDocItem As ClassInfoAdicDocItem) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tInfoAdicDocItem As typeInfoAdicDocItem
Dim tDetExp As typeInfoAdicDocItemDetExp
Dim objDetExp As ClassInfoAdicDocItemDetExp

On Error GoTo Erro_InfoAdicDocItem_Le

    'Abertura comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201241
    Next

    With tInfoAdicDocItem
        .sMsg = String(STRING_INFOADIC_ITEM_MSG, 0)
        .sMsg2 = String(STRING_INFOADIC_ITEM_MSG, 0)
        .sNumPedidoCompra = String(STRING_INFOADIC_ITEM_NUMPEDIDOCOMPRA, 0)
        .sNumRE = String(STRING_MAXIMO, 0)
    
        lErro = Comando_Executar(alComando(1), "SELECT DataLimiteFaturamento, NumPedidoCompra, ItemPedCompra, IncluiValorTotal, Msg, Msg2, NumIntDE,NumRE FROM InfoAdicDocItem WHERE TipoDoc = ? AND NumIntDocItem = ?", _
            .dtDataLimiteFaturamento, .sNumPedidoCompra, .lItemPedCompra, .iIncluiValorTotal, .sMsg, .sMsg2, .lNumIntDE, .sNumRE, objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 201231
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201232
    If lErro <> AD_SQL_SUCESSO Then gError 201233
    
    With objInfoAdicDocItem
        .dtDataLimiteFaturamento = tInfoAdicDocItem.dtDataLimiteFaturamento
        .sNumPedidoCompra = tInfoAdicDocItem.sNumPedidoCompra
        .lItemPedCompra = tInfoAdicDocItem.lItemPedCompra
        .iIncluiValorTotal = tInfoAdicDocItem.iIncluiValorTotal
        .sMsg = tInfoAdicDocItem.sMsg + tInfoAdicDocItem.sMsg2
        .lNumIntDE = tInfoAdicDocItem.lNumIntDE
        .sNumRE = tInfoAdicDocItem.sNumRE
    
    End With
    
    Set objInfoAdicDocItem.colDetExportacao = New Collection
    
    tDetExp.sNumDrawback = String(STRING_MAXIMO, 0)
    tDetExp.sNumRegistExport = String(STRING_MAXIMO, 0)
    tDetExp.sChvNFe = String(STRING_NFE_CHNFE, 0)
    
    lErro = Comando_Executar(alComando(2), "SELECT NumDrawback, NumRegistExport, ChvNFe, QuantExport FROM InfoAdicDocItemDetExport WHERE TipoDoc = ? AND NumIntDocItem = ? ORDER BY Seq", _
        tDetExp.sNumDrawback, tDetExp.sNumRegistExport, tDetExp.sChvNFe, tDetExp.dQuantExport, objInfoAdicDocItem.iTipoDoc, objInfoAdicDocItem.lNumIntDocItem)
    If lErro <> AD_SQL_SUCESSO Then gError 201231
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201232
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objDetExp = New ClassInfoAdicDocItemDetExp
    
        objDetExp.sChvNFe = tDetExp.sChvNFe
        objDetExp.sNumDrawback = tDetExp.sNumDrawback
        objDetExp.sNumRegistExport = tDetExp.sNumRegistExport
        objDetExp.dQuantExport = tDetExp.dQuantExport
    
        objInfoAdicDocItem.colDetExportacao.Add objDetExp
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201232
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    InfoAdicDocItem_Le = SUCESSO
    
    Exit Function
    
Erro_InfoAdicDocItem_Le:

    InfoAdicDocItem_Le = gErr

    Select Case gErr

        Case 201231, 201232, 201233, 201241
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICDOCITEM", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201230)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
