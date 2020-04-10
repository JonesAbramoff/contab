VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoesPV"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ItemPedido_QuantFaltaReservar(objItemPedido As ClassItemPedido, dQuantidadeFaltaReservar As Double) As Long
'Calcula a quantidade que falta reservar deste ítem no Pedido.
'Desconta o que já tem reservado do total a reservar.
'O ítem deve ser um Produto de Reserva
'Devolve em dQuantidadeFaltaReservar a quantidade na Unidade de Estoque.
Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedido_QuantFaltaReservar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITEMPEDIDO_QUANTFALTARESERVAR", X)
    If lErro <> SUCESSO Then Error 25073
    
    
    ItemPedido_QuantFaltaReservar = X.ItemPedido_QuantFaltaReservar(objItemPedido, dQuantidadeFaltaReservar)

    Exit Function
    
Erro_ItemPedido_QuantFaltaReservar:
    
    ItemPedido_QuantFaltaReservar = Err
    
    Select Case Err

        Case 25073

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146446)
        
    End Select
    
    Exit Function

End Function


Function ItemPedido_QuantReservar(objItemPedido As ClassItemPedido, dQuantidadeReservar As Double) As Long
'Calcula a quantidade total que se deve reservar deste ítem no Pedido
'Ignora a quantidade já reservada.
'O ítem deve ser um Produto de Reserva
'Devolve em dQuantReservar a quantidade na Unidade de Estoque.
Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedido_QuantReservar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITEM_PEDIDO_QUANT_RESERVAR", X)
    If lErro <> SUCESSO Then Error 25068
    
    
    ItemPedido_QuantReservar = X.ItemPedido_QuantReservar(objItemPedido, dQuantidadeReservar)

    Exit Function
    
Erro_ItemPedido_QuantReservar:
    
    ItemPedido_QuantReservar = Err
    
    Select Case Err

        Case 25068

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146447)
        
    End Select
    
    Exit Function

End Function

Function Reserva_Automatico3(lComando As Long, lComando2 As Long, lCodigo As Long) As Long
'grava o proximo codigo de reserva disponivel
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Automatico3
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVA_AUTOMATICO3", X)
    If lErro <> SUCESSO Then Error 36702
    
    
    Reserva_Automatico3 = X.Reserva_Automatico3(lComando, lComando2, lCodigo)

    Exit Function
    
Erro_Reserva_Automatico3:
    
    Reserva_Automatico3 = Err
    
    Select Case Err

        Case 36702

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146448)
        
    End Select
    
    Exit Function

End Function

Function Reserva_Insere(lComando As Long, lComando1 As Long, objItemPedido As ClassItemPedido, dQuantidadeReserva As Double, lCodigoReserva As Long, iAlmoxarifado As Integer) As Long
'Insere uma reserva a partir dos dados passados como parametro
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Insere
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVA_INSERE", X)
    If lErro <> SUCESSO Then Error 36703
    
    
    Reserva_Insere = X.Reserva_Insere(lComando, lComando1, objItemPedido, dQuantidadeReserva, lCodigoReserva, iAlmoxarifado)

    Exit Function
    
Erro_Reserva_Insere:
    
    Reserva_Insere = Err
    
    Select Case Err

        Case 36703

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146449)
        
    End Select
    
    Exit Function

End Function

Function Reserva_Automatico1(lComando As Long, lCodigo As Long, iFilialEmpresa As Integer) As Long
'descobre o próximo código de reserva disponível e loca-o
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Automatico1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVA_AUTOMATICO1", X)
    If lErro <> SUCESSO Then Error 36704
    
    
    Reserva_Automatico1 = X.Reserva_Automatico1(lComando, lCodigo, iFilialEmpresa)

    Exit Function
    
Erro_Reserva_Automatico1:
    
    Reserva_Automatico1 = Err
    
    Select Case Err

        Case 36704

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146450)
        
    End Select
    
    Exit Function

End Function

Function Reserva_Automatico2(lComando1 As Long, lCodigo As Long) As Long
'descobre o proximo codigo de reserva disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Automatico2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVA_AUTOMATICO2", X)
    If lErro <> SUCESSO Then Error 36705
    
    
    Reserva_Automatico2 = X.Reserva_Automatico2(lComando1, lCodigo)

    Exit Function
    
Erro_Reserva_Automatico2:
    
    Reserva_Automatico2 = Err
    
    Select Case Err

        Case 36705

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146451)
        
    End Select
    
    Exit Function

End Function

Function ItemPV_ObtemQtdeReservaAlmox(objItemPV As ClassItemPedido, iAlmoxarifado As Integer, dQuantReserva As Double) As Long
'preenche em dQuantReserva a qtde reservada do item no almoxarifado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPV_ObtemQtdeReservaAlmox
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITEMPV_OBTEMQTDERESERVAALMOX", X)
    If lErro <> SUCESSO Then Error 44340
    
    
    ItemPV_ObtemQtdeReservaAlmox = X.ItemPV_ObtemQtdeReservaAlmox(objItemPV, iAlmoxarifado, dQuantReserva)

    Exit Function
    
Erro_ItemPV_ObtemQtdeReservaAlmox:
    
    ItemPV_ObtemQtdeReservaAlmox = Err
    
    Select Case Err

        Case 44340

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146452)
        
    End Select
    
    Exit Function

End Function

Function ReservasItem_Le(objItemPedido As ClassItemPedido, colReservaBD As colReservaItem) As Long
'Lê nas tabelas Reserva e Almoxarifado as reservas do item

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservasItem_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_RESERVASITEM_LE", X)
    If lErro <> SUCESSO Then Error 45572
    
    
    ReservasItem_Le = X.ReservasItem_Le(objItemPedido, colReservaBD)

    Exit Function
    
Erro_ReservasItem_Le:
    
    ReservasItem_Le = Err
    
    Select Case Err

        Case 45572

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146453)
        
    End Select
    
    Exit Function

End Function

Function ItemPedido_Le(objItemPedido As ClassItemPedido) As Long
'Pesquisa no Banco de Dados a ItemPedido cujo código do Pedido de Venda e o ítem foram passados no objeto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedido_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITEMPEDIDO_LE", X)
    If lErro <> SUCESSO Then Error 45715
    
    
    ItemPedido_Le = X.ItemPedido_Le(objItemPedido)
    
    Exit Function
    
Erro_ItemPedido_Le:

    ItemPedido_Le = Err
    
    Select Case Err

        Case 45715

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146454)
        
    End Select
    
    Exit Function
    
End Function


Function OrdemDeProducao_Grava(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ou atualiza uma ordem de producao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ORDEMDEPRODUCAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45712
    
    
    OrdemDeProducao_Grava = X.OrdemDeProducao_Grava(objOrdemDeProducao)
    
    Exit Function
    
Erro_OrdemDeProducao_Grava:

    OrdemDeProducao_Grava = Err
    
    Select Case Err

        Case 45712

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146455)
        
    End Select
    
    Exit Function
    
End Function

Function OrdemDeProducao_Exclui(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ORDEMDEPRODUCAO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45713
    
    
    OrdemDeProducao_Exclui = X.OrdemDeProducao_Exclui(objOrdemDeProducao)
    
    Exit Function
    
Erro_OrdemDeProducao_Exclui:

    OrdemDeProducao_Exclui = Err
    
    Select Case Err

        Case 45713

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146456)
        
    End Select
    
    Exit Function
    
End Function


Function PedidoDeVenda_Le_Itens(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê no BD a coleção de ítens do Pedido de Venda e devolve em colItensPedido dentro de objPedidoVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le_Itens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE_ITENS", X)
    If lErro <> SUCESSO Then Error 36998
    
    
    PedidoDeVenda_Le_Itens = X.PedidoDeVenda_Le_Itens(objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Le_Itens:
    
    PedidoDeVenda_Le_Itens = Err
    
    Select Case Err

        Case 36998

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146457)
        
    End Select
    
    Exit Function

End Function

Function BloqueioPV_Automatico(lCodigo As Long, iFilialEmpresa As Integer, iProxSeqBloqueio As Integer) As Long
'retorna o número da próximo Sequencial de Bloqueio disponivel
'OBS: ESTA FUNÇÃO DEVE SER CHAMADA DE DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioPV_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_BLOQUEIOPV_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 45776
    
    
    BloqueioPV_Automatico = X.BloqueioPV_Automatico(lCodigo, iFilialEmpresa, iProxSeqBloqueio)
    
    Exit Function
    
Erro_BloqueioPV_Automatico:

    BloqueioPV_Automatico = Err
    
    Select Case Err

        Case 45776

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146458)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueioEstoquePV_Libera(objBloqueioPV As ClassBloqueioPV) As Long
'Libera o Bloqueio do Pedido de Venda passado em objBloqueioPV

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioEstoquePV_Libera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_BLOQUEIOESTOQUEPV_LIBERA", X)
    If lErro <> SUCESSO Then Error 45774
    
    
    BloqueioEstoquePV_Libera = X.BloqueioEstoquePV_Libera(objBloqueioPV)
    
    Exit Function
    
Erro_BloqueioEstoquePV_Libera:

    BloqueioEstoquePV_Libera = Err
    
    Select Case Err

        Case 45774

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146459)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueioEstoquePV_Le(objBloqueioPV As ClassBloqueioPV) As Long
'Lê o bloqueio com o Tipo BLOQUEIO_TOTAL OU BLOQUEIO_PARCIAL
'ligado ao Pedido de Venda passado em objBloqueio

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioEstoquePV_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_BLOQUEIOESTOQUEPV_LE", X)
    If lErro <> SUCESSO Then Error 45773
    
    
    BloqueioEstoquePV_Le = X.BloqueioEstoquePV_Le(objBloqueioPV)
    
    Exit Function
    
Erro_BloqueioEstoquePV_Le:

    BloqueioEstoquePV_Le = Err
    
    Select Case Err

        Case 45773

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146460)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Automatico(lCodigo As Long) As Long
'retorna o número da proxima próxima disponível

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 20709
    
    
    Reserva_Automatico = X.Reserva_Automatico(lCodigo)
    
    Exit Function
    
Erro_Reserva_Automatico:

    Reserva_Automatico = Err
    
    Select Case Err

        Case 20709

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146461)
        
    End Select
    
    Exit Function
    
End Function

Function GeracaoOP_Grava(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'cria ordens de produção a partir de pedidos de venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoOP_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_GERACAOOP_GRAVA", X)
    If lErro <> SUCESSO Then Error 45797
    
    
    GeracaoOP_Grava = X.GeracaoOP_Grava(objOrdemDeProducao)
    
    Exit Function
    
Erro_GeracaoOP_Grava:

    GeracaoOP_Grava = Err
    
    Select Case Err

        Case 45797

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146462)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava o Pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_GRAVA", X)
    If lErro <> SUCESSO Then Error 45798
    
    
    PedidoVenda_Grava = X.PedidoVenda_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_PedidoVenda_Grava:

    PedidoVenda_Grava = Err
    
    Select Case Err

        Case 45798

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146463)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Grava_BD(objPedidoVenda As ClassPedidoDeVenda, Optional lComando1 As Long = 0) As Long
'Insere o Pedido de Venda passado na Tabela de Pedidos de Venda
'ou atualiza caso o Pedido já exista

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 45799
    
    
    PedidoVenda_Grava_BD = X.PedidoVenda_Grava_BD(objPedidoVenda, lComando1)
    
    Exit Function
    
Erro_PedidoVenda_Grava_BD:

    PedidoVenda_Grava_BD = Err
    
    Select Case Err

        Case 45799

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146464)
        
    End Select
    
    Exit Function
    
End Function

Function ComissoesPV_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere no Bd as Comissões do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_COMISSOESPV_GRAVA", X)
    If lErro <> SUCESSO Then Error 45800
    
    
    ComissoesPV_Grava = X.ComissoesPV_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_ComissoesPV_Grava:

    ComissoesPV_Grava = Err
    
    Select Case Err

        Case 45800

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146465)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelasPV_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere as no BD Parcelas do Pedido passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PARCELASPV_GRAVA", X)
    If lErro <> SUCESSO Then Error 45801
    
    
    ParcelasPV_Grava = X.ParcelasPV_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_ParcelasPV_Grava:

    ParcelasPV_Grava = Err
    
    Select Case Err

        Case 45801

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146466)
        
    End Select
    
    Exit Function
    
End Function

Function ItensPV_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere os itens do Pedido de Venda passado no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPV_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITENSPV_GRAVA", X)
    If lErro <> SUCESSO Then Error 45802
    
    
    ItensPV_Grava = X.ItensPV_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_ItensPV_Grava:

    ItensPV_Grava = Err
    
    Select Case Err

        Case 45802

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146467)
        
    End Select
    
    Exit Function
    
End Function

Function ReservasPV_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Grava as Reservas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservasPV_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVASPV_GRAVA", X)
    If lErro <> SUCESSO Then Error 45803
    
    
    ReservasPV_Grava = X.ReservasPV_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_ReservasPV_Grava:

    ReservasPV_Grava = Err
    
    Select Case Err

        Case 45803

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146468)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueioCredito_Adiciona(objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioCredito_Adiciona
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_BLOQUEIOCREDITO_ADICIONA", X)
    If lErro <> SUCESSO Then Error 45804
    
    
    BloqueioCredito_Adiciona = X.BloqueioCredito_Adiciona(objPedidoVenda)
    
    Exit Function
    
Erro_BloqueioCredito_Adiciona:

    BloqueioCredito_Adiciona = Err
    
    Select Case Err

        Case 45804

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146469)
        
    End Select
    
    Exit Function
    
End Function

Function ComissoesPV_Exclui(objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui do BD as Comissões do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_COMISSOESPV_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45805
    
    
    ComissoesPV_Exclui = X.ComissoesPV_Exclui(objPedidoVenda)
    
    Exit Function
    
Erro_ComissoesPV_Exclui:

    ComissoesPV_Exclui = Err
    
    Select Case Err

        Case 45805

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146470)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelasPV_Exclui(objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui do BD as Parcelas do Pedido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PARCELASPV_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45806
    
    
    ParcelasPV_Exclui = X.ParcelasPV_Exclui(objPedidoVenda)
    
    Exit Function
    
Erro_ParcelasPV_Exclui:

    ParcelasPV_Exclui = Err
    
    Select Case Err

        Case 45806

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146471)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueiosPV_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere no BD os Bloqueios do Pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_BLOQUEIOSPV_GRAVA", X)
    If lErro <> SUCESSO Then Error 45807
    
    
    BloqueiosPV_Grava = X.BloqueiosPV_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_BloqueiosPV_Grava:

    BloqueiosPV_Grava = Err
    
    Select Case Err

        Case 45807

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146472)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Lock(objPedidoVenda As ClassPedidoDeVenda, lComando As Long) As Long
'Lê e faz o Lock no Pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_LOCK", X)
    If lErro <> SUCESSO Then Error 45808
    
    
    PedidoVenda_Lock = X.PedidoVenda_Lock(objPedidoVenda, lComando)
    
    Exit Function
    
Erro_PedidoVenda_Lock:

    PedidoVenda_Lock = Err
    
    Select Case Err

        Case 45808

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146473)
        
    End Select
    
    Exit Function
    
End Function

Function ItensPV_Le_Lock(objPedVenda As ClassPedidoDeVenda) As Long
'Faz lock exclusive nos itens de PedVenda e lê os NumIntDoc desses itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPV_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITENSPV_LOCK", X)
    If lErro <> SUCESSO Then Error 45809
    
    
    ItensPV_Le_Lock = X.ItensPV_Le_Lock(objPedVenda)
    
    Exit Function
    
Erro_ItensPV_Le_Lock:

    ItensPV_Le_Lock = Err
    
    Select Case Err

        Case 45809

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146474)
        
    End Select
    
    Exit Function
    
End Function

Function ReservasPV_Exclui(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long) As Long
'Exclui as Reservas do pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservasPV_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVASPV_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45810
    
    
    ReservasPV_Exclui = X.ReservasPV_Exclui(objPedidoVenda, lComando, lComando1)
    
    Exit Function
    
Erro_ReservasPV_Exclui:

    ReservasPV_Exclui = Err
    
    Select Case Err

        Case 45810

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146475)
        
    End Select
    
    Exit Function
    
End Function

Function ItemOP_Le_ItemPV(objItemOp As ClassItemOP, objItemPV As ClassItemPedido) As Long
'Lê o item de Ordem de Produção vinculado ao Item de Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemOP_Le_ItemPV
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITEMOP_LE_ITEMPV", X)
    If lErro <> SUCESSO Then Error 45811
    
    
    ItemOP_Le_ItemPV = X.ItemOP_Le_ItemPV(objItemOp, objItemPV)
    
    Exit Function
    
Erro_ItemOP_Le_ItemPV:

    ItemOP_Le_ItemPV = Err
    
    Select Case Err

        Case 45811

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146476)
        
    End Select
    
    Exit Function
    
End Function

Function ItensPV_Exclui(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long) As Long
'Exclui os itens presentes na coleção de itens do Pedido de vanda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPV_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITENSPV_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45812
    
    
    ItensPV_Exclui = X.ItensPV_Exclui(objPedidoVenda, lComando, lComando1, lComando2, lComando3)
    
    Exit Function
    
Erro_ItensPV_Exclui:

    ItensPV_Exclui = Err
    
    Select Case Err

        Case 45812

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146477)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVendaBaixado_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda, faz o lock e o insere na tabela de Pedidos baixados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVendaBaixado_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDABAIXADO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45813
    
    
    PedidoVendaBaixado_Grava = X.PedidoVendaBaixado_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_PedidoVendaBaixado_Grava:

    PedidoVendaBaixado_Grava = Err
    
    Select Case Err

        Case 45813

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146478)
        
    End Select
    
    Exit Function
    
End Function

Function ItensPVBaixados_Grava(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPVBaixados_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITENSPVBAIXADOS_GRAVA", X)
    If lErro <> SUCESSO Then Error 45814
    
    
    ItensPVBaixados_Grava = X.ItensPVBaixados_Grava(objPedidoVenda)
    
    Exit Function
    
Erro_ItensPVBaixados_Grava:

    ItensPVBaixados_Grava = Err
    
    Select Case Err

        Case 45814

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146479)
        
    End Select
    
    Exit Function
    
End Function

Function ComissoesPV_Baixa(objPedidoVenda As ClassPedidoDeVenda, lComando, lComando1, lComando2) As Long
'Faz a baixa das Comissões do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Baixa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_COMISSOESPV_BAIXA", X)
    If lErro <> SUCESSO Then Error 45815
    
    
    ComissoesPV_Baixa = X.ComissoesPV_Baixa(objPedidoVenda, lComando, lComando1, lComando2)
    
    Exit Function
    
Erro_ComissoesPV_Baixa:

    ComissoesPV_Baixa = Err
    
    Select Case Err

        Case 45815

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146480)
        
    End Select
    
    Exit Function
    
End Function

Function ParcelasPV_Baixa(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long, lComando2 As Long) As Long
'Faz a baixa das Parcelas do Pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Baixa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PARCELASPV_BAIXA", X)
    If lErro <> SUCESSO Then Error 45816
    
    
    ParcelasPV_Baixa = X.ParcelasPV_Baixa(objPedidoVenda, lComando, lComando1, lComando2)
    
    Exit Function
    
Erro_ParcelasPV_Baixa:

    ParcelasPV_Baixa = Err
    
    Select Case Err

        Case 45816

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146481)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueiosPV_Baixa(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long, lComando2 As Long) As Long
'Faz a Baixa dos Bloqueios do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Baixa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_BLOQUEIOSPV_BAIXA", X)
    If lErro <> SUCESSO Then Error 45817
    
    
    BloqueiosPV_Baixa = X.BloqueiosPV_Baixa(objPedidoVenda, lComando, lComando1, lComando2)
    
    Exit Function
    
Erro_BloqueiosPV_Baixa:

    BloqueiosPV_Baixa = Err
    
    Select Case Err

        Case 45817

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146482)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Exclui_BD(objPedVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long, lComando4 As Long, lComando5 As Long) As Long
'Exclui o Pedido de Venda da tabela de Pedido de Vendas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Exclui_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_EXCLUI_BD", X)
    If lErro <> SUCESSO Then Error 45818
    
    
    PedidoVenda_Exclui_BD = X.PedidoVenda_Exclui_BD(objPedVenda, lComando, lComando1, lComando2, lComando3, lComando4, lComando5)
    
    Exit Function
    
Erro_PedidoVenda_Exclui_BD:

    PedidoVenda_Exclui_BD = Err
    
    Select Case Err

        Case 45818

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146483)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVendaBaixado_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda Baixado com o Código e a Filial passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVendaBaixado_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDABAIXADO_LE", X)
    If lErro <> SUCESSO Then Error 45819
    
    
    PedidoVendaBaixado_Le = X.PedidoVendaBaixado_Le(objPedidoVenda)
    
    Exit Function
    
Erro_PedidoVendaBaixado_Le:

    PedidoVendaBaixado_Le = Err
    
    Select Case Err

        Case 45819

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146484)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Exclui(objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui o Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45820
    
    
    PedidoVenda_Exclui = X.PedidoVenda_Exclui(objPedidoVenda)
    
    Exit Function
    
Erro_PedidoVenda_Exclui:

    PedidoVenda_Exclui = Err
    
    Select Case Err

        Case 45820

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146485)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueiosPV_Exclui(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long) As Long
'Exclui os Bloqueios do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_BLOQUEIOSPV_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45821
    
    
    BloqueiosPV_Exclui = X.BloqueiosPV_Exclui(objPedidoVenda, lComando, lComando1)
    
    Exit Function
    
Erro_BloqueiosPV_Exclui:

    BloqueiosPV_Exclui = Err
    
    Select Case Err

        Case 45821

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146486)
        
    End Select
    
    Exit Function
    
End Function

Function ItemPV_GravaTributacao(objItemPV As ClassItemPedido, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPV_GravaTributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITEMPV_GRAVATRIBUTACAO", X)
    If lErro <> SUCESSO Then Error 45822
    
    
    ItemPV_GravaTributacao = X.ItemPV_GravaTributacao(objItemPV, lComando)
    
    Exit Function
    
Erro_ItemPV_GravaTributacao:

    ItemPV_GravaTributacao = Err
    
    Select Case Err

        Case 45822

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146487)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_GravaTributacao(objPedidoVenda As ClassPedidoDeVenda, lComando1 As Long, lComando2 As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_GravaTributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_GRAVATRIBUTACAO", X)
    If lErro <> SUCESSO Then Error 45823
    
    
    PedidoVenda_GravaTributacao = X.PedidoVenda_GravaTributacao(objPedidoVenda, lComando1, lComando2)
    
    Exit Function
    
Erro_PedidoVenda_GravaTributacao:

    PedidoVenda_GravaTributacao = Err
    
    Select Case Err

        Case 45823

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146488)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_GravaTributacaoCompl(objPedidoVenda As ClassPedidoDeVenda, objTribComplPV As ClassTributacaoItemPV, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_GravaTributacaoCompl
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_GRAVATRIBUTACAOCOMPL", X)
    If lErro <> SUCESSO Then Error 45824
    
    
    PedidoVenda_GravaTributacaoCompl = X.PedidoVenda_GravaTributacaoCompl(objPedidoVenda, objTribComplPV, lComando)
    
    Exit Function
    
Erro_PedidoVenda_GravaTributacaoCompl:

    PedidoVenda_GravaTributacaoCompl = Err
    
    Select Case Err

        Case 45824

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146489)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_ExcluiTributacao(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVenda_ExcluiTributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDOVENDA_EXCLUITRIBUTACAO", X)
    If lErro <> SUCESSO Then Error 45825
    
    
    PedidoVenda_ExcluiTributacao = X.PedidoVenda_ExcluiTributacao(objPedidoVenda, lComando, lComando1, lComando2, lComando3)
    
    Exit Function
    
Erro_PedidoVenda_ExcluiTributacao:

    PedidoVenda_ExcluiTributacao = Err
    
    Select Case Err

        Case 45825

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146490)
        
    End Select
    
    Exit Function
    
End Function

Function ItemPV_ExcluiTributacao(objItemPV As ClassItemPedido, lComando As Long, lComando1 As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPV_ExcluiTributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITEMPV_EXCLUITRIBUTACAO", X)
    If lErro <> SUCESSO Then Error 45826
    
    
    ItemPV_ExcluiTributacao = X.ItemPV_ExcluiTributacao(objItemPV, lComando, lComando1)
    
    Exit Function
    
Erro_ItemPV_ExcluiTributacao:

    ItemPV_ExcluiTributacao = Err
    
    Select Case Err

        Case 45826

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146491)
        
    End Select
    
    Exit Function
    
End Function

Function ReservasPedido_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as Reservas dos itens do Pedido passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservasPedido_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVASPEDIDO_LE", X)
    If lErro <> SUCESSO Then Error 45827
    
    
    ReservasPedido_Le = X.ReservasPedido_Le(objPedidoVenda)
    
    Exit Function
    
Erro_ReservasPedido_Le:

    ReservasPedido_Le = Err
    
    Select Case Err

        Case 45827

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146492)
        
    End Select
    
    Exit Function
    
End Function

Function ComissoesPV_Altera(objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza as comissões do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_COMISSOESPV_ALTERA", X)
    If lErro <> SUCESSO Then Error 45829
    
    
    ComissoesPV_Altera = X.ComissoesPV_Altera(objPedidoVenda)
    
    Exit Function
    
Erro_ComissoesPV_Altera:

    ComissoesPV_Altera = Err
    
    Select Case Err

        Case 45829

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146493)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoDeVenda_Baixar(alComando() As Long, objPedidoDeVenda As ClassPedidoDeVenda) As Long
'transfere os dados do pedido para as tabelas de pedidos baixados
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Baixar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDOSDEVENDA_BAIXAR", X)
    If lErro <> SUCESSO Then Error 36727
    
    
    PedidoDeVenda_Baixar = X.PedidoDeVenda_Baixar(alComando(), objPedidoDeVenda)

    Exit Function
    
Erro_PedidoDeVenda_Baixar:
    
    PedidoDeVenda_Baixar = Err
    
    Select Case Err

        Case 36727

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146494)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Baixar1(alComando() As Long, objPedidoDeVenda As ClassPedidoDeVenda, iFaturaTudo As Integer) As Long
'transfere os dados do pedido para as tabelas de pedidos baixados
'IMPORTANTE: Tem que ser chamado dentro de TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Baixar1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDOSDEVENDA_BAIXAR1", X)
    If lErro <> SUCESSO Then Error 44338
    
    
    PedidoDeVenda_Baixar1 = X.PedidoDeVenda_Baixar1(alComando(), objPedidoDeVenda, iFaturaTudo)

    Exit Function
    
Erro_PedidoDeVenda_Baixar1:
    
    PedidoDeVenda_Baixar1 = Err
    
    Select Case Err

        Case 44338

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146495)
        
    End Select
    
    Exit Function

End Function

Function ParcelasPV_Altera(objPedidoVenda As ClassPedidoDeVenda) As Long
'Atualiza as pracelas do Pedido de Venda Passdados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PARCELASPV_ALTERA", X)
    If lErro <> SUCESSO Then Error 45830
    
    
    ParcelasPV_Altera = X.ParcelasPV_Altera(objPedidoVenda)
    
    Exit Function
    
Erro_ParcelasPV_Altera:

    ParcelasPV_Altera = Err
    
    Select Case Err

        Case 45830

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146496)
        
    End Select
    
    Exit Function
    
End Function

Function ReservasPV_Altera(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long)
'Atualiza as Reservas do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservasPV_Altera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVASPV_ALTERA", X)
    If lErro <> SUCESSO Then Error 45831
    
    
    ReservasPV_Altera = X.ReservasPV_Altera(objPedidoVenda)
    
    Exit Function
    
Erro_ReservasPV_Altera:

    ReservasPV_Altera = Err
    
    Select Case Err

        Case 45831

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146497)
        
    End Select
    
    Exit Function
    
End Function

Function ItensPV_ItensOP_Vinculados(objPedidoVenda As ClassPedidoDeVenda) As Long
'verifica se existe alguma Item de Ordem de Produção vincula do a algum Item do Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPV_ItensOP_Vinculados
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_ITENSPV_ITENSOP_VINCULADOS", X)
    If lErro <> SUCESSO Then Error 45832
    
    
    ItensPV_ItensOP_Vinculados = X.ItensPV_ItensOP_Vinculados(objPedidoVenda)
    
    Exit Function
    
Erro_ItensPV_ItensOP_Vinculados:

    ItensPV_ItensOP_Vinculados = Err
    
    Select Case Err

        Case 45832

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146498)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoDeVenda_Automatico(lNumPedido As Long) As Long
'Gera o próximo Número de um Pedido De Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDODEVENDA_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 45833
    
    
    PedidoDeVenda_Automatico = X.PedidoDeVenda_Automatico(lNumPedido)
    
    Exit Function
    
Erro_PedidoDeVenda_Automatico:

    PedidoDeVenda_Automatico = Err
    
    Select Case Err

        Case 45833

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146499)
        
    End Select
    
    Exit Function
    
End Function

Function PVObjetosVinculados_Lock(objPedidoVenda As ClassPedidoDeVenda) As Long
'Faz "lock" nos objetos associados ao Pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PVObjetosVinculados_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PVOBJETOSVINCULADOS_LOCK", X)
    If lErro <> SUCESSO Then Error 45834
    
    
    PVObjetosVinculados_Lock = X.PVObjetosVinculados_Lock(objPedidoVenda)
    
    Exit Function
    
Erro_PVObjetosVinculados_Lock:

    PVObjetosVinculados_Lock = Err
    
    Select Case Err

        Case 45834

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146500)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueiosPV_Credito_Aprovado_Testa(objPedidoVenda As ClassPedidoDeVenda, iCreditoAprovado As Integer) As Long
'Verifica se tem bloqueio de crédito liberado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Credito_Aprovado_Testa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIOSPV_CREDITO_APROVADO_TESTA", X)
    If lErro <> SUCESSO Then Error 44383
    
    
    BloqueiosPV_Credito_Aprovado_Testa = X.BloqueiosPV_Credito_Aprovado_Testa(objPedidoVenda, iCreditoAprovado)

    Exit Function
    
Erro_BloqueiosPV_Credito_Aprovado_Testa:
    
    BloqueiosPV_Credito_Aprovado_Testa = Err
    
    Select Case Err

        Case 44383

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146501)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPV_Exclui_Credito(objPedidoVenda As ClassPedidoDeVenda) As Long
'Exclui os Bloqueios de Creditos não liberados ligados ao Pedido de Venda passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Exclui_Credito
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIOSPV_EXCLUI_CREDITO", X)
    If lErro <> SUCESSO Then Error 44391
    
    
    BloqueiosPV_Exclui_Credito = X.BloqueiosPV_Exclui_Credito(objPedidoVenda)

    Exit Function
    
Erro_BloqueiosPV_Exclui_Credito:
    
    BloqueiosPV_Exclui_Credito = Err
    
    Select Case Err

        Case 44391

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146502)
        
    End Select
    
    Exit Function

End Function

Function PedVenda_Grava_Estatisticas(objPedVenda As ClassPedidoDeVenda, Optional objPedVenda_BD As ClassPedidoDeVenda) As Long
'Grava a estatística dos produtos Pedidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedVenda_Grava_Estatisticas
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDVENDA_GRAVA_ESTATISTICAS", X)
    If lErro <> SUCESSO Then Error 54665
    
    
    PedVenda_Grava_Estatisticas = X.PedVenda_Grava_Estatisticas(objPedVenda, objPedVenda_BD)
    
    Exit Function
    
Erro_PedVenda_Grava_Estatisticas:

    PedVenda_Grava_Estatisticas = Err
    
    Select Case Err

        Case 54665

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146503)
        
    End Select
    
    Exit Function
    
End Function

Function SldMesFat_Le_Lock(lComando As Long, objSldMesFat As ClassSldMesFat) As Long
'Lê todos atributos da tabela SldMesFat e faz lock exclusive
'Filtro: filial/produto/ano passados em objSldMesFat

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesFat_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_SLDMESFAT_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 54666
    
    
    SldMesFat_Le_Lock = X.SldMesFat_Le_Lock(lComando, objSldMesFat)
    
    Exit Function
    
Erro_SldMesFat_Le_Lock:

    SldMesFat_Le_Lock = Err
    
    Select Case Err

        Case 54666

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146504)
        
    End Select
    
    Exit Function
    
End Function

Function SldDiaFat_Le_Lock(lComando As Long, objSldDiaFat As ClassSldDiaFat) As Long
'Lê todos atributos da tabela SldDiaFat e faz lock exclusive
'Filtro: filial/produto/data passados em objSldDiaFat

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldDiaFat_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("PV_SLDDIAFAT_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 54667
    
    
    SldDiaFat_Le_Lock = X.SldDiaFat_Le_Lock(lComando, objSldDiaFat)
    
    Exit Function
    
Erro_SldDiaFat_Le_Lock:

    SldDiaFat_Le_Lock = Err
    
    Select Case Err

        Case 54667

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146505)
        
    End Select
    
    Exit Function
    
End Function

Function Cliente_Le_Estendida(objCliente As ClassCliente, objClienteEstatistica As ClassFilialClienteEst) As Long
'Le todos os Dados do cliente inclusive os dados de Estatistica

Dim X As Object, lErro As Long

On Error GoTo Erro_Cliente_Le_Estendida

    
    

    lErro = Rotina_ObterObjetoFuncao("FAT_CLIENTE_LE_ESTENDIDA", X)
    If lErro <> SUCESSO Then Error 54524

    
    Cliente_Le_Estendida = X.Cliente_Le_Estendida(objCliente, objClienteEstatistica)

    Exit Function

Erro_Cliente_Le_Estendida:

    Cliente_Le_Estendida = Err

    Select Case Err

        Case 54524
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146506)

    End Select

    Exit Function

End Function

Function Cliente_Le_Estatistica(objClienteEstatistica As ClassFilialClienteEst) As Long
'Le os Dados de Estatistica do Cliente

Dim X As Object, lErro As Long

On Error GoTo Erro_Cliente_Le_Estatistica

    
    

    lErro = Rotina_ObterObjetoFuncao("FAT_CLIENTE_LE_ESTATISTICA", X)
    If lErro <> SUCESSO Then Error 54525

    
    Cliente_Le_Estatistica = X.Cliente_Le_Estatistica(objClienteEstatistica)

    Exit Function

Erro_Cliente_Le_Estatistica:

    Cliente_Le_Estatistica = Err

    Select Case Err

        Case 54525
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146507)

    End Select

    Exit Function

End Function

Function Cliente_Le_Estatistica_Credito(objClienteEstatistica As ClassFilialClienteEst) As Long
'Le os Dados de Estatistica do Cliente para Crédito

Dim X As Object, lErro As Long

On Error GoTo Erro_Cliente_Le_Estatistica_Credito

    
    

    lErro = Rotina_ObterObjetoFuncao("PV_CLIENTE_LE_ESTATISTICA_CREDITO", X)
    If lErro <> SUCESSO Then Error 59586

    
    Cliente_Le_Estatistica_Credito = X.Cliente_Le_Estatistica_Credito(objClienteEstatistica)

    Exit Function

Erro_Cliente_Le_Estatistica_Credito:

    Cliente_Le_Estatistica_Credito = Err

    Select Case Err

        Case 59586
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146508)

    End Select

    Exit Function

End Function

Function Cliente_Le_NomeRed_Estendida(objCliente As ClassCliente, objClienteEstatistica As ClassFilialClienteEst) As Long
'Le todos os Dados do cliente inclusive os dados de Estatistica com o NomeReduzido sendo passado como parametro

Dim X As Object, lErro As Long

On Error GoTo Erro_Cliente_Le_NomeRed_Estendida

    
    

    lErro = Rotina_ObterObjetoFuncao("FAT_CLIENTE_LE_NOMERED_ESTENDIDA", X)
    If lErro <> SUCESSO Then Error 54526

    
    Cliente_Le_NomeRed_Estendida = X.Cliente_Le_NomeRed_Estendida(objCliente, objClienteEstatistica)

    Exit Function

Erro_Cliente_Le_NomeRed_Estendida:

    Cliente_Le_NomeRed_Estendida = Err

    Select Case Err

        Case 54526
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146509)

    End Select

    Exit Function

End Function

Function FilCliFilEmp_Le(objFilCliFilEmp As ClassFilCliFilEmp) As Long
'Le em FilialClienteFilEmp os dados estatísticos de compras da filial cliente

Dim X As Object, lErro As Long

On Error GoTo Erro_FilCliFilEmp_Le

    
    

    lErro = Rotina_ObterObjetoFuncao("FAT_FILCLIFILEMP_LE", X)
    If lErro <> SUCESSO Then Error 54529

    
    FilCliFilEmp_Le = X.FilCliFilEmp_Le(objFilCliFilEmp)

    Exit Function

Erro_FilCliFilEmp_Le:

    FilCliFilEmp_Le = Err

    Select Case Err

        Case 54529
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146510)

    End Select

    Exit Function

End Function

Function FilialCliente_Le_Estendida(objFilialCliente As ClassFilialCliente, objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le todos os Dados da Filial Cliente inclusive os dados de Estatística

Dim X As Object, lErro As Long

On Error GoTo Erro_FilialCliente_Le_Estendida

    
    

    lErro = Rotina_ObterObjetoFuncao("FAT_FILIALCLIENTE_LE_ESTENDIDA", X)
    If lErro <> SUCESSO Then Error 54530

    
    FilialCliente_Le_Estendida = X.FilialCliente_Le_Estendida(objFilialCliente, objFilialClienteEstatistica)

    Exit Function

Erro_FilialCliente_Le_Estendida:

    FilialCliente_Le_Estendida = Err

    Select Case Err

        Case 54530
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146511)

    End Select

    Exit Function

End Function

Function FilialClienteEstatistica_Le(objFilialClienteEstatistica As ClassFilialClienteEst) As Long
'Le os Dados de Estatística da Filial Cliente

Dim X As Object, lErro As Long

On Error GoTo Erro_FilialClienteEstatistica_Le

    
    

    lErro = Rotina_ObterObjetoFuncao("FAT_FILIALCLIENTEESTATISTICA_LE", X)
    If lErro <> SUCESSO Then Error 54531

    
    FilialClienteEstatistica_Le = X.FilialClienteEstatistica_Le(objFilialClienteEstatistica)

    Exit Function

Erro_FilialClienteEstatistica_Le:

    FilialClienteEstatistica_Le = Err

    Select Case Err

        Case 54531
            

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146512)

    End Select

    Exit Function

End Function

'Function FilialFornecedor_Grava_EstatisticasDev(objNFiscal As ClassNFiscal) As Long
''Tem que ser chamado dentro de Transação
''Atualiza a gravação da parte de estatística de devolução de Fornecedor
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FilialFornecedor_Grava_EstatisticasDev
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("PV_FILIALFORNECEDOR_GRAVA_ESTATISTICADEV", X)
'    If lErro <> SUCESSO Then Error 54711
'
'
'    FilialFornecedor_Grava_EstatisticasDev = X.FilialFornecedor_Grava_EstatisticasDev(objNFiscal)
'
'    Exit Function
'
'Erro_FilialFornecedor_Grava_EstatisticasDev:
'
'    FilialFornecedor_Grava_EstatisticasDev = Err
'
'    Select Case Err
'
'        Case 54711
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146513)
'
'    End Select
'
'    Exit Function
'
'End Function


Function PedidosBaixar_Batch(colPVInfo As Collection) As Long
'recebe colecao de PVInfo
'transfere pedidos de venda p/pedidos de venda baixados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosBaixar_Batch
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDOSBAIXAR_BATCH", X)
    If lErro <> SUCESSO Then Error 36726
    
    
    PedidosBaixar_Batch = X.PedidosBaixar_Batch(colPVInfo)

    Exit Function
    
Erro_PedidosBaixar_Batch:
    
    PedidosBaixar_Batch = Err
    
    Select Case Err

        Case 36726

    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146514)
        
    End Select
    
    Exit Function

End Function

Function ReservasItemPV_Le_NumIntOrigem(objItemPV As ClassItemPedido, colReservaBD As colReservaItem) As Long
'Lê nas tabelas Reserva e Almoxarifado as reservas do item de pedido de venda com o NumintDoc passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservasItemPV_Le_NumIntOrigem
    
    lErro = Rotina_ObterObjetoFuncao("PV_RESERVASITEMPV_LE_NUMINTORIGEM", X)
    If lErro <> SUCESSO Then Error 59872
    
    ReservasItemPV_Le_NumIntOrigem = X.ReservasItemPV_Le_NumIntOrigem(objItemPV, colReservaBD)

    Exit Function
    
Erro_ReservasItemPV_Le_NumIntOrigem:
    
    ReservasItemPV_Le_NumIntOrigem = Err
    
    Select Case Err

        Case 59872
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146515)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Baixar_Pedido(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Baixar_Pedido
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDODEVENDA_BAIXAR_PEDIDO", X)
    If lErro <> SUCESSO Then Error 65084
    
    PedidoDeVenda_Baixar_Pedido = X.PedidoDeVenda_Baixar_Pedido(objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Baixar_Pedido:
    
    PedidoDeVenda_Baixar_Pedido = Err
    
    Select Case Err

        Case 65084
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146516)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVendaBaixado_LeMod(lComando As Long, objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVendaBaixado_LeMod
    
    lErro = Rotina_ObterObjetoFuncao("PV_PEDIDODEVENDABAIXADO_LEMOD", X)
    If lErro <> SUCESSO Then Error 65104
    
    PedidoDeVendaBaixado_LeMod = X.PedidoDeVendaBaixado_LeMod(lComando, objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVendaBaixado_LeMod:
    
    PedidoDeVendaBaixado_LeMod = Err
    
    Select Case Err

        Case 65104
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146517)
        
    End Select
    
    Exit Function

End Function

Function TributacaoPV_Desfaz_Baixa(objPedidoVenda As ClassPedidoDeVenda) As Long
'Baixa Tributação relacionada ao pedido de vendas passado como parâmetro
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TributacaoPV_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("PV_TRIBUTACAOPV_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then Error 65105
    
    TributacaoPV_Desfaz_Baixa = X.TributacaoPV_Desfaz_Baixa(objPedidoVenda)

    Exit Function
    
Erro_TributacaoPV_Desfaz_Baixa:
    
    TributacaoPV_Desfaz_Baixa = Err
    
    Select Case Err

        Case 65105
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146518)
        
    End Select
    
    Exit Function

End Function

Function ParcelasPV_Desfaz_Baixa(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long) As Long
'Desfaz a baixa das parcelas de pedido de venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("PV_PARCELASPV_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then Error 65106
    
    ParcelasPV_Desfaz_Baixa = X.ParcelasPV_Desfaz_Baixa(objPedidoVenda, lComando, lComando1)

    Exit Function
    
Erro_ParcelasPV_Desfaz_Baixa:
    
    ParcelasPV_Desfaz_Baixa = Err
    
    Select Case Err

        Case 65106
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146519)
        
    End Select
    
    Exit Function

End Function

Function ComissoesPV_Desfaz_Baixa(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long) As Long
'Desfaz a Baixa das comissoes do PV

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("PV_COMISSOESPV_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then Error 65107
    
    ComissoesPV_Desfaz_Baixa = X.ComissoesPV_Desfaz_Baixa(objPedidoVenda, lComando, lComando1)

    Exit Function
    
Erro_ComissoesPV_Desfaz_Baixa:
    
    ComissoesPV_Desfaz_Baixa = Err
    
    Select Case Err

        Case 65107
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146520)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPV_Desfaz_Baixa(objPedidoVenda As ClassPedidoDeVenda, lComando As Long, lComando1 As Long) As Long
'Besfaz a baixa nos Bloqueios de pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("PV_BLOQUEIOSPV_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then Error 65108
    
    BloqueiosPV_Desfaz_Baixa = X.BloqueiosPV_Desfaz_Baixa(objPedidoVenda, lComando, lComando1)

    Exit Function
    
Erro_BloqueiosPV_Desfaz_Baixa:
    
    BloqueiosPV_Desfaz_Baixa = Err
    
    Select Case Err

        Case 65108
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146521)
        
    End Select
    
    Exit Function

End Function

Function ParcelasPV_Baixados_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as parcelas associadas ao objPedidoVenda Baixados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Baixados_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PARCELASPV_BAIXADOS_LE", X)
    If lErro <> SUCESSO Then Error 59881
    
    ParcelasPV_Baixados_Le = X.ParcelasPV_Baixados_Le(objPedidoVenda)

    Exit Function
    
Erro_ParcelasPV_Baixados_Le:
    
    ParcelasPV_Baixados_Le = Err
    
    Select Case Err

        Case 59881
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146522)
        
    End Select
    
    Exit Function

End Function


Function ComissoesPV_Baixadas_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as Comissões Baixadas associadas ao Pedido de Venda passado e objPedidoVenda.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Baixadas_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_COMISSOESPV_BAIXADAS_LE", X)
    If lErro <> SUCESSO Then Error 59879
    
    ComissoesPV_Baixadas_Le = X.ComissoesPV_Baixadas_Le(objPedidoVenda)

    Exit Function
    
Erro_ComissoesPV_Baixadas_Le:
    
    ComissoesPV_Baixadas_Le = Err
    
    Select Case Err

        Case 59879
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146523)
        
    End Select
    
    Exit Function

End Function


Function BloqueiosPV_Baixados_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê os Bloqueios Baixados associados ao objPedidoVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Baixados_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIOSPV_BAIXADOS_LE", X)
    If lErro <> SUCESSO Then Error 59880
    
    BloqueiosPV_Baixados_Le = X.BloqueiosPV_Baixados_Le(objPedidoVenda)

    Exit Function
    
Erro_BloqueiosPV_Baixados_Le:
    
    BloqueiosPV_Baixados_Le = Err
    
    Select Case Err

        Case 59880
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146524)
        
    End Select
    
    Exit Function

End Function

'**** VERSÃO 2

Function Produto_Le_QuantEmPedidoCompra(sProduto As String, dQuantEmPedido As Double) As Long
'Lê a quantidade do produto passado que foram requisitadas para
'estoque e não estão ligadas a ordens de produção, ou seja , não estão amarradas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_QuantEmPedidoCompra
    
    lErro = Rotina_ObterObjetoFuncao("PV_PRODUTO_LE_QUANTEMPEDIDOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71173

    Produto_Le_QuantEmPedidoCompra = X.Produto_Le_QuantEmPedidoCompra(sProduto, dQuantEmPedido)

    Exit Function
    
Erro_Produto_Le_QuantEmPedidoCompra:
    
    Produto_Le_QuantEmPedidoCompra = gErr
    
    Select Case gErr

        Case 71173
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146525)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_QuantRequisitada(sProduto As String, dQuantEmRequisicao As Double) As Long
'Lê a quantidade do produto passado que foram requisitadas para
'estoque e não estão ligadas a ordens de produção, ou seja , não estão amarradas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_QuantRequisitada
    
    lErro = Rotina_ObterObjetoFuncao("PV_PRODUTO_LE_QUANTREQUISITADA", X)
    If lErro <> SUCESSO Then gError 71174

    Produto_Le_QuantRequisitada = X.Produto_Le_QuantRequisitada(sProduto, dQuantEmRequisicao)

    Exit Function
    
Erro_Produto_Le_QuantRequisitada:
    
    Produto_Le_QuantRequisitada = gErr
    
    Select Case gErr

        Case 71174
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146526)
        
    End Select
    
    Exit Function

End Function

Function Produto_Calcula_QuantTransferencia(sCodProduto As String, dQuantTransf As Double) As Long
'Calcula a quantidade em Transferencia do produto passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Calcula_QuantTransferencia
    
    lErro = Rotina_ObterObjetoFuncao("PV_PRODUTO_LE_QUANTTRANSFERENCIA", X)
    If lErro <> SUCESSO Then gError 71175

    Produto_Calcula_QuantTransferencia = X.Produto_Calcula_QuantTransferencia(sCodProduto, dQuantTransf)

    Exit Function
    
Erro_Produto_Calcula_QuantTransferencia:
    
    Produto_Calcula_QuantTransferencia = gErr
    
    Select Case gErr

        Case 71175
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146527)
        
    End Select
    
    Exit Function

End Function

Function IntervaloProducao_Le(iIntervalo As Integer) As Long
'Busca em MATConfig o Intervalo entre o fim da produção dos insumos de um produto e
'o início na produção desse produto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_IntervaloProducao_Le
    
    lErro = Rotina_ObterObjetoFuncao("PV_INTERVALOPRODUCAO_LE", X)
    If lErro <> SUCESSO Then gError 71262

    IntervaloProducao_Le = X.IntervaloProducao_Le(iIntervalo)

    Exit Function
    
Erro_IntervaloProducao_Le:
    
    IntervaloProducao_Le = gErr
    
    Select Case gErr

        Case 71262
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146528)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Le_QuantItemNF(objNFiscal As ClassNFiscal, dQuantidade As Double, iTipoMov As Integer, iItemNF As Integer, iAlmoxarifado As Integer) As Long
'Lê a quantidade do movimento de estoque referente ao item e escaninho da nota fiscal passados como parametros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_QuantItemNF
    
    lErro = Rotina_ObterObjetoFuncao("PV_MOVESTOQUE_LE_QUANTITEMNF", X)
    If lErro <> SUCESSO Then gError 88010

    MovEstoque_Le_QuantItemNF = X.MovEstoque_Le_QuantItemNF(objNFiscal, dQuantidade, iTipoMov, iItemNF, iAlmoxarifado)

    Exit Function
    
Erro_MovEstoque_Le_QuantItemNF:
    
    MovEstoque_Le_QuantItemNF = gErr
    
    Select Case gErr

        Case 88010
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146529)
        
    End Select
    
    Exit Function

End Function

