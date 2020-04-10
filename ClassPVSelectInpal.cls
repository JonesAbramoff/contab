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

'??? "ERRO_LEITURA_EMBALAGENS_PV"

'Alterações INPAL
'1 Marcio - 08/2000 - Função Cliente_Le_UltimoItemPedido foi criada
'2 Marcio - 09/2000 - Função de leitura de Itens de Pedido  de venda foi alterada para incluir preço Base


'************************** Inicio - Alterações INPAL **************************

'1 Marcio - 08/2000 - Função Cliente_Le_UltimoItemPedido foi criada

Function Cliente_Le_UltimoItemPedido(ByVal objItemPedido As ClassItemPedido, ByVal lCliente As Long) As Long
'Lê Descrição e preço unitário do último Pedido de Vendas relacionado ao Produto e cliente passados

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
Dim dPrecoBase As Double, sUnidadeMed As String
Dim dPrecoUnitario As Double

On Error GoTo Erro_Cliente_Le_UltimoItemPedido

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 500016
        
    sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    'Procura por pedido de venda de mesmo cliente passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
    lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoBase, PrecoUnitario FROM ItensPedidoDeVenda, PedidosDeVenda WHERE ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND PedidosDeVenda.FilialEmpresa = ? AND ItensPedidoDeVenda.Produto = ? AND PedidosDeVenda.Cliente = ? Order By PedidosDeVenda.DataEmissao DESC, PedidosDeVenda.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoBase, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 500017
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500018
    
    'Se não encontrou ItemPV, erro
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Procura por pedido de venda baixado de mesmo cliente passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
        lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoBase, PrecoUnitario FROM ItensPedidoDeVendaBaixados, PedidosDeVendaBaixados WHERE ItensPedidoDeVendaBaixados.CodPedido = PedidosDeVendaBaixados.Codigo AND ItensPedidoDeVendaBaixados.FilialEmpresa = PedidosDeVendaBaixados.FilialEmpresa AND PedidosDeVendaBaixados.FilialEmpresa = ? AND ItensPedidoDeVendaBaixados.Produto = ? AND PedidosDeVendaBaixados.Cliente = ? ORDER BY PedidosDeVendaBaixados.DataEmissao DESC, PedidosDeVendaBaixados.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoBase, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 500017
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500018
        
        'Se não encontrou ItemPV, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 500019
    
    End If
    
    'Guarda preço unitário e Descrição do produto
    objItemPedido.sUnidadeMed = sUnidadeMed
    objItemPedido.sDescricao = sDescricao
    objItemPedido.dPrecoBase = dPrecoBase
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'************************** Fim - Alterações INPAL **************************
'Incluida por Leo em 04/03/02

Function ClienteFilial_Le_UltimoItemPedido(ByVal objItemPedido As ClassItemPedido, ByVal lCliente As Long, ByVal iFilialCli As Integer) As Long
'Lê Descrição e preço unitário do último Pedido de Vendas relacionado ao Produto, cliente e FilialCliente passados

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String, sUnidadeMed As String
Dim dPrecoBase As Double
Dim dPrecoUnitario As Double

On Error GoTo Erro_ClienteFilial_Le_UltimoItemPedido

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94407 '500016
        
    sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    'Procura por pedido de venda de mesmo cliente e filial passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
    lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoBase, PrecoUnitario FROM ItensPedidoDeVenda, PedidosDeVenda WHERE ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND PedidosDeVenda.FilialEmpresa = ? AND ItensPedidoDeVenda.Produto = ? AND PedidosDeVenda.Cliente = ? AND PedidosDeVenda.Filial = ? Order By PedidosDeVenda.DataEmissao DESC, PedidosDeVenda.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoBase, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente, iFilialCli)
    If lErro <> AD_SQL_SUCESSO Then gError 94408 '500017
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94409 '500018
    
    'Se não encontrou ItemPV
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Procura por pedido de venda baixado de mesmo cliente e filial passado e com o ItemPV com o mesmo Produto passado por ordem descrescente de DataEmissao
        lErro = Comando_Executar(lComando, "SELECT UnidadeMed, Descricao, PrecoBase, PrecoUnitario FROM ItensPedidoDeVendaBaixados, PedidosDeVendaBaixados WHERE ItensPedidoDeVendaBaixados.CodPedido = PedidosDeVendaBaixados.Codigo AND ItensPedidoDeVendaBaixados.FilialEmpresa = PedidosDeVendaBaixados.FilialEmpresa AND PedidosDeVendaBaixados.FilialEmpresa = ? AND ItensPedidoDeVendaBaixados.Produto = ? AND PedidosDeVendaBaixados.Cliente = ? AND PedidosDeVendaBaixados.Filial = ? ORDER BY PedidosDeVendaBaixados.DataEmissao DESC, PedidosDeVendaBaixados.NumIntDoc DESC", sUnidadeMed, sDescricao, dPrecoBase, dPrecoUnitario, objItemPedido.iFilialEmpresa, objItemPedido.sProduto, lCliente, iFilialCli)
        If lErro <> AD_SQL_SUCESSO Then gError 94410 '500017
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94411 '500018
        
        'Se não encontrou ItemPV, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 94412 '500019
    
    End If
    
    'Guarda preço unitário e Descrição do produto
    objItemPedido.sUnidadeMed = sUnidadeMed
    objItemPedido.sDescricao = sDescricao
    objItemPedido.dPrecoBase = dPrecoBase
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PV_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda Baixado com o Código e a Filial passados

Dim lErro As Long
Dim lComando As Long
Dim iProgramacao As Integer

On Error GoTo Erro_PV_Le_Customizado

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178565

    lErro = Comando_Executar(lComando, "SELECT Programacao FROM PedidosdeVenda WHERE Codigo = ? AND FilialEmpresa = ? ", iProgramacao, objPV.lCodigo, objPV.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 178566

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178567
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(lComando, "SELECT Programacao FROM PedidosdeVendaBaixados WHERE Codigo = ? AND FilialEmpresa = ? ", iProgramacao, objPV.lCodigo, objPV.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 178568
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178569
    
        If lErro <> AD_SQL_SUCESSO Then gError 178570
    
    End If

    objPV.iProgramacao = iProgramacao

    'Fecha o comando
    Call Comando_Fechar(lComando)

    PV_Le_Customizado = SUCESSO

    Exit Function

Erro_PV_Le_Customizado:

    PV_Le_Customizado = gErr

    Select Case gErr

        Case 178565
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178566, 178567
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", gErr, objPV.lCodigo)
        
        Case 178568, 178569
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDABAIXADOS", gErr, objPV.lCodigo)
        
        Case 178570
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPV.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178571)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
