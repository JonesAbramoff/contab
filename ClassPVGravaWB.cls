VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGravaWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ItensPV_Insere(ByVal lComando As Long, ByVal objItemPV As ClassItemPedido) As Long
'cria ordens de produção a partir de pedidos de venda

Dim lErro As Long
Dim objItemPVWB As ClassItemPedidoWB

On Error GoTo Erro_ItensPV_Insere

    Set objItemPVWB = objItemPV.objInfoUsu

    'Insere o item no BD
    lErro = Comando_Executar(lComando, "INSERT INTO ItensPedidoDeVenda (NumIntDoc,FilialEmpresa,CodPedido,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC, PrecoTabela, DescTabela) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objItemPV.lNumIntDoc, objItemPV.iFilialEmpresa, objItemPV.lCodPedido, objItemPV.sProduto, objItemPV.dQuantidade, objItemPV.dPrecoUnitario, objItemPV.dPrecoTotal, objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.dValorDesconto, objItemPV.dtDataEntrega, objItemPV.sDescricao, objItemPV.sLote, objItemPV.iStatus, objItemPV.dValorAbatComissao, objItemPV.dQuantCancelada, objItemPV.dQuantReservada, objItemPV.dQuantFaturada, objItemPV.dQuantOP, objItemPV.dQuantSC, objItemPVWB.dPrecoTabela, objItemPVWB.dDescTabela)
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

Function ItensPVBaixados_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda, iFaturaTudo As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItemPV As typeItemPedido
Dim iTemItens As Integer 'Sinaliza se o Pedido já vem carregado com seus itens
Dim dPrecoTabela As Double, dDescTabela As Double, objItemPedido As ClassItemPedido

On Error GoTo Erro_ItensPVBaixados_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 46104

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 46105

    tItemPV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPV.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemPV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC, PrecoTabela, DescTabela FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ?", 0, _
        tItemPV.lNumIntDoc, tItemPV.sProduto, tItemPV.dQuantidade, tItemPV.dPrecoUnitario, tItemPV.dPrecoTotal, tItemPV.iClasseUM, tItemPV.sUnidadeMed, tItemPV.dValorDesconto, tItemPV.dtDataEntrega, tItemPV.sDescricao, tItemPV.sLote, tItemPV.iStatus, tItemPV.dValorAbatComissao, tItemPV.dQuantCancelada, tItemPV.dQuantReservada, tItemPV.dQuantFaturada, tItemPV.dQuantOP, tItemPV.dQuantSC, dPrecoTabela, dDescTabela, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 46106

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46107

    iTemItens = False
    If objPedidoVenda.colItensPedido.Count > 0 Then iTemItens = True

    Do While lErro = AD_SQL_SUCESSO

        With tItemPV
            
            If Not iTemItens Then
                Set objItemPedido = objPedidoVenda.colItensPedido.Add(objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantFaturada, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, .lNumIntDoc)
                Set objItemPedido.objInfoUsu = New ClassItemPedidoWB
                objItemPedido.objInfoUsu.dPrecoTabela = dPrecoTabela
                objItemPedido.objInfoUsu.dDescTabela = dDescTabela
            End If
            
            .lCodPedido = objPedidoVenda.lCodigo
            .iFilialEmpresa = objPedidoVenda.iFilialEmpresa
            .dQuantReservada = 0

            lErro = Comando_Executar(lComando1, "INSERT INTO ItensPedidoDeVendaBaixados (FilialEmpresa, CodPedido, NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC, PrecoTabela, DescTabela) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .lCodPedido, .lNumIntDoc, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, dPrecoTabela, dDescTabela)
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

Function PVItens_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim objItemPVWB As ClassItemPedidoWB
Dim objItemPV As ClassItemPedido
Dim lNumIntDoc As Long
Dim sTabela As String

On Error GoTo Erro_PVItens_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141994
    Next

    For Each objItemPV In objPV.colItensPedido
        
        sTabela = "ItensPedidoDeVenda"
        
        'Pesquisa os ítens vinculados à Nota Fiscal
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141995

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141996

        If lErro <> AD_SQL_SUCESSO Then

            sTabela = "ItensPedidoDeVendaBaixados"

            'Pesquisa os ítens vinculados à Nota Fiscal
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 141997
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141998

        End If

        If lErro = AD_SQL_SUCESSO Then
        
            Set objItemPVWB = objItemPV.objInfoUsu
            
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE " & sTabela & " SET PrecoTabela = ?, DescTabela = ?", alComando(0), objItemPVWB.dPrecoTabela, objItemPVWB.dDescTabela)
            If lErro <> AD_SQL_SUCESSO Then gError 141999

        End If

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PVItens_Grava_Customizado = SUCESSO

    Exit Function

Erro_PVItens_Grava_Customizado:

    PVItens_Grava_Customizado = gErr

    Select Case gErr

        Case 141994
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141995 To 141998
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 141999
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177000)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

