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

Public Function BloqueiosPV_Cust_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
Dim lErro As Long

On Error GoTo Erro_BloqueiosPV_Cust_Adiciona

    lErro = BloqueiosPV_CargoCondPagto_Adiciona(objPedidoVenda)
    If lErro <> SUCESSO Then gError 201181
    
    lErro = BloqueiosPV_CargoTabPreco_Adiciona(objPedidoVenda)
    If lErro <> SUCESSO Then gError 201182
    
    BloqueiosPV_Cust_Adiciona = SUCESSO
    
    Exit Function
    
Erro_BloqueiosPV_Cust_Adiciona:

    BloqueiosPV_Cust_Adiciona = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201170)

    End Select
    
    Exit Function

End Function

Public Function PV_CalculaPrecoUnitario(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iCodTabelaPreco As Integer, ByVal dtDataPV As Date, dPrecoUnitario As Double) As Long
Dim lErro As Long, lComando As Long, dPreco As Double

On Error GoTo Erro_PV_CalculaPrecoUnitario

    If dPrecoUnitario = 0 Then
    
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 201172
        
        lErro = Comando_Executar(lComando, "SELECT Preco FROM TabelasDePrecoItens WHERE CodTabela = ? AND FilialEmpresa = ? AND CodProduto = ? AND DataVigencia <= ? ORDER BY DataVigencia DESC", _
            dPreco, iCodTabelaPreco, iFilialEmpresa, sProduto, dtDataPV)
        If lErro <> AD_SQL_SUCESSO Then gError 201173
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201174
        If lErro <> AD_SQL_SUCESSO Then gError 201175
        
        dPrecoUnitario = dPreco
        
        Call Comando_Fechar(lComando)
    
    End If
    
    PV_CalculaPrecoUnitario = SUCESSO
    
    Exit Function
    
Erro_PV_CalculaPrecoUnitario:

    PV_CalculaPrecoUnitario = gErr

    Select Case gErr

        Case 201172
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201171)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function BloqueiosPV_CargoCondPagto_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long, iCargoMinimo As Integer
Dim objVendedor As New ClassVendedor, lComando As Long

On Error GoTo Erro_BloqueiosPV_CargoCondPagto_Adiciona

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 201177
    
    If objPedidoVenda.iCondicaoPagto <> 0 And objPedidoVenda.colComissoes.Count <> 0 Then
    
        'verificar se deve gerar bloqueio por condicao de pagto
        lErro = Comando_Executar(lComando, "SELECT CargoMinimo FROM CondicoesPagto WHERE Codigo = ?", iCargoMinimo, objPedidoVenda.iCondicaoPagto)
        If lErro <> AD_SQL_SUCESSO Then gError 201178
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201179
    
        If lErro = AD_SQL_SUCESSO And iCargoMinimo <> 0 Then
        
            objVendedor.iCodigo = objPedidoVenda.colComissoes(1).iCodVendedor
            lErro = CF("Vendedor_Le", objVendedor)
            If lErro <> SUCESSO And lErro <> 12582 Then gError 201180
            If lErro = SUCESSO Then
            
                If objVendedor.iCargo < iCargoMinimo Then
                
                    'Adiciona o bloqueio ao Pedido
                    objPedidoVenda.colBloqueiosPV.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_CARGO_COND_PAGTO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""

                End If
            
            End If
                
        End If
        
    End If
    
    Call Comando_Fechar(lComando)
    
    BloqueiosPV_CargoCondPagto_Adiciona = SUCESSO
    
    Exit Function
    
Erro_BloqueiosPV_CargoCondPagto_Adiciona:

    BloqueiosPV_CargoCondPagto_Adiciona = gErr

    Select Case gErr

        Case 201177
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201178, 201179
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARGO_CONDPAGTO", gErr)
        
        Case 201180
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201176)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function BloqueiosPV_CargoTabPreco_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long, iCargoMinimo As Integer
Dim objVendedor As New ClassVendedor, lComando As Long
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_BloqueiosPV_CargoTabPreco_Adiciona

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 201184
    
    If objPedidoVenda.colComissoes.Count <> 0 Then
    
        objVendedor.iCodigo = objPedidoVenda.colComissoes(1).iCodVendedor
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 201187
    
        If objVendedor.iCargo <> 0 Then
        
            For Each objItemPV In objPedidoVenda.colItensPedido
            
                If objItemPV.iTabelaPreco <> 0 Then
                
                    lErro = Comando_Executar(lComando, "SELECT CargoMinimo FROM TabelasDePreco WHERE Codigo = ?", iCargoMinimo, objItemPV.iTabelaPreco)
                    If lErro <> AD_SQL_SUCESSO Then gError 201185
                    
                    lErro = Comando_BuscarProximo(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201186
                
                    If lErro = AD_SQL_SUCESSO And iCargoMinimo <> 0 Then
                    
                        If objVendedor.iCargo < iCargoMinimo Then
                        
                            'Adiciona o bloqueio ao Pedido
                            objPedidoVenda.colBloqueiosPV.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_CARGO_TAB_PRECO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""
                            Exit For
        
                        End If
                    
                    End If
                    
                End If
            
            Next
    
        End If
    
    End If
    
    Call Comando_Fechar(lComando)
    
    BloqueiosPV_CargoTabPreco_Adiciona = SUCESSO
    
    Exit Function
    
Erro_BloqueiosPV_CargoTabPreco_Adiciona:

    BloqueiosPV_CargoTabPreco_Adiciona = gErr

    Select Case gErr

        Case 201184
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201185, 201186
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARGO_CONDPAGTO", gErr)
        
        Case 201187
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201183)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

