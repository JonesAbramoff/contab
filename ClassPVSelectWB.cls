VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVSelectWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PVItens_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long
'Lê os valores customizados dos ítens do Pedido de Venda passada em objNFiscal

Dim lErro As Long
Dim lComando As Long
Dim objItemPVWB As ClassItemPedidoWB
Dim objItemPV As ClassItemPedido
Dim dPrecoTabela As Double, dDescTabela As Double

On Error GoTo Erro_PVItens_Le_Customizado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141984

    For Each objItemPV In objPV.colItensPedido
        
        'Pesquisa os ítens vinculados ao pedido de venda
        lErro = Comando_Executar(lComando, "SELECT PrecoTabela, DescTabela FROM ItensPedidoDeVenda WHERE NumIntDoc = ? ", dPrecoTabela, dDescTabela, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141985

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141986

        If lErro <> AD_SQL_SUCESSO Then
        
            'Pesquisa os ítens vinculados ao pedido baixado
            lErro = Comando_Executar(lComando, "SELECT PrecoTabela, DescTabela FROM ItensPedidoDeVendaBaixados WHERE NumIntDoc = ? ", dPrecoTabela, dDescTabela, objItemPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 141987
    
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141988
        
        End If
                
        If lErro = AD_SQL_SUCESSO Then
            
            Set objItemPVWB = New ClassItemPedidoWB
            
            objItemPVWB.dPrecoTabela = dPrecoTabela
            objItemPVWB.dDescTabela = dDescTabela
            
            Set objItemPV.objInfoUsu = objItemPVWB

        End If

    Next

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    PVItens_Le_Customizado = SUCESSO

    Exit Function

Erro_PVItens_Le_Customizado:

    PVItens_Le_Customizado = gErr

    Select Case gErr

        Case 141984
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141985 To 141988
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141989)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

