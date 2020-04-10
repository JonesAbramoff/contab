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

Function PedidoDeVenda_Fatura_Customizado(ByVal objNFiscal As ClassNFiscal, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
Dim objPVAla As New ClassPedidoDeVendaAla
    If objNFiscal.dtDataSaida <> DATA_NULA Then
        objPVAla.dtDataEntregaReal = objNFiscal.dtDataSaida
    Else
        objPVAla.dtDataEntregaReal = objNFiscal.dtDataEmissao
    End If
    Set objPedidoVenda.objInfoUsu = objPVAla
    PedidoDeVenda_Fatura_Customizado = PV_Grava_Customizado(objPedidoVenda)
End Function

Function PV_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim lComando As Long, sTabela As String
Dim iIndice As Integer, lNumIntDoc As Long
Dim dtDataEntrega As Date
Dim objPVAla As ClassPedidoDeVendaAla

On Error GoTo Erro_PV_Grava_Customizado

    If Not (objPV.objInfoUsu Is Nothing) Then
    
        Set objPVAla = objPV.objInfoUsu

        'Abre Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 140300
        Next
        
        sTabela = "PedidosDeVenda"
       
        'Tenta ler Pedido de Venda e Pedido de Venda Baixado
        For iIndice = 1 To 2
           
            'Le o Pedido de Venda
            lErro = Comando_ExecutarPos(alComando(iIndice), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc= ? ", 0, lNumIntDoc, objPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 140301
       
            lErro = Comando_BuscarPrimeiro(alComando(iIndice))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140302
            
            lComando = alComando(iIndice)
       
            If lErro = AD_SQL_SUCESSO Then Exit For
           
            sTabela = "PedidosDeVendaBaixados"
       
        Next
       
        'Não encontrou
        If lErro <> AD_SQL_SUCESSO Then gError 140303
       
        'Altera a DataEntregaReal
        'Ao acrescentar campos novos tem que alterar a função PedidoDeVenda_Fatura_Customizado
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE " & sTabela & " SET DataEntregaReal = ? ", lComando, objPVAla.dtDataEntregaReal)
        If lErro <> AD_SQL_SUCESSO Then gError 140304
        
        'Fecha comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
       
       
    End If
    
    PV_Grava_Customizado = SUCESSO

    Exit Function

Erro_PV_Grava_Customizado:

    PV_Grava_Customizado = gErr

        Select Case gErr

        Case 140300
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140301, 140302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 140303
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPV.lCodigo)
    
        Case 140304
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

