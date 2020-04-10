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
'Alterações INPAL

'1 - 08/2000 Márcio - Função de gravação alterada para incluir preço base no ItemPV

Public Sub Arredonda_Insumos(ByVal sUnidadeMed As String, ByVal dQuantidadeOrig As Double, dQuantidadeArr As Double)
'Arredonda a quantidade de Insumos
Dim lErro As Long
Dim dFixQuant As Double
Dim dQuantTemp As Double

On Error GoTo Erro_Arredonda_Insumos

    Select Case sUnidadeMed
    
        Case "UN", "PC", "TB", "BD"
            dFixQuant = Fix(dQuantidadeOrig + 0.5 * Sgn(dQuantidadeOrig))
            dQuantTemp = dQuantidadeOrig - Int(dQuantidadeOrig)
            Select Case dQuantTemp
            
                Case 0
                    dQuantidadeArr = dFixQuant
                Case 0.5
                    dQuantidadeArr = dFixQuant
                Case Else
                    If (dQuantTemp < 0.5) Then
                        dFixQuant = dFixQuant + 1
                        dQuantidadeArr = dFixQuant
                    Else
                        dQuantidadeArr = dFixQuant
                    End If
            End Select
        
        Case Else
            If dQuantidadeOrig > 10 Then
                dQuantidadeArr = Round(dQuantidadeOrig, 0)
            Else
                dQuantidadeArr = Round(dQuantidadeOrig, 1)
            End If
    End Select

    Exit Sub
    
Erro_Arredonda_Insumos:

    lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
    
    Exit Sub
End Sub

Public Function PV_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sTabela As String
Dim alComando(0 To 2) As Long
Dim lComando As Long
Dim iProgramacao As Integer

On Error GoTo Erro_PV_Grava_Customizado

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178573
    Next
    
    iProgramacao = objPV.iProgramacao
        
    sTabela = "PedidosDeVenda"
    lComando = alComando(0)

    'Busca o pedido da tabela dos Pedidos de Venda Baixados
    lErro = CF("PedidoDeVenda_Le_Mod", objPV, alComando(0))
    If lErro <> SUCESSO And lErro <> 51357 Then gError 178574

    If lErro <> SUCESSO Then
    
        sTabela = "PedidosDeVendaBaixados"
        lComando = alComando(1)
    
        'Busca o pedido da tabela dos Pedidos de Venda Baixados
        lErro = CF("PedidoDeVendaBaixado_LeMod", alComando(1), objPV)
        If lErro <> SUCESSO And lErro <> 62205 Then gError 178575
        
    End If
        
    'Se não encontrou ---> erro
    If lErro <> SUCESSO Then gError 178576
        
    'Desfaz a baixa do pedido de venda
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE " & sTabela & " SET Programacao = ? ", lComando, iProgramacao)
    If lErro <> AD_SQL_SUCESSO Then gError 178578

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PV_Grava_Customizado = SUCESSO
    
    Exit Function
    
Erro_PV_Grava_Customizado:

    PV_Grava_Customizado = gErr
    
    Select Case gErr
    
        Case 178573
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178574, 178575, 178577
        
        Case 178576
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPV.lCodigo)

        Case 178578
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPV.lCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178579)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function
