VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPVInpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PrecoBase(ByVal objCT As Object) As Object
     Set PrecoBase = objCT.objUserControl.Controls("PrecoBase")
End Property

Public Property Get Programacao(ByVal objCT As Object) As Object
     Set Programacao = objCT.objUserControl.Controls("Programacao")
End Property

Public Function PV_CargaPosFormLoad(ByVal objCT As Object, bTrazendoDoc As Boolean) As Long
    PrecoBase(objCT).Format = gobjFAT.sFormatoPrecoUnitario
End Function

Public Function PV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Qtd Pedida")
    objGridInt.colColuna.Add ("Qtd Canc.")
    objGridInt.colColuna.Add ("Preço Base")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")
    objGridInt.colColuna.Add ("Prioridade")
    
    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colColuna.Add ("% a Menos Fat")
        objGridInt.colColuna.Add ("% a Mais Fat")
        objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
        objGridInt.colColuna.Add ("Quant Fat. A Mais")
    End If

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.QuantCancelada.Name)
    objGridInt.colCampo.Add (PrecoBase(objCT).Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
    objGridInt.colCampo.Add (objCT.QuantReservadaPV.Name)
    objGridInt.colCampo.Add (objCT.QuantFaturada.Name)
    objGridInt.colCampo.Add (objCT.Prioridade.Name)

    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colCampo.Add (objCT.PercentMenosReceb.Name)
        objGridInt.colCampo.Add (objCT.PercentMaisReceb.Name)
        objGridInt.colCampo.Add (objCT.RecebForaFaixa.Name)
        objGridInt.colCampo.Add (objCT.QuantFatAMais.Name)
    Else
        objCT.PercentMenosReceb.left = POSICAO_FORA_TELA
        objCT.PercentMaisReceb.left = POSICAO_FORA_TELA
        objCT.RecebForaFaixa.left = POSICAO_FORA_TELA
        objCT.QuantFatAMais.left = POSICAO_FORA_TELA
    End If
    
    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_QuantPedida_Col = 4
    objCT.iGrid_QuantCancel_Col = 5
    objCT.gobjInfoUsu.iGrid_PrecoBase_Col = 6
    objCT.iGrid_ValorUnitario_Col = 7
    objCT.iGrid_PercDesc_Col = 8
    objCT.iGrid_Desconto_Col = 9
    objCT.iGrid_PrecoTotal_Col = 10
    objCT.iGrid_PrecoTotalB_Col = 11
    objCT.iGrid_DataEntrega_Col = 12
    objCT.iGrid_QuantReservada_Col = 13
    objCT.iGrid_QuantFaturada_Col = 14
    objCT.iGrid_Prioridade_Col = 15
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        objCT.iGrid_PercentMenosReceb_Col = 16
        objCT.iGrid_PercentMaisReceb_Col = 17
        objCT.iGrid_RecebForaFaixa_Col = 18
        objCT.iGrid_QuantFatAMais_Col = 19

    End If

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Function PV_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Traz_Pedido_Tela

    If objPedidoVenda.iProgramacao = DESMARCADO Then
        Programacao(objCT).Value = vbUnchecked
    Else
        Programacao(objCT).Value = vbChecked
    End If
       
    PV_Traz_Pedido_Tela = SUCESSO
     
    Exit Function
    
Erro_PV_Traz_Pedido_Tela:

    PV_Traz_Pedido_Tela = gErr
     
    Select Case gErr
          
        Case 124189
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205060)
     
    End Select
     
    Exit Function
    
End Function

Function PV_Carrega_Grid_Itens(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal bProcessaReserva As Boolean) As Long

Dim lErro As Long, iIndice As Integer, objItemPedido As ClassItemPedido, sTipo As String

On Error GoTo Erro_PV_Carrega_Grid_Itens

    For iIndice = 1 To objPedidoVenda.colItensPedido.Count
       
        Set objItemPedido = objPedidoVenda.colItensPedido(iIndice)
        
        objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoBase, gobjFAT.sFormatoPrecoUnitario)

    Next
    
    PV_Carrega_Grid_Itens = SUCESSO
     
    Exit Function
    
Erro_PV_Carrega_Grid_Itens:

    PV_Carrega_Grid_Itens = gErr
     
    Select Case gErr
          
        Case 124182 To 124184
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205061)
     
    End Select
     
    Exit Function

End Function

Public Sub Programacao_Click(ByVal objCT As Object)
        
Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantFaturada As Double

On Error GoTo Erro_Programacao_Click

    'Se marcou a Programação
    If Programacao(objCT).Value = vbChecked Then
                
'        'Zera as quantidades Reservadas do Grid de Alocações
'        For iIndice = objCT.objGridAlocacao.iLinhasExistentes To 1 Step -1
'            lErro = Grid_Exclui_Linha(objCT.objGridAlocacao, iIndice)
'        Next

        Call Grid_Limpa(objCT.objGridAlocacao)
        
        objCT.ProdutoDescricao.Caption = ""
            
        'Zera as quantidades Reservadas no Grid de Itens
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_QuantReservada_Col) = ""
        Next
        
    'Se desmarcou
    Else
           
        'Para cada Item do GridItens
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
        
            'Reprocessa a reserva do produto
            dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_QuantPedida_Col))
            dQuantCancelada = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_QuantCancel_Col))
            dQuantFaturada = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_QuantFaturada_Col))
    
            lErro = objCT.Reserva_Processa(iIndice, dQuantidade, dQuantCancelada, dQuantFaturada)
            If lErro <> SUCESSO Then gError 75485
        
        Next
        
    End If
        
    Exit Sub
    
Erro_Programacao_Click:

    Select Case gErr
        
        Case 75485
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205062)
    
    End Select
    
    Exit Sub
    
End Sub

Function PV_Cliente_Validate(ByVal objCT As Object, ByVal lCliente As Long) As Long

Dim lErro As Long, dSaldoAtraso As Double
Dim objClienteInpal As ClassClienteInpal
Dim objCliente As New ClassCliente
Dim dTaxaCliente As Double, dPercAcresFin As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_PV_Cliente_Validate

    objCliente.lCodigo = lCliente

    lErro = CF("Cliente_Le_Customizado", objCliente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192978
    
    If Not (objCliente.objInfoUsu Is Nothing) Then
                
        Set objClienteInpal = objCliente.objInfoUsu
                
        'Se o cliente possui taxa padrão ou o cliente não está preenchido
        If objClienteInpal.iPadraoTaxaFin = Padrao Then
            
            'Calcula a taxa Mensal a partir da taxa diária utilizando juros compostos
            dTaxaCliente = (((1 + gobjCRFAT.dTaxaFinanceira) ^ 30) - 1)
        Else
            dTaxaCliente = objClienteInpal.dTaxaFinanceira
        End If
        
        'Altera PercAcrescFin
        If Len(Trim(objCT.PercAcrescFin.clipText)) > 0 Then
    
            dPercAcresFin = StrParaDbl(objCT.PercAcrescFin.Text) / 100
            If dPercAcresFin <> Round(dTaxaCliente, 2) Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
                If vbMsgRes = vbYes Then
                    objCT.PercAcrescFin.Text = Format(dTaxaCliente * 100, "Fixed")
                    objCT.gobjInfoUsu.giPercAcresFinAlterado = 0
                End If
            End If
        Else
            objCT.PercAcrescFin.Text = Format(dTaxaCliente * 100, "Fixed")
            objCT.gobjInfoUsu.giPercAcresFinAlterado = 0
        End If
        
    End If
        
    PV_Cliente_Validate = SUCESSO
     
    Exit Function
    
Erro_PV_Cliente_Validate:

    PV_Cliente_Validate = gErr
     
    Select Case gErr
          
        Case 132307
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205063)
     
    End Select
     
    Exit Function

End Function

Function PV_Cliente_Validate2(ByVal objCT As Object, ByVal lCliente As Long) As Long

Dim lErro As Long, iIndice As Integer
Dim objClienteInpal As ClassClienteInpal
Dim objCliente As New ClassCliente

On Error GoTo Erro_PV_Cliente_Validate2

    'Para cada linha do Grid de Itens
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
    
        '4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
        '5 - Marcio - 08/2000 - incluido colocar default para Preço Unitário e descrição do Produto a partir do último Pedido cadastrado p/ o cliente
        'Preenche a linha do Grid com dados da último Pedido feito pelo cliente em questão
        lErro = Trata_Cliente_UltimoPedido(objCT, iIndice, lCliente)
        If lErro <> SUCESSO Then gError 500025
        
    Next
        
    PV_Cliente_Validate2 = SUCESSO
     
    Exit Function
    
Erro_PV_Cliente_Validate2:

    PV_Cliente_Validate2 = gErr
     
    Select Case gErr
    
        Case 500025
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205064)
     
    End Select
     
    Exit Function

End Function

Public Function PV_CondicaoPagamento_Click(ByVal objCT As Object, ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
    PV_CondicaoPagamento_Click = Recalcula_PrecoUnitario_Itens(objCT)
End Function

Function Recalcula_PrecoUnitario(ByVal objCT As Object, iLinhaGrid As Integer) As Long

Dim dPrecoBase As Double
Dim dPrecoUnitario As Double
Dim dDias As Double
Dim dTaxaCliente As Double
Dim dTaxaFinDia As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim lErro As Long

On Error GoTo Erro_Recalcula_PrecoUnitario

    'precobase * ( 1 + ((1 + ((1 + taxamensal/100) ^ (1 / 30)) - 1) ^ dias) - 1)
    
    'Se o acréscimo financeiro é igual a zero
    If StrParaDbl(objCT.PercAcrescFin.Text) = 0 Or Len(Trim(objCT.CondicaoPagamento.Text)) = 0 Then
        
        'Preço unitário = preço base
        objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_ValorUnitario_Col) = objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)
    
    'Se é maior que zero
    Else
    
        'Lê Condição a partir do código
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCT.CondicaoPagamento.Text)
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 500028
        
        'Se a condição de pagamento não está cadastrada, erro
        If lErro = 19205 Then gError 500029
        
        'Calcula a média ponderada de dias para a taxa financeira
        Call CalculaDias_CondPagamento(objCondicaoPagto, dDias)
        
        'Calcula a taxa financeira diária
        Call Calcula_TaxaFinanceira(dDias, StrParaDbl(objCT.PercAcrescFin.Text) / 100, dTaxaFinDia)
                            
        'Se o preço Base estiver preenchido
        If Len(Trim(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))) > 0 Then
        
            'Recalcula o Preço Unitário a partir do Preço Base
            dPrecoBase = StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))
            dPrecoUnitario = dPrecoBase * (1 + dTaxaFinDia)
        
            objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            
        End If
            
    End If
    
    'Recalcula preço total
    Call objCT.PrecoTotal_Calcula(iLinhaGrid)

    Recalcula_PrecoUnitario = SUCESSO
    
    Exit Function
    
Erro_Recalcula_PrecoUnitario:

    Recalcula_PrecoUnitario = gErr
    
    Select Case gErr
    
        Case 500028
        
        Case 500029
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205065)
    
    End Select
    
    Exit Function
    
End Function

Sub CalculaDias_CondPagamento(objCondPagto As ClassCondicaoPagto, dDias As Double)
    
Dim iDiaParcelaAtual As Integer
Dim iTotalDias As Integer
Dim iNumeroParcelas As Integer
Dim iPeso As Integer
Dim iIntervalo As Integer

    'Se a condição de pagamento for à vista
    If objCondPagto.iCodigo = COD_A_VISTA Or objCondPagto.iCodigo = 0 Then
        
        dDias = 0
    
    'Se não for à vista
    Else
        
        'Se a condição de pagamento for Mensal
        If objCondPagto.iMensal = 1 Then
            iIntervalo = 30
        Else
            iIntervalo = objCondPagto.iIntervaloParcelas
        End If
        
        'Guarda o número de parcelas tirando a primeira
        iNumeroParcelas = objCondPagto.iNumeroParcelas - 1
                    
        'Se possui mais de duas parcelas, calcula a média
        If iNumeroParcelas > 1 Then
        
            'Calcula total de dias da condição de pagamento
            Do While iNumeroParcelas >= 0
                
                'Calcula o número de dias que faltam para chegar a parcela em questão
                iDiaParcelaAtual = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo * iNumeroParcelas)
                
                'Acumula o número de dias de todas as parcelas
                iTotalDias = iTotalDias + iDiaParcelaAtual
                
                'Decrementa o número de parcelas
                iNumeroParcelas = iNumeroParcelas - 1
                                        
            Loop
            
            'Calcula a média ponderada de dias
            dDias = iTotalDias / objCondPagto.iNumeroParcelas
        
        'Se possui exatamente duas parcelas
        ElseIf iNumeroParcelas = 1 Then
            
            dDias = objCondPagto.iDiasParaPrimeiraParcela + (iIntervalo / 2)
        
        'Se só possui exatamente uma parcela
        Else
            
            dDias = objCondPagto.iDiasParaPrimeiraParcela
            
        End If
        
    End If
    
End Sub

Sub Calcula_TaxaFinanceira(dDias As Double, dTaxaFin As Double, dTaxaFinDia As Double)
'Calcula taxa financeira diária a partir dos número de dias e da taxa financeira mensal passada
        
Dim dTaxaUmDia As Double

    'Calcula os juros compostos ao dia a partir da taxa financeira (ao mês) passada
    'TaxaFin ao dia = (1 + TaxaFin ao mês) elevado a (1 / 30 dias) -1
    dTaxaUmDia = ((1 + dTaxaFin) ^ (0.0333)) - 1
    
    'Calcula a taxa financeira para iDias a partir da taxa financeira ao dia
    dTaxaFinDia = ((1 + dTaxaUmDia) ^ dDias) - 1
            
End Sub

Public Sub PrecoBase_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

    '*** 11/04/02 - Luiz G.F.Nogueira ***
    'Seta iComissoesAlterada
    objCT.iComissoesAlterada = REGISTRO_ALTERADO
    '************************************

End Sub

Public Sub PrecoBase_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub PrecoBase_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub PrecoBase_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = PrecoBase(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function PV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.gobjInfoUsu.iGrid_PrecoBase_Col
            lErro = Saida_Celula_PrecoBase(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 124164

    End Select

    PV_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_PV_Saida_Celula_GridItens:

    PV_Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 124164

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205066)

    End Select

    Exit Function

End Function

Public Function PV_Saida_Celula_UM(ByVal objCT As Object, ByVal sUmAnterior As String, bIgnora As Boolean) As Long
    If sUmAnterior <> objCT.UnidadeMed.Text And Programacao(objCT).Value = vbUnchecked Then bIgnora = True
    PV_Saida_Celula_UM = SUCESSO
End Function

Public Function PV_Saida_Celula_Quantidade(ByVal objCT As Object, ByVal bQuantidadeIgual As Boolean, bIgnora As Boolean) As Long
    If Not bQuantidadeIgual And Programacao(objCT).Value = vbUnchecked Then bIgnora = True
    PV_Saida_Celula_Quantidade = SUCESSO
End Function

Public Function PV_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Rotina_Grid_Enable

    Select Case objControl.Name
        
        Case PrecoBase(objCT).Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
    
    End Select
    
    PV_Rotina_Grid_Enable = SUCESSO
     
    Exit Function
    
Erro_PV_Rotina_Grid_Enable:

    PV_Rotina_Grid_Enable = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205067)
     
    End Select
     
    Exit Function

End Function

Public Function PV_Produto_Saida_Celula6(ByVal objCT As Object) As Long
    PV_Produto_Saida_Celula6 = Trata_Cliente_UltimoPedido(objCT, objCT.GridItens.Row)
End Function

Public Function PV_Traz_Produto_Tela2(ByVal objCT As Object) As Long
    PV_Traz_Produto_Tela2 = Trata_Cliente_UltimoPedido(objCT, objCT.GridItens.Row)
End Function

Private Function Trata_Cliente_UltimoPedido(ByVal objCT As Object, iLinhaGrid As Integer, Optional lCodCli As Long = 0) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objItemPedido As New ClassItemPedido
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iFilial As Integer

On Error GoTo Erro_Trata_Cliente_UltimoPedido

    'Se o cliente foi preenchido
    If Len(Trim(objCT.Cliente.clipText)) > 0 And Len(Trim(objCT.Filial.Text)) > 0 Then 'Trecho incluido por Leo em 04/03/02
        
        If lCodCli = 0 Then
        
            'Lê o Cliente
            objCliente.sNomeReduzido = objCT.Cliente.Text
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 500021
            
            'Se o cliente não está cadastrado, erro
            If lErro = 12348 Then gError 500022
        
            lCodCli = objCliente.lCodigo
                           
            iFilial = Codigo_Extrai(objCT.Filial.Text)
                            
        End If
        
        'Formata o produto
        lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 500026
        
        'Procura por último Pedido feito pelo cliente com o mesmo Item
        objItemPedido.sProduto = sProdutoFormatado
        objItemPedido.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("ClienteFilial_Le_UltimoItemPedido", objItemPedido, lCodCli, iFilial)
        If lErro <> SUCESSO And lErro <> 94412 Then gError 500015 'Alterado por Leo em 05/03/02

        'Se encontrou o Item
        If lErro = SUCESSO Then 'alterado por Leo em 05/03/02
        
            objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = Format(IIf(objItemPedido.dPrecoBase <> 0, objItemPedido.dPrecoBase, objItemPedido.dPrecoUnitario), gobjFAT.sFormatoPrecoUnitario)
            
            '??? desativado para a inpal sc em 27/06/2016
            If glEmpresa <> 14 Then objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_DescProduto_Col) = objItemPedido.sDescricao
                                
            '4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
            'Calcula o preço unitário
            lErro = Recalcula_PrecoUnitario(objCT, iLinhaGrid)
            If lErro <> SUCESSO Then gError 500032
                
        End If
    
    End If
    
    Trata_Cliente_UltimoPedido = SUCESSO
    
    Exit Function
    
Erro_Trata_Cliente_UltimoPedido:
    
    Trata_Cliente_UltimoPedido = gErr
    
    Select Case gErr
    
        Case 500015, 500021, 500026, 500032
        
        Case 500022
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205068)
    
    End Select

    Exit Function
    
End Function

Public Function PV_Saida_Celula_PrecoUnitario3(ByVal objCT As Object, ByVal bPrecoUnitarioIgual As Boolean) As Long
    'Se mudou o preço unitário mudou ou o Preço base da linha atual não está preenchido
    If (Not bPrecoUnitarioIgual Or StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) = 0) And Len(Trim(objCT.PrecoUnitario.Text)) > 0 Then
        'Recalcula Preço Base
        PV_Saida_Celula_PrecoUnitario3 = PrecoBase_Recalcula(objCT, objCT.GridItens.Row)
    Else
        PV_Saida_Celula_PrecoUnitario3 = SUCESSO
    End If
End Function

Function PrecoBase_Recalcula(ByVal objCT As Object, iLinhaGrid As Integer) As Long
'Recalcula o preço base a partir do preço unitário e do acréscimo financeiro

Dim lErro As Long
Dim dTaxaCliente As Double
Dim dTaxaFinDia As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dPrecoUnitario As Double
Dim dPrecoBase As Double
Dim dDias As Double

On Error GoTo Erro_PrecoBase_Recalcula

    'Se o acréscimo financeiro não foi preenchido ou é igual a zero
    If StrParaDbl(objCT.PercAcrescFin.clipText) = 0 Or Len(Trim(objCT.CondicaoPagamento.Text)) = 0 Then
        
        'Preço Base = Preço Unitário
        objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_ValorUnitario_Col)
    
    'Se o acréscimo financeiro foi preenchido
    Else
            
        'Lê Condição a partir do código
        objCondicaoPagto.iCodigo = Codigo_Extrai(objCT.CondicaoPagamento.Text)
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 500014
        
        'Se a condição de pagamento não está cadastrada, erro
        If lErro = 19205 Then gError 500015
        
        'Calcula a média ponderada de dias para a taxa financeira
        Call CalculaDias_CondPagamento(objCondicaoPagto, dDias)
        
        'Calcula a taxa financeira diária
        Call Calcula_TaxaFinanceira(dDias, StrParaDbl(objCT.PercAcrescFin.Text) / 100, dTaxaFinDia)
                                                        
        dPrecoUnitario = StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_ValorUnitario_Col))
                                
        'Calcula o preço base
        dPrecoBase = dPrecoUnitario / (1 + dTaxaFinDia)
                
        'Coloca o preço Base na tela
        If dPrecoBase > 0 Then
            objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = Format(dPrecoBase, gobjFAT.sFormatoPrecoUnitario)
        Else
            objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = ""
        End If
                
    End If
    
    PrecoBase_Recalcula = SUCESSO
    
    Exit Function
    
Erro_PrecoBase_Recalcula:

    PrecoBase_Recalcula = gErr
    
    Select Case gErr
    
        Case 500014
        
        Case 500015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205069)
    
    End Select
    
    Exit Function
    
End Function

'4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
Function Saida_Celula_PrecoBase(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Base que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoBaseIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoBase

    bPrecoBaseIgual = False

    Set objGridInt.objControle = PrecoBase(objCT)

    If Len(Trim(PrecoBase(objCT).Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoBase(objCT).Text)
        If lErro <> SUCESSO Then gError 500007

    End If

    'Comparação com Preço Base anterior
    If StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) = StrParaDbl(PrecoBase(objCT).Text) Then bPrecoBaseIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 500008

    'Se mudou o preço Base ou o Preço unitário da linha atual não está preenchido
    If (Not bPrecoBaseIgual Or StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_ValorUnitario_Col)) = 0) And Len(Trim(PrecoBase(objCT).Text)) > 0 Then
                        
        'Reclacula o preço unitário
        lErro = Recalcula_PrecoUnitario(objCT, objCT.GridItens.Row)
        If lErro <> SUCESSO Then gError 500016
        
        lErro = objCT.ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 500009

    End If

    Saida_Celula_PrecoBase = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoBase:

    Saida_Celula_PrecoBase = gErr

    Select Case gErr

        Case 500007, 500008, 500009, 500010, 500016
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205070)

    End Select

    Exit Function

End Function

Public Function PV_BotaoReserva_Click(ByVal objCT As Object, ByVal objItemPedido As ClassItemPedido) As Long
    PV_BotaoReserva_Click = Trata_Cliente_UltimoPedido(objCT, objItemPedido.iItem)
End Function

Public Function PV_Move_Pedido_Memoria(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Move_Pedido_Memoria
    
    If Programacao(objCT).Value = vbUnchecked Then
        objPedidoVenda.iProgramacao = DESMARCADO
    Else
        objPedidoVenda.iProgramacao = MARCADO
    End If
    
    PV_Move_Pedido_Memoria = SUCESSO
     
    Exit Function
    
Erro_PV_Move_Pedido_Memoria:

    PV_Move_Pedido_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205071)
     
    End Select
     
    Exit Function

End Function

Function PV_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPedido As ClassItemPedido, ByVal iIndice As Integer, ByVal sProduto1 As String) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Move_GridItem_Memoria

    objItemPedido.dPrecoBase = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))
    
    PV_Move_GridItem_Memoria = SUCESSO
     
    Exit Function
    
Erro_PV_Move_GridItem_Memoria:

    PV_Move_GridItem_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205072)
     
    End Select
     
    Exit Function

End Function

Function PV_Valida_Grid_Itens(ByVal objCT As Object, ByVal iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Valida_Grid_Itens

    lErro = Valor_Positivo_Critica(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))
    If lErro <> SUCESSO Then gError 500036
        
    'Se é um Pedido Programado e a quantidade Reservada foi preenchida, erro
    If Programacao(objCT).Value = vbChecked And StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_QuantReservada_Col)) > 0 Then gError 75486
    
    PV_Valida_Grid_Itens = SUCESSO
     
    Exit Function
    
Erro_PV_Valida_Grid_Itens:

    PV_Valida_Grid_Itens = gErr
     
    Select Case gErr
          
        Case 75486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_PROGRAMADO_COM_RESERVA", gErr)
                
        Case 500036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_ITEM_NAO_PREENCHIDO", gErr, iIndice)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205073)
     
    End Select
     
    Exit Function

End Function

Public Sub PercAcrescFin_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PercAcrescFin_Validate

    If objCT.giPercAcresFinAlterado = 0 Then Exit Sub

    If Len(Trim(objCT.PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(objCT.PercAcrescFin.Text)
        If lErro <> SUCESSO Then gError 26717
    End If
        
    lErro = Recalcula_PrecoUnitario_Itens(objCT)
    If lErro <> SUCESSO Then gError 500031
    
    objCT.giPercAcresFinAlterado = 0

    Exit Sub

Erro_PercAcrescFin_Validate:

    Cancel = True

    Select Case gErr

        Case 26717, 46190, 500031, 500044

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205074)

    End Select

    Exit Sub

End Sub

Public Function PV_Limpa_Tela(ByVal objCT As Object) As Long
    Programacao(objCT).Value = vbUnchecked
    PV_Limpa_Tela = SUCESSO
End Function

Private Function Recalcula_PrecoUnitario_Itens(ByVal objCT As Object) As Long

Dim lErro As Long, iIndice As Integer, bRecalcularValorTotal As Boolean
        
On Error GoTo Erro_Recalcula_PrecoUnitario_Itens

    bRecalcularValorTotal = False

    'Para cada linha do GridItens
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes

        'Recalcula os preços unitários do GridItens
        lErro = Recalcula_PrecoUnitario(objCT, iIndice)
        If lErro <> SUCESSO Then gError 500035

        If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_ValorUnitario_Col)) <> 0 Then bRecalcularValorTotal = True
        
    Next
    
    If bRecalcularValorTotal Then
        
        'Recalcula ValorTotal
        lErro = objCT.ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 500044
        
    End If

    Recalcula_PrecoUnitario_Itens = SUCESSO
    
    Exit Function
     
Erro_Recalcula_PrecoUnitario_Itens:

    Recalcula_PrecoUnitario_Itens = gErr
    
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205075)
     
    End Select
     
    Exit Function

End Function
