VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFatPedHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#######################################################
'Inserido por Wagner
'Public Property Get PercDesconto(ByVal objCT As Object) As Object
'     Set PercDesconto = objCT.objUserControl.Controls("PercDesconto")
'End Property
'#######################################################

'--- inicio dos properties get dos controles da tela
Public Property Get Moeda(ByVal objCT As Object) As Object
     Set Moeda = objCT.objUserControl.Controls("Moeda")
End Property

Public Property Get PrecoUnitarioMoeda(ByVal objCT As Object) As Object
     Set PrecoUnitarioMoeda = objCT.objUserControl.Controls("PrecoUnitarioMoeda")
End Property

Public Property Get TaxaConversao(ByVal objCT As Object) As Object
     Set TaxaConversao = objCT.objUserControl.Controls("TaxaConversao")
End Property

Public Property Get ValorFrete1(ByVal objCT As Object) As Object
     Set ValorFrete1 = objCT.objUserControl.Controls("ValorFrete1")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Moeda_Change(ByVal objCT As Object)
    
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Moeda_Click(ByVal objCT As Object)
    If Not objCT.gbCarregandoTela Then
        'se trocar de dolar (ou outra moeda) p/real, colocar taxa =1 e tratar feito se tivesse trocado a taxa.
        If Len(Trim(TaxaConversao(objCT).Text)) <> 0 And Moeda(objCT).ListIndex = 0 Then
            TaxaConversao(objCT).Text = ""
            Call TaxaConversao_Validate(objCT, bSGECancelDummy)
        End If
    End If
End Sub

Public Function NFFP_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Pr. Unitário")
    objGridInt.colColuna.Add ("Unitário R$")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Data Entrega")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitarioMoeda(objCT).Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)

    'Colunas da Grid
    With objCT
        .iGrid_Produto_Col = 1
        .iGrid_DescProduto_Col = 2
        .iGrid_UnidadeMed_Col = 3
        .iGrid_Quantidade_Col = 4
        .gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col = 5
        .iGrid_ValorUnitario_Col = 6
        .iGrid_PercDesc_Col = 7
        .iGrid_Desconto_Col = 8
        .iGrid_PrecoTotal_Col = 9
        .iGrid_DataEntrega_Col = 10
    End With
    
    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS_HICARE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFFP_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Function NFFP_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    'liberar objetos globais referentes a customizacao
    
    NFFP_Form_Unload = SUCESSO
    
End Function

Function NFFP_CargaPosFormLoad(ByVal objCT As Object, ByVal bTrazendoDoc As Boolean) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_CargaPosFormLoad

    PrecoUnitarioMoeda(objCT).Format = gobjFAT.sFormatoPrecoUnitario
        
    'carregar moedas
    lErro = Carrega_Moeda(objCT)
    If lErro <> SUCESSO Then gError 124170
    
    NFFP_CargaPosFormLoad = SUCESSO
     
    Exit Function
    
Erro_NFFP_CargaPosFormLoad:

    NFFP_CargaPosFormLoad = gErr
     
    Select Case gErr
          
        Case 124170
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179287)
     
    End Select
     
    Exit Function

End Function

Private Function Carrega_Moeda(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda(objCT).AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda(objCT).ItemData(Moeda(objCT).NewIndex) = objMoeda.iCodigo
    
    Next

    Moeda(objCT).ListIndex = 0
    
    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179288)
    
    End Select

End Function

Public Sub BotaoTrazCotacao_Click(ByVal objCT As Object)
'Traz a última cotação da moeda selecionada

Dim lErro As Long
Dim iLinha As Integer
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_BotaoTrazCotacao_Click

    'Carrega objCotacao
    objCotacao.dtData = gdtDataAtual
    
    'Se a moeda não foi selecionada => Erro
    If Len(Trim(Moeda(objCT).Text)) = 0 Then gError 108943
        
    'Preeche com a Moeda selecionada
    objCotacao.iMoeda = Codigo_Extrai(Moeda(objCT).List(Moeda(objCT).ListIndex))
    objCotacaoAnterior.iMoeda = Codigo_Extrai(Moeda(objCT).List(Moeda(objCT).ListIndex))

    'Chama função de leitura
    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
    If lErro <> SUCESSO Then gError 108944
    
    'Se nao existe cotacao para a data informada => Mostra a última. Se mesmo assim nao existir => Colocar 1,00
    TaxaConversao(objCT).Text = IIf(objCotacaoAnterior.dValor = 0, 1, IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000")))
    
    Call TaxaConversao_Validate(objCT, bSGECancelDummy)
    
    Exit Sub
    
Erro_BotaoTrazCotacao_Click:

    Select Case gErr
    
        Case 108943
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_PREENCHIDA", gErr)
            
        Case 108944
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179289)
    
    End Select
    
End Sub

Public Sub TaxaConversao_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TaxaConversao_Validate(ByVal objCT As Object, Cancel As Boolean)
' recalcular precounitario em R$ e repercutir consequencias desta alteracao
    
Dim lErro As Long
Dim iLinha As Integer
Dim sProduto As String, dPrecoUnitario As Double
Dim iPreenchido As Integer, dTaxaConversao As Double

On Error GoTo Erro_TaxaConversao_Validate

    If Not objCT.gbCarregandoTela Then
    
        'Verifica se algum valor foi digitado
        If Len(Trim(TaxaConversao(objCT).Text)) > 0 Then
    
            'Critica se é valor Positivo
            lErro = Valor_Positivo_Critica_Double(TaxaConversao(objCT).Text)
            If lErro <> SUCESSO Then gError 103378
        
            'Põe o valor formatado na tela
            TaxaConversao(objCT).Text = Format(TaxaConversao(objCT).Text, FORMATO_TAXA_CONVERSAO_MOEDA)
            
        End If
        
        dTaxaConversao = StrParaDbl(TaxaConversao(objCT).Text)
        If dTaxaConversao = 0 Then dTaxaConversao = 1
        
        For iLinha = 1 To objCT.objGridItens.iLinhasExistentes
    
            'Verifica se o Produto está preenchido
            lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 46857
    
            If iPreenchido <> PRODUTO_VAZIO Then
            
                dPrecoUnitario = dTaxaConversao * StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col))
                If dPrecoUnitario > 0 Then
                    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                Else
                    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = ""
                End If
                
                'Calcula o Preco Total do item
                Call objCT.PrecoTotal_Calcula(iLinha, False)
    
                ''*** incluidos p/tratamento de tributacao *******************************
                lErro = objCT.gobjTribTab.Alteracao_Item_Grid(iLinha)
                If lErro <> SUCESSO Then gError 56862
                '*** fim tributacao
            
            End If
            
        Next
        
        'Calcula o Valor Total da Nota
        Call objCT.ValorTotal_Calcula
    
    End If
    
    Exit Sub
     
Erro_TaxaConversao_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case 46857, 56862, 103378
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179290)
     
    End Select
     
    Exit Sub

End Sub

Public Sub PrecoUnitarioMoeda_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitarioMoeda_GotFocus(ByVal objCT As Object)
    
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub PrecoUnitarioMoeda_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub PrecoUnitarioMoeda_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = PrecoUnitarioMoeda(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorFrete1_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorFrete1_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim dValorFrete As Double

On Error GoTo Erro_ValorFrete1_Validate

    'Verifica se o Valor foi preenchido
    If Len(Trim(ValorFrete1(objCT).Text)) > 0 Then
    
        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorFrete1(objCT).Text)
        If lErro <> SUCESSO Then gError 46815

        dValorFrete = CDbl(ValorFrete1(objCT).Text)
        
        'Coloca o Valor Formatado na tela
        ValorFrete1(objCT).Text = Format(dValorFrete, "Fixed")

    End If
    
    Exit Sub

Erro_ValorFrete1_Validate:

    Cancel = True

    Select Case gErr

        Case 46815

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179291)

    End Select

    Exit Sub

End Sub

Function NFFP_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
Dim lErro As Long

On Error GoTo Erro_NFFP_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
    
        Case PrecoUnitarioMoeda(objCT).Name
            If iProdutoPreenchido = PRODUTO_VAZIO Or objCT.gcolItemPedido.Count = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
    End Select
    
    NFFP_Rotina_Grid_Enable = SUCESSO
     
    Exit Function
    
Erro_NFFP_Rotina_Grid_Enable:

    NFFP_Rotina_Grid_Enable = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179292)
     
    End Select
     
    Exit Function

End Function

Function NFFP_Traz_NFiscal_Tela2(ByVal objCT As Object, objNFiscal As ClassNFiscal, lNumIntPedido As Long) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_Traz_NFiscal_Tela2

    If objNFiscal.lNumIntDoc <> 0 Then
    
        lErro = CF("NFiscal_Le_InfoUsu", objNFiscal)
        If lErro <> SUCESSO Then gError 124196
        
        If objNFiscal.objInfoUsu.dValorFrete1 > 0 Then ValorFrete1(objCT).Text = Format(objNFiscal.objInfoUsu.dValorFrete1, "Standard")
        
    End If
    
    Call Combo_Seleciona_ItemData(Moeda(objCT), objNFiscal.iMoeda)
    
    If objNFiscal.dTaxaMoeda <> 0 Then
        TaxaConversao(objCT).Text = Format(objNFiscal.dTaxaMoeda, FORMATO_TAXA_CONVERSAO_MOEDA)
    Else
        TaxaConversao(objCT).Text = ""
    End If
        
    NFFP_Traz_NFiscal_Tela2 = SUCESSO
     
    Exit Function
    
Erro_NFFP_Traz_NFiscal_Tela2:

    NFFP_Traz_NFiscal_Tela2 = gErr
     
    Select Case gErr
          
        Case 124196
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179293)
     
    End Select
     
    Exit Function

End Function

Function NFFP_Preenche_GridItem(ByVal objCT As Object, objNFiscal As ClassNFiscal, objItemNF As ClassItemNF, iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_Preenche_GridItem

    objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col) = Format(objItemNF.dPrecoUnitarioMoeda, gobjFAT.sFormatoPrecoUnitario)
    
    NFFP_Preenche_GridItem = SUCESSO
     
    Exit Function
    
Erro_NFFP_Preenche_GridItem:

    NFFP_Preenche_GridItem = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179294)
     
    End Select
     
    Exit Function

End Function

Function NFFP_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
        
        'Preco Unitário em Moeda
        Case objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col
            lErro = Saida_Celula_PrecoUnitarioMoeda(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 124191
    
    End Select
    
    NFFP_Saida_Celula_GridItens = SUCESSO
     
    Exit Function
    
Erro_NFFP_Saida_Celula_GridItens:

    NFFP_Saida_Celula_GridItens = gErr
     
    Select Case gErr
          
        Case 124191
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179295)
     
    End Select
     
    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitarioMoeda(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long, dTaxaConversao As Double, dPrecoUnitario As Double
Dim dQuantidade As Double, iLinha As Integer
Dim dQuantCancelada As Double

On Error GoTo Erro_Saida_Celula_PrecoUnitarioMoeda

    Set objGridInt.objControle = PrecoUnitarioMoeda(objCT)

    'Verifica o preenchimento de Preço Unitário
    If Len(Trim(PrecoUnitarioMoeda(objCT).Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitarioMoeda(objCT).Text)
        If lErro <> SUCESSO Then gError 46931

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46932
    
    'recalcular preco unitario em R$
    dTaxaConversao = StrParaDbl(TaxaConversao(objCT).Text)
    If dTaxaConversao = 0 Then dTaxaConversao = 1
        
    iLinha = objGridInt.objGrid.Row

    dPrecoUnitario = dTaxaConversao * StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col))
    If dPrecoUnitario > 0 Then
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    Else
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = ""
    End If
    
    'Calcula o Preco Total do item
    Call objCT.PrecoTotal_Calcula(iLinha, False)

    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = objCT.gobjTribTab.Alteracao_Item_Grid(iLinha)
    If lErro <> SUCESSO Then gError 56862
    '*** fim tributacao
    
    'Calcula o Valor Total da Nota
    Call objCT.ValorTotal_Calcula

    Saida_Celula_PrecoUnitarioMoeda = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitarioMoeda:

    Saida_Celula_PrecoUnitarioMoeda = gErr

    Select Case gErr

        Case 46930, 46931, 46932, 56862
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179296)

    End Select

    Exit Function

End Function

Function NFFP_Move_Tela_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_Move_Tela_Memoria

    Set objNFiscal.objInfoUsu = New ClassNFiscalHic

    objNFiscal.objInfoUsu.dValorFrete1 = StrParaDbl(ValorFrete1(objCT).Text)
    objNFiscal.iMoeda = Codigo_Extrai(Moeda(objCT).Text)
    objNFiscal.dTaxaMoeda = StrParaDbl(TaxaConversao(objCT).Text)
    
    NFFP_Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFFP_Move_Tela_Memoria:

    NFFP_Move_Tela_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179297)
     
    End Select
     
    Exit Function

End Function

Function NFFP_Move_GridItem_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal, objItemNF As ClassItemNF, iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_Move_GridItem_Memoria

    objItemNF.dPrecoUnitarioMoeda = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col))
    
    NFFP_Move_GridItem_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFFP_Move_GridItem_Memoria:

    NFFP_Move_GridItem_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179298)
     
    End Select
     
    Exit Function

End Function

Public Function NFFP_Transfere_Dados_Pedido_NFiscal2(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objNFiscal As ClassNFiscal, ByVal bFaturadoParc As Boolean) As Long
    
Dim lErro As Long, objItemNF As ClassItemNF

On Error GoTo Erro_NFFP_Transfere_Dados_Pedido_NFiscal2

    objNFiscal.iMoeda = objPedidoVenda.iMoeda
    
    For Each objItemNF In objNFiscal.ColItensNF
    
        objItemNF.dPrecoUnitarioMoeda = objItemNF.dPrecoUnitario
        
    Next
    
    NFFP_Transfere_Dados_Pedido_NFiscal2 = SUCESSO
     
    Exit Function
    
Erro_NFFP_Transfere_Dados_Pedido_NFiscal2:

    NFFP_Transfere_Dados_Pedido_NFiscal2 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179299)
     
    End Select
     
    Exit Function

End Function

'###########################################################################
'Inserido por Wagner
Function NFFP_Botao_Estoque(ByVal objCT As Object, ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer
Dim sInterroga As String

On Error GoTo Erro_NFFP_Botao_Estoque
    
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
    
        lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 131874
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            If Len(Trim(sInterroga)) > 0 Then
                sInterroga = sInterroga & ", ?"
            Else
                sInterroga = "?"
            End If
    
            objEstoqueProd.sProduto = sProdutoFormatado
    
            colSelecao.Add sProdutoFormatado
    
            lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 131875
    
            If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado
    
    
        End If
        
    Next
    
    colSelecao.Add "%"
    
    sSelecao = "Produto IN (" & sInterroga & ")"

    NFFP_Botao_Estoque = SUCESSO
    
    Exit Function

Erro_NFFP_Botao_Estoque:

    NFFP_Botao_Estoque = gErr

    Select Case gErr

        Case 131874, 131875 'Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179300)

    End Select

    Exit Function
    
End Function

Public Function NFFP_Atualiza_PrecoUnitarioMoeda(ByVal objCT As Object, ByVal iLinha As Integer, ByVal dFator As Double) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dTaxaConversao As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dQuantCancelada As Double

On Error GoTo Erro_NFFP_Atualiza_PrecoUnitarioMoeda
   
    'recalcular preco unitario em R$
    dTaxaConversao = StrParaDbl(TaxaConversao(objCT).Text)
    If dTaxaConversao = 0 Then dTaxaConversao = 1
        
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col) = Format(StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col)) / dFator, gobjFAT.sFormatoPrecoUnitario)
        
    dPrecoUnitario = Arredonda_Moeda(dTaxaConversao * StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col)))
    If dPrecoUnitario > 0 Then
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    Else
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = ""
    End If

    NFFP_Atualiza_PrecoUnitarioMoeda = SUCESSO

    Exit Function

Erro_NFFP_Atualiza_PrecoUnitarioMoeda:

    NFFP_Atualiza_PrecoUnitarioMoeda = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179301)

    End Select

    Exit Function

End Function
'###########################################################################

'#########################################################
'Inserido por Wagner
'Public Sub PercDesconto_Change(ByVal objCT As Object)
'
'    objCT.iAlterado = REGISTRO_ALTERADO
'    objCT.iValorDescontoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercDesconto_Validate(ByVal objCT As Object, Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorProdutos As Double
'Dim iIndice As Integer
'Dim dPercDesconto As Double
'
'On Error GoTo Erro_PercDesconto_Validate
'
'    'Vale o desconto que foi colocado aqui
'    objCT.giValorDescontoManual = 1
'
'    dPercDesconto = StrParaDbl(PercDesconto(objCT).Text)
'
'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
'        If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) > 0 Then
'            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    'Verifica se o Valor está preenchido
'    If dPercDesconto > 0 Then
'
'        'Faz a Crítica do Valor digitado
'        lErro = Porcentagem_Critica(PercDesconto(objCT).Text)
'        If lErro <> SUCESSO Then gError 132017
'
'    End If
'
'    objCT.ValorDesconto.Text = Format((StrParaDbl(objCT.ValorProdutos.Caption) + StrParaDbl(objCT.ValorDesconto)) * (dPercDesconto / 100), "STANDARD")
'
'    'Recalcula o Valor do desconto
'    lErro = objCT.ValorDesconto_Validate(bSGECancelDummy)
'    If lErro <> SUCESSO Then gError 132018
'
'    objCT.iValorDescontoAlterado = 0
'
'    Exit Sub
'
'Erro_PercDesconto_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 132017 To 132018
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179302)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
Function NFFP_ValorDesconto_Validate(ByVal objCT As Object) As Long
'
'Dim lErro As Long
'Dim dValorDesconto As Double
'Dim dValorProdutos As Double
'
'On Error GoTo Erro_NFFP_ValorDesconto_Validate
'
'    dValorDesconto = StrParaDbl(objCT.ValorDesconto)
'    dValorProdutos = StrParaDbl(objCT.ValorProdutos.Caption)
'
'    If 0 > dValorProdutos Then gError 132016
'
'    If dValorDesconto + dValorProdutos <> 0 Then
'        PercDesconto(objCT).Text = (dValorDesconto / (dValorDesconto + dValorProdutos)) * 100
'    Else
'        PercDesconto(objCT).Text = "0"
'    End If
'
'    NFFP_ValorDesconto_Validate = SUCESSO
'
'    Exit Function
'
'Erro_NFFP_ValorDesconto_Validate:
'
'    NFFP_ValorDesconto_Validate = gErr
'
'    Select Case gErr
'
'        Case 132016
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos + dValorDesconto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179303)
'
'    End Select
'
'    Exit Function
'
End Function
'###########################################################################

Public Function NFFP_Atualiza_PrecoUnitarioMoeda2(ByVal objCT As Object, ByVal iLinha As Integer) As Long
'Calcula o Preço unitário moeda a partir do Preço unitário

Dim lErro As Long
Dim dTaxaConversao As Double
Dim dPrecoTabela As Double

On Error GoTo Erro_NFFP_Atualiza_PrecoUnitarioMoeda2
   
    'recalcular preco unitario em R$
    dTaxaConversao = StrParaDbl(TaxaConversao(objCT).Text)
    If dTaxaConversao = 0 Then dTaxaConversao = 1
        
    dPrecoTabela = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col))
        
    'Preço Moeda = Preço Tabela de Preço
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitarioMoeda_Col) = Format(dPrecoTabela, gobjFAT.sFormatoPrecoUnitario)

    'Preço Real = Preço Tabela de Preço * dTaxaConversao
    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = Format(dPrecoTabela * dTaxaConversao, gobjFAT.sFormatoPrecoUnitario)

    NFFP_Atualiza_PrecoUnitarioMoeda2 = SUCESSO

    Exit Function

Erro_NFFP_Atualiza_PrecoUnitarioMoeda2:

    NFFP_Atualiza_PrecoUnitarioMoeda2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179304)

    End Select

    Exit Function

End Function
