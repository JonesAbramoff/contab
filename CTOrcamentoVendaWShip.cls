VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcVendaWShip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PrecoUnitImp(ByVal objCT As Object) As Object
     Set PrecoUnitImp = objCT.objUserControl.Controls("PrecoUnitImp")
End Property

Public Property Get PrecoTotalImp(ByVal objCT As Object) As Object
     Set PrecoTotalImp = objCT.objUserControl.Controls("PrecoTotalImp")
End Property

Public Property Get TabelaPrecoImp(ByVal objCT As Object) As Object
     Set TabelaPrecoImp = objCT.objUserControl.Controls("TabelaPrecoImp")
End Property

Public Function OV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim lErro As Long
Dim iIncremento As Integer
Dim objUserControl As Object

On Error GoTo Erro_OV_Inicializa_Grid_Itens

    Set objGridInt.objForm = objCT
    Set objUserControl = objCT

    'Títulos das colunas
    If gobjFAT.iOVExibeOrdItens = MARCADO Then
        objGridInt.colColuna.Add (" ")
    Else
        objGridInt.colColuna.Add ("Item")
    End If
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colColuna.Add ("S")
    End If
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens4", objGridInt)
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Importação")
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens1", objGridInt)
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Total Import.")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Data Entrega")
    
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Motivo Perda")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colCampo.Add (objCT.SomaItem.Name)
    End If
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (objCT.VersaoKit.Name)
        objGridInt.colCampo.Add (objCT.VersaoKitBase.Name)
    Else
        objCT.VersaoKit.left = POSICAO_FORA_TELA
        objCT.VersaoKit.TabStop = False
        objCT.VersaoKitBase.left = POSICAO_FORA_TELA
        objCT.VersaoKitBase.TabStop = False
    End If

    Call CF("Orcamento_Inicializa_Grid_Itens5", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (PrecoUnitImp(objCT).Name)
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens2", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    objGridInt.colCampo.Add (PrecoTotalImp(objCT).Name)
    objGridInt.colCampo.Add (objCT.PrazoEntregaItem.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
    
    objGridInt.colCampo.Add (objCT.StatusItem.Name)
    objGridInt.colCampo.Add (objCT.MotivoPerdaItem.Name)
    objGridInt.colCampo.Add (objCT.Observacao.Name)
    
    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    If gobjFAT.iOVExibeTS = MARCADO Then
        objCT.iGrid_SomaItem_Col = 1
        iIncremento = iIncremento + 1
    End If
    objCT.iGrid_Produto_Col = 1 + iIncremento
    objCT.iGrid_DescProduto_Col = 2 + iIncremento
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objCT.iGrid_VersaoKit_Col = 3 + iIncremento
        objCT.iGrid_VersaoKitBase_Col = 4 + iIncremento
        iIncremento = iIncremento + 2
        
    End If
    
    Call CF("Orcamento_Inicializa_Grid_Itens6", iIncremento)
    objCT.iGrid_UnidadeMed_Col = 3 + iIncremento
    objCT.iGrid_quantidade_Col = 4 + iIncremento
    objCT.iGrid_PrecoUnitario_Col = 5 + iIncremento
    
    objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col = 6 + iIncremento
    iIncremento = iIncremento + 1

    Call CF("Orcamento_Inicializa_Grid_Itens3", iIncremento)
    objCT.iGrid_PercDesc_Col = 6 + iIncremento
    objCT.iGrid_Desconto_Col = 7 + iIncremento
    objCT.iGrid_PrecoTotal_Col = 8 + iIncremento
    objCT.iGrid_PrecoTotalB_Col = 9 + iIncremento
    
    objCT.gobjInfoUsu.iGrid_PrecoTotalImp_Col = 10 + iIncremento
    iIncremento = iIncremento + 1
    
    objCT.iGrid_PrazoEntrega_Col = 10 + iIncremento
    objCT.iGrid_DataEntrega_Col = 11 + iIncremento
    
    objCT.iGrid_Status_Col = 12 + iIncremento
    objCT.iGrid_MotivoPerda_Col = 13 + iIncremento
    objCT.iGrid_Observacao_Col = 14 + iIncremento
    
    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.iLinhasVisiveis = 7
    Else
        objGridInt.iLinhasVisiveis = 8
    End If

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    OV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

Erro_OV_Inicializa_Grid_Itens:

    OV_Inicializa_Grid_Itens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213594)

    End Select

    Exit Function

End Function

Public Function OV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objItemOVWShip As New ClassItemOVWShip

On Error GoTo Erro_OV_Move_GridItens_Memoria

    objItemOVWShip.dPrecoUnitImp = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col))
    objItemOVWShip.iTabelaPrecoImp = Codigo_Extrai(TabelaPrecoImp(objCT).Text)
    
    Set objItemOV.objInfoUsu = objItemOVWShip

    OV_Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_OV_Move_GridItens_Memoria:

    OV_Move_GridItens_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213594)

    End Select

    Exit Function

End Function

Public Function OV_Traz_GridItens_Tela(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objItemOVWShip As ClassItemOVWShip

On Error GoTo Erro_OV_Traz_GridItens_Tela

    Set objItemOVWShip = objItemOV.objInfoUsu

    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col) = Format(objItemOVWShip.dPrecoUnitImp, gobjFAT.sFormatoPrecoUnitario)
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoTotalImp_Col) = Format(objItemOVWShip.dPrecoUnitImp * objItemOV.dQuantidade, "STANDARD")
    
    If objItemOVWShip.iTabelaPrecoImp > 0 Then
        TabelaPrecoImp(objCT).Text = objItemOVWShip.iTabelaPrecoImp
        Call TabelaPrecoImp_Validate(objCT, bSGECancelDummy)
    Else
        TabelaPrecoImp(objCT).Text = ""
    End If
    
    OV_Traz_GridItens_Tela = SUCESSO

    Exit Function

Erro_OV_Traz_GridItens_Tela:

    OV_Traz_GridItens_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213595)

    End Select

    Exit Function

End Function

Public Function OV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col
            lErro = Saida_Celula_PrecoUnitImp(objCT, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End Select

    OV_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_OV_Saida_Celula_GridItens:

    OV_Saida_Celula_GridItens = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213596)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitImp(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Base que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PrecoUnitImp

    Set objGridInt.objControle = PrecoUnitImp(objCT)

    If Len(Trim(PrecoUnitImp(objCT).Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitImp(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        PrecoUnitImp(objCT).Text = Format(PrecoUnitImp(objCT).Text, gobjFAT.sFormatoPrecoUnitario)

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Recalcula_PrecoTotalImp(objCT, objCT.GridItens.Row)

    Saida_Celula_PrecoUnitImp = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitImp:

    Saida_Celula_PrecoUnitImp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213597)

    End Select

    Exit Function

End Function

Private Function Recalcula_PrecoTotalImp(ByVal objCT As Object, ByVal iLinhaGrid As Integer) As Long

Dim lErro As Long
Dim dPreco As Double, dQtde As Double

On Error GoTo Erro_Recalcula_PrecoTotalImp

    If iLinhaGrid <> 0 Then
        
        dPreco = StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col))
        dQtde = StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_quantidade_Col))
        
        objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoTotalImp_Col) = Format(dPreco * dQtde, "STANDARD")
    
    End If
    
    Recalcula_PrecoTotalImp = SUCESSO
    
    Exit Function
    
Erro_Recalcula_PrecoTotalImp:

    Recalcula_PrecoTotalImp = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213598)
    
    End Select
    
    Exit Function
    
End Function

Public Sub PrecoUnitImp_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrecoUnitImp_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub PrecoUnitImp_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub PrecoUnitImp_Validate(ByVal objCT As Object, Cancel As Boolean)
Dim lErro As Long
    Set objCT.objGridItens.objControle = PrecoUnitImp(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function OV_Rotina_Grid_Enable(ByVal objCT As Object, ByVal sNomeControle As String, ByVal iProdutoPreenchido As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Rotina_Grid_Enable

    Select Case sNomeControle
        
        Case PrecoUnitImp(objCT).Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or Codigo_Extrai(TabelaPrecoImp(objCT)) = 0 Then
                PrecoUnitImp(objCT).Enabled = False
            Else
                PrecoUnitImp(objCT).Enabled = True
            End If
    
        Case PrecoTotalImp(objCT).Name
            PrecoTotalImp(objCT).Enabled = False
    
    End Select
    
    OV_Rotina_Grid_Enable = SUCESSO
     
    Exit Function
    
Erro_OV_Rotina_Grid_Enable:

    OV_Rotina_Grid_Enable = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213599)
     
    End Select
     
    Exit Function

End Function

Public Function OV_Form_Load(ByVal objCT As Object) As Long

    PrecoUnitImp(objCT).Format = gobjFAT.sFormatoPrecoUnitario
    
    Call Carrega_TabelaPreco(objCT)

    OV_Form_Load = SUCESSO
    
End Function

Public Function OV_Saida_Celula_Quantidade(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_Quantidade

    'Reclacula o preço unitário
    lErro = Recalcula_PrecoTotalImp(objCT, objCT.GridItens.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    OV_Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_OV_Saida_Celula_Quantidade:

    OV_Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213600)

    End Select

    Exit Function
    
End Function

Public Function OV_Produto_Saida_Celula(ByVal objCT As Object, objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Produto_Saida_Celula

    lErro = Trata_TabelaPrecoImp_Item(objCT, objCT.GridItens.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    OV_Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_OV_Produto_Saida_Celula:

    OV_Produto_Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213601)

    End Select

    Exit Function

End Function

Public Function OV_Limpa_OrcamentoVenda2(ByVal objCT As Object) As Long
    TabelaPrecoImp(objCT).Text = ""
    OV_Limpa_OrcamentoVenda2 = SUCESSO
End Function

Public Sub TabelaPrecoImp_Click(ByVal objCT As Object)

Dim lErro As Long

On Error GoTo Erro_TabelaPrecoImp_Click

    objCT.iAlterado = REGISTRO_ALTERADO

    If TabelaPrecoImp(objCT).ListIndex = -1 Then Exit Sub

    If objCT.objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPrecoImp(objCT)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_TabelaPrecoImp_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213602)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPrecoImp_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPrecoImp_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPrecoImp(objCT).Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPrecoImp(objCT).Text = TabelaPrecoImp(objCT).List(TabelaPrecoImp(objCT).ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPrecoImp(objCT), iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM

        If lErro <> SUCESSO Then gError 213603 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPrecoImp(objCT).Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPrecoImp(objCT)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 213604

    Exit Sub

Erro_TabelaPrecoImp_Validate:

    Cancel = True

    Select Case gErr

    Case ERRO_SEM_MENSAGEM

    Case 213603  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPrecoImp(objCT).Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPrecoImp(objCT).Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        End If

    Case 213604
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPrecoImp(objCT).Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213605)

    End Select

    Exit Sub

End Sub

Private Function Trata_TabelaPrecoImp(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objTabelaPreco As New ClassTabelaPreco
Dim objMoeda As New ClassMoedas

On Error GoTo Erro_Trata_TabelaPrecoImp

    If Not objCT.gbCarregandoTela Then
        For iLinha = 1 To objCT.objGridItens.iLinhasExistentes
            lErro = Trata_TabelaPrecoImp_Item(objCT, iLinha)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Next
    End If

    objTabelaPreco.iCodigo = Codigo_Extrai(TabelaPrecoImp(objCT))

    'Tenta ler TabelaPreço com esse código no BD
    lErro = CF("TabelaPreco_Le", objTabelaPreco)
    If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 84016  'Não encontrou Tabela Preço no BD

    objMoeda.iCodigo = objTabelaPreco.iMoeda

    lErro = CF("Moedas_Le", objMoeda)
    If lErro <> SUCESSO And lErro <> 108821 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
        objCT.GridItens.TextMatrix(0, objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col) = "Importação(" & objMoeda.sSimbolo & ")"
    End If

    Trata_TabelaPrecoImp = SUCESSO

    Exit Function

Erro_Trata_TabelaPrecoImp:

    Trata_TabelaPrecoImp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213606)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPrecoImp_Item(ByVal objCT As Object, ByVal iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_Trata_TabelaPrecoImp_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objTabelaPrecoItem.sCodProduto = sProduto
    objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPrecoImp(objCT).Text)
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

    If objTabelaPrecoItem.iCodTabela <> 0 And iPreenchido <> PRODUTO_VAZIO Then

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
        'Se não encontrar
        If lErro = 28014 Then
        
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
        End If
    
        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoUnitImp_Col) = Format(objTabelaPrecoItem.dPreco, gobjFAT.sFormatoPrecoUnitario)
        
            Call Recalcula_PrecoTotalImp(objCT, iLinha)
        End If

    End If

    Trata_TabelaPrecoImp_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPrecoImp_Item:

    Trata_TabelaPrecoImp_Item = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM  'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213607)

    End Select

    Exit Function

End Function

Private Function Carrega_TabelaPreco(ByVal objCT As Object) As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPrecoImp(objCT).AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPrecoImp(objCT).ItemData(TabelaPrecoImp(objCT).NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213608)

    End Select

    Exit Function

End Function
