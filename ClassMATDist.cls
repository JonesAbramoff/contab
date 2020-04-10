VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public objTela As Object

Public objGridDist As AdmGrid
Public iGrid_ItemNFDist_Col As Integer
Public iGrid_ProdutoAlmoxDist_Col As Integer
Public iGrid_AlmoxDist_Col As Integer
Public iGrid_QuantItemNFDist_Col As Integer
Public iGrid_QuantDist_Col As Integer
Public iGrid_UMDist_Col As Integer

Public bTela As Boolean 'indica se é uma tela = True (UserControl) ou um classe que contem o codigo da tela (Ex. CTNFiscalEntrada.cls)

Private objUserControl As Object

Public objRastreamento As Object

Public iAlmoxarifadoPadraoTela As Integer

'mario distribuicao
Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1

Private Sub Class_Initialize()
        
        Set objEventoLocalizacao = New AdmEvento
        Set objGridDist = New AdmGrid
        bTela = False

End Sub

Private Sub Class_Terminate()
    
    Set objEventoLocalizacao = Nothing
    
End Sub

Public Property Get ItemNFDist() As Object
     Set ItemNFDist = objUserControl.Controls("ItemNFDist")
End Property

Public Property Get ProdutoAlmoxDist() As Object
     Set ProdutoAlmoxDist = objUserControl.Controls("ProdutoAlmoxDist")
End Property

Public Property Get AlmoxDist() As Object
     Set AlmoxDist = objUserControl.Controls("AlmoxDist")
End Property

Public Property Get QuantItemNFDist() As Object
     Set QuantItemNFDist = objUserControl.Controls("QuantItemNFDist")
End Property

Public Property Get QuantDist() As Object
     Set QuantDist = objUserControl.Controls("QuantDist")
End Property

Public Property Get UMDist() As Object
     Set UMDist = objUserControl.Controls("UMDist")
End Property

Public Property Get GridDist() As Object
     Set GridDist = objUserControl.Controls("GridDist")
End Property

Public Property Get BotaoLocalizacaoDist() As Object
     Set BotaoLocalizacaoDist = objUserControl.Controls("BotaoLocalizacaoDist")
End Property

Public Function Inicializa_GridDist() As Long
'Inicializa o Grid de Itens

'Const NUM_MAX_ALOCACOES = 50
Dim objGridInt As AdmGrid

    Set objGridInt = objGridDist

    Set objGridInt.objForm = objTela

    If bTela Then
        Set objUserControl = objTela
    Else
        Set objUserControl = objTela.objUserControl
    End If

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("ItemNF")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. ItemNF")
    objGridInt.colColuna.Add ("Quant. Almox")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemNFDist.Name)
    objGridInt.colCampo.Add (ProdutoAlmoxDist.Name)
    objGridInt.colCampo.Add (AlmoxDist.Name)
    objGridInt.colCampo.Add (UMDist.Name)
    objGridInt.colCampo.Add (QuantItemNFDist.Name)
    objGridInt.colCampo.Add (QuantDist.Name)

    'Colunas da Grid
    iGrid_ItemNFDist_Col = 1
    iGrid_ProdutoAlmoxDist_Col = 2
    iGrid_AlmoxDist_Col = 3
    iGrid_UMDist_Col = 4
    iGrid_QuantItemNFDist_Col = 5
    iGrid_QuantDist_Col = 6
    
    'Grid do GridInterno
    objGridInt.objGrid = GridDist

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ALOCACOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8
    
    'Largura da primeira coluna
    GridDist.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridDist = SUCESSO

    Exit Function

End Function

Public Function Move_GridDist_Memoria(objNFiscal As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim objItemAloc As Object, objItemAloc2 As Object
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim bExecuta As Boolean
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_GridDist_Memoria

    'Se o número de linhas existentes e  maior que zero
    If objGridDist.iLinhasExistentes > 0 Then

        iIndice = 0
        'Percorre o grid de alocações
        For iIndice = 1 To objGridDist.iLinhasExistentes

            iItem = CInt(GridDist.TextMatrix(iIndice, iGrid_ItemNFDist_Col))
            
            bExecuta = False
            
            If Not Possui_Grade() Then
                bExecuta = True
            ElseIf objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade.Count = 0 Then
                bExecuta = True
            End If
                        
            If bExecuta Then

                Set objItemAloc = CreateObject("GlobaisCRFAT.ClassItemNFAlocacao")
                
                'Preenche o objItemAloc
                objItemAloc.dQuantidade = StrParaDbl(GridDist.TextMatrix(iIndice, iGrid_QuantDist_Col))
                objItemAloc.sAlmoxarifado = GridDist.TextMatrix(iIndice, iGrid_AlmoxDist_Col)
                objItemAloc.sUnidadeMed = GridDist.TextMatrix(iIndice, iGrid_UMDist_Col)
                
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iIndice, iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 199166

                objItemAloc.sProduto = sProdutoFormatado
                
                If colCodigoNome.Count > 0 Then
                
                    'Verifica se o Almoxarifado já foi lido
                    For Each objCodigoNome In colCodigoNome
                        
                        If objCodigoNome.sNome = objItemAloc.sAlmoxarifado Then
                            objItemAloc.iAlmoxarifado = objCodigoNome.iCodigo
                            Exit For
                        End If
                    Next
                End If
                
                'Se o almoxarifado ainda não foi lido
                If objItemAloc.iAlmoxarifado = 0 Then
                
                    objAlmoxarifado.sNomeReduzido = objItemAloc.sAlmoxarifado
                    
                    'Lê o Almoxarifado à partir do nome reduzido
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 35939
                    
                    If lErro = 25060 Then gError 35940
                    
                    'Adiciona o almoxarifado na coleção de almoxarifados lidos
                    objItemAloc.iAlmoxarifado = objAlmoxarifado.iCodigo
    
                    colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                End If
            
            
            
                'Adiciona a alocação do grid na coleção de itens
                With objItemAloc
                    Set objItemAloc2 = objNFiscal.ColItensNF(iItem).ColAlocacoes.Add(.iAlmoxarifado, .sAlmoxarifado, .dQuantidade)
                    objItemAloc2.sUnidadeMed = .sUnidadeMed
                    objItemAloc2.sProduto = .sProduto
                End With

                objNFiscal.ColItensNF(iItem).iAlmoxarifado = objItemAloc.iAlmoxarifado
            
            End If
                            
        Next
        
    End If



    Move_GridDist_Memoria = SUCESSO

    Exit Function

Erro_Move_GridDist_Memoria:

    Move_GridDist_Memoria = gErr
    
    Select Case gErr

        Case 35939, 86375, 199166

        Case 35940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151562)

    End Select

    Exit Function

End Function

Public Sub ItemNFDist_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ItemNFDist_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDist)

End Sub

Public Sub ItemNFDist_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDist)

End Sub

Public Sub ItemNFDist_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDist.objControle = ItemNFDist
    lErro = Grid_Campo_Libera_Foco(objGridDist)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AlmoxDist_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AlmoxDist_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDist)

End Sub

Public Sub AlmoxDist_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDist)

End Sub

Public Sub AlmoxDist_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDist.objControle = AlmoxDist
    lErro = Grid_Campo_Libera_Foco(objGridDist)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantDist_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantDist_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDist)

End Sub

Public Sub QuantDist_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDist)

End Sub

Public Sub QuantDist_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDist.objControle = QuantDist
    lErro = Grid_Campo_Libera_Foco(objGridDist)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Distribuicao_Obtem_Produto_Cust(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF, sProdutoMascarado As String) As Long
    Distribuicao_Obtem_Produto_Cust = SUCESSO
End Function

Private Function Saida_Celula_ItemNFDist(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dFator As Double
Dim iItem As Integer
Dim iItem_Atual As Integer
Dim bExecuta As Boolean
Dim sProdutoMascarado As String
Dim objNF As ClassNFiscal
Dim objItemNF As ClassItemNF

On Error GoTo Erro_Saida_Celula_ItemNFDist

    Set objGridInt.objControle = ItemNFDist

    bExecuta = False
    
    Set objNF = objTela.gobjnfiscal
    
    'Verifica se o Produto esta preenchido
    If Len(Trim(ItemNFDist.Text)) = 0 And Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col))) > 0 Then gError 89468
    
    If Len(Trim(ItemNFDist.Text)) > 0 Then

        iItem = CInt(ItemNFDist.Text)
      
        If iItem < 1 Or iItem > objTela.objGridItens.iLinhasExistentes Then gError 89469

        'Se o produto não está preenchido ==> erro
        sProdutoMascarado = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col)
        
        If Not (objNF Is Nothing) Then
            Set objItemNF = objNF.ColItensNF.Item(iItem)
            lErro = CF("Distribuicao_Obtem_Produto_Cust", objNF, objItemNF, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 89471
        End If
        If Len(Trim(sProdutoMascarado)) = 0 Then gError 89470
        
        If Not Possui_Grade Then
            bExecuta = True
        ElseIf objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade.Count = 0 Then
            bExecuta = True
        End If
        
        If bExecuta Then
        
            'Formata o Produto para o BD
            lErro = CF("Produto_Formata", sProdutoMascarado, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 89471
                
            'Lê os demais atributos do Produto
            objProduto.sCodigo = sProdutoFormatado
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 89472
                
            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 89473
                    
            'descobre qual o item atual
            If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col))) > 0 Then
                iItem_Atual = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col))
            End If
            
            'se o item que se está preenchendo é diferente do item atual, ==> limpa os campos de lote
            If iItem <> iItem_Atual Then
            
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantDist_Col) = ""
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxDist_Col) = ""
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoAlmoxDist_Col) = sProdutoMascarado
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89474
                
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantItemNFDist_Col) = Formata_Estoque(StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Quantidade_Col)) * dFator)
                
            End If
            
        End If
    
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89493

    Saida_Celula_ItemNFDist = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFDist:

    Saida_Celula_ItemNFDist = gErr

    Select Case gErr
        
        Case 89468, 89469
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMDISTRIBUICAO_NAO_ITEMNF", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 89470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID_ITENS", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89471, 89472, 89474, 89493
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 89473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151563)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AlmoxDist(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_AlmoxDist

    Set objGridInt.objControle = AlmoxDist

    'Se o Almoxarifado está preenchido
    If Len(Trim(AlmoxDist.Text)) > 0 Then

        lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 89475

        'Valida o ALmoxarifado
        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, AlmoxDist, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 89476
        
        'Se não for encontrado --> Erro
        If lErro = 25157 Then gError 89477
        
        If lErro = 25162 Then gError 89478

        For iLinha = 1 To objGridInt.iLinhasExistentes
        
            If iLinha <> objGridInt.objGrid.Row Then
                If objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col) = objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col) And _
                   objGridInt.objGrid.TextMatrix(iLinha, iGrid_AlmoxDist_Col) = objAlmoxarifado.sNomeReduzido And _
                   objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoAlmoxDist_Col) = objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoAlmoxDist_Col) Then gError 89513
            End If
            
        Next

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89479

    If objGridInt.objGrid.Row <= objGridInt.iLinhasExistentes And Not objRastreamento Is Nothing Then
        lErro = objRastreamento.Atualiza_Troca_Almoxarifado1(objGridInt.objGrid.Row)
        If lErro <> SUCESSO Then gError 89480
    End If

    Saida_Celula_AlmoxDist = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxDist:

    Saida_Celula_AlmoxDist = gErr

    Select Case gErr

        Case 89475, 89476, 89479, 89480
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89477
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", AlmoxDist.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                objAlmoxarifado.sNomeReduzido = AlmoxDist.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 89478

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(AlmoxDist.Text))
            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(AlmoxDist.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Almoxarifados
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 89513
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOX_REPETIDO_ITEMNF_GRID", gErr, AlmoxDist.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 151564)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantDist(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantDist

    Set objGridInt.objControle = QuantDist

    'Se a quantidade alocada foi preenchida
    If Len(Trim(QuantDist.clipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantDist.Text)
        If lErro <> SUCESSO Then gError 89485

        'Se a quantidade alocada for maior que a quantidade do item da nota ==> erro
        If StrParaDbl(QuantDist.Text) > StrParaDbl(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantItemNFDist_Col)) Then gError 89486
        
        QuantDist.Text = Formata_Estoque(QuantDist.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89487

    Saida_Celula_QuantDist = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantDist:

    Saida_Celula_QuantDist = gErr

    Select Case gErr

        Case 89485, 89487
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTDIST_MAIOR_QUANTITEMNF", gErr, QuantDist.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantItemNFDist_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151565)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Dist() As Long

Dim lErro As Long
Dim objGridInt As AdmGrid

On Error GoTo Erro_Saida_Celula_Dist

    Set objGridInt = objGridDist

    Select Case objGridInt.objGrid.Col
    
        Case iGrid_ItemNFDist_Col
            lErro = Saida_Celula_ItemNFDist(objGridInt)
            If lErro <> SUCESSO Then gError 89489
    
        Case iGrid_ProdutoAlmoxDist_Col
            lErro = Saida_Celula_ProdutoAlmox(objGridInt)
            If lErro <> SUCESSO Then gError 89489
    
        Case iGrid_AlmoxDist_Col
            lErro = Saida_Celula_AlmoxDist(objGridInt)
            If lErro <> SUCESSO Then gError 89490
        
        'QuantLoteRastro Rastreamento
        Case iGrid_QuantDist_Col
            lErro = Saida_Celula_QuantDist(objGridInt)
            If lErro <> SUCESSO Then gError 89491
    
    End Select

    Saida_Celula_Dist = SUCESSO

    Exit Function

Erro_Saida_Celula_Dist:

    Saida_Celula_Dist = gErr

    Select Case gErr

        Case 89489, 89490, 89491

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151566)

    End Select

    Exit Function

End Function

Public Sub BotaoLocalizacaoDist_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_BotaoLocalizacaoDist_Click:

    'Verifica se tem alguma linha do Grid Selecionada
    If GridDist.Row = 0 Then gError 89495

    'Verifica se o Produto está preenchido
    If Len(Trim(GridDist.TextMatrix(GridDist.Row, iGrid_ProdutoAlmoxDist_Col))) > 0 Then

        lErro = CF("Produto_Formata", GridDist.TextMatrix(GridDist.Row, iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 89496

        If iProdutoPreenchido <> PRODUTO_VAZIO Then
            colSelecao.Add sProdutoFormatado
            
            Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoLocalizacao)
        End If
        
    Else
        gError 89497
    End If

    Exit Sub

Erro_BotaoLocalizacaoDist_Click:

    Select Case gErr

        Case 89495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 89496
        
        Case 89497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151567)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLocalizacao_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto

    Set objEstoqueProduto = obj1

    'Se não tiver nenhuma linha do Grid selecionada --> Sai
    If GridDist.Row = 0 Then Exit Sub

    'Verifica se o Produto está preenchido
    If Len(Trim(GridDist.TextMatrix(GridDist.Row, iGrid_ProdutoAlmoxDist_Col))) > 0 Then
    
        'coloca o Almoxarifado no Grid na linha selecionada
        GridDist.TextMatrix(GridDist.Row, iGrid_AlmoxDist_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        AlmoxDist.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
        
    End If

    objTela.Show

    Exit Sub

End Sub

Public Function Preenche_GridDistribuicao(objNFiscal As Object) As Long
    Dim objCT As Object
    Set objCT = Me
    Preenche_GridDistribuicao = CF("Preenche_GridDistribuicaoAux", objCT, objNFiscal)
End Function

Public Function Preenche_GridDistribuicaoAux(ByVal objCT As Object, objNFiscal As Object) As Long
'Preenche o Grid com as Alocações da Nota Fiscal

Dim objItemAloc As Object
Dim iIndice As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer

On Error GoTo Erro_Preenche_GridDistribuicaoAux
    
    'Limpa o grid de alocações
    Call Grid_Limpa(objCT.objGridDist)
    
    'Para cada item da NF
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        
        'Para cada alocação do Item de NF
        For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).ColAlocacoes

            objCT.objGridDist.iLinhasExistentes = objCT.objGridDist.iLinhasExistentes + 1
        
            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 51081
            If lErro <> SUCESSO Then gError 51082
        
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 51083
            
            'Coloca os dados da alocação na tela
            objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_ItemNFDist_Col) = objNFiscal.ColItensNF(iIndice).iItem
            objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_ProdutoAlmoxDist_Col) = objCT.objTela.objGridItens.objGrid.TextMatrix(objNFiscal.ColItensNF(iIndice).iItem, objCT.objTela.iGrid_Produto_Col)
            objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_AlmoxDist_Col) = objItemAloc.sAlmoxarifado
            objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
            objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_QuantDist_Col) = Formata_Estoque(objItemAloc.dQuantidade)
            dQuantidade = StrParaDbl(objCT.objTela.objGridItens.objGrid.TextMatrix(iIndice, objCT.objTela.iGrid_Quantidade_Col)) * dFator
            
            '#############################################################################
            'Verifica se o Formata estoque está arredondando o valor da quantidade p\ menos
'            If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
'
'                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                dQuantidade = dQuantidade + dAcrescimo
'
'            End If
            dQuantidade = Arredonda_Estoque(dQuantidade)
            '#############################################################################
            
            objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantidade)
        Next
        
        If objNFiscal.ColItensNF(iIndice).iPossuiGrade Then Call objCT.Atualiza_Grid_Distribuicao(objNFiscal.ColItensNF(iIndice))

    Next

    Preenche_GridDistribuicaoAux = SUCESSO

    Exit Function

Erro_Preenche_GridDistribuicaoAux:

    Preenche_GridDistribuicaoAux = gErr
    
    Select Case gErr
    
        Case 51081, 51083
        
        Case 51082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objCT.ProdutoAlmoxDist.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151568)
            
    End Select

End Function

Public Function Preenche_GridDistribuicaoPC(iItemNF As Integer, objItemPC As Object, objProduto As ClassProduto) As Long
'Preenche o Grid com as Alocações do item de Pedido de compra
'Jones 13/04/07: acho que nao é usada e se for deve ter erro de conversao de UM.
Dim objItemAloc As Object
Dim lErro As Long
Dim dFator As Double
Dim dQuantidade As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer
Dim objLocalizacao As Object
Dim iLinha As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iAchou As Integer

On Error GoTo Erro_Preenche_GridDistribuicaoPC
    
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 89616
    
    dQuantidade = StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_Quantidade_Col)) * dFator
    
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMCompra, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 89617
    
    'Para cada alocação do Item de NF
    For Each objLocalizacao In objItemPC.colLocalizacao

        objAlmoxarifado.iCodigo = objLocalizacao.iAlmoxarifado

        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 89618

        'Se não achou o Almoxarifado --> erro
        If lErro = 25056 Then gError 89619

        iAchou = 0

        For iLinha = 1 To objGridDist.iLinhasExistentes

            'se o par item nf/almoxarifado já estiver presente no grid de distribuicao ==> adiciona a quantidade
            If iItemNF = StrParaInt(GridDist.TextMatrix(iLinha, iGrid_ItemNFDist_Col)) And objAlmoxarifado.sNomeReduzido = objGridDist.objGrid.TextMatrix(iLinha, iGrid_AlmoxDist_Col) Then

                GridDist.TextMatrix(iLinha, iGrid_QuantDist_Col) = Formata_Estoque(StrParaDbl(GridDist.TextMatrix(iLinha, iGrid_QuantDist_Col)) + objLocalizacao.dQuantidade * dFator)
                
                '##############################################################################
                'Alterado por Wagner 16/11/04
'                'Verifica se o Formata estoque está arredondando o valor da quantidade p\ menos
'                If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
'
'                    'Calcula o número de casas decimais do Formato de Estoque
'                    iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                    dAcrescimo = 0
'                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                    dQuantidade = dQuantidade + dAcrescimo
'
'                End If
                dQuantidade = Arredonda_Estoque(dQuantidade)
                '##############################################################################
                
                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantidade)
                
                iAchou = 1
                Exit For
                
            End If
            
        Next
        
        'se o par item nf/almoxarifado ainda não estava no grid ==>
        If iAchou = 0 Then
        
            objGridDist.iLinhasExistentes = objGridDist.iLinhasExistentes + 1
        
            'Coloca os dados da alocação na tela
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ItemNFDist_Col) = CStr(iItemNF)
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ProdutoAlmoxDist_Col) = objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_Produto_Col)
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_AlmoxDist_Col) = objAlmoxarifado.sNomeReduzido
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantDist_Col) = Formata_Estoque(objLocalizacao.dQuantidade * dFator)
        
        
            '##############################################################
            'Alterado por Wagner 16/11/04
            'Verifica se o Formata estoque está arredondando o valor da quantidade p\ menos
'            If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
'
'                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                dQuantidade = dQuantidade + dAcrescimo
'
'            End If

            dQuantidade = Arredonda_Estoque(dQuantidade)
            '##############################################################
            
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantidade)
            
        End If
    Next

    Preenche_GridDistribuicaoPC = SUCESSO

    Exit Function

Erro_Preenche_GridDistribuicaoPC:

    Preenche_GridDistribuicaoPC = gErr
    
    Select Case gErr
    
        Case 89616, 89617, 89618
        
        Case 89619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151569)
            
    End Select

End Function

Public Function Preenche_GridDistribuicaoPC1(colItemPedCompraInfo As Collection) As Long
'Percorre cada item do grid de Itens e faz a alocação segundo as informações dos itens de pedido de compra
'recebidos e sua localização conforme fornecido pelo pedido de compra. Se não contiver localização, faz a alocação pelo almoxarifado padrão

Dim lErro As Long
Dim dFator As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iLinha As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iAchou As Integer
Dim dQuantidadeItemNF As Double
Dim dQuantidadeItemNF1 As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim iItemNF As Integer
Dim objProduto As New ClassProduto
Dim objItemPCInfo As Object
Dim objLocalizacao As Object
Dim dQuantidadeItemPC As Double
Dim dQuantidadeDist As Double
Dim bTeste As Boolean

On Error GoTo Erro_Preenche_GridDistribuicaoPC1
    
    'Limpa o grid de alocações
    Call Grid_Limpa(objGridDist)
    
    For iItemNF = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Formata o Produto
        lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 89620
    
        objProduto.sCodigo = sProduto
    
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89621
        
        'se o produto não está cadastrado ==> erro
        If lErro <> SUCESSO Then gError 89622
    
        'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
    
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 89623
        
            dQuantidadeItemNF = StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_Quantidade_Col)) * dFator
            'guarda a quantidade do item NF que ainda falta distribuir
            dQuantidadeItemNF1 = dQuantidadeItemNF
        
            For Each objItemPCInfo In colItemPedCompraInfo
                
                '#################################################################
                'Inserido por Wagner 20/07/2006
                bTeste = True
                Call CF2(objTela, "NFEntCom_Produto_Igual_Customizado", objItemPCInfo.objInfoUsu, iItemNF, bTeste)
                '#################################################################
                
                If objItemPCInfo.sProduto = sProduto And bTeste Then
                
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPCInfo.sUM, objProduto.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 89624
        
                    dQuantidadeItemPC = objItemPCInfo.dQuantRecebida * dFator
                
                    'se tiver quantidade do item do pedido de compra a ser alocada
                    If dQuantidadeItemPC > 0 Then
                
                        For Each objLocalizacao In objItemPCInfo.colLocalizacao
                    
                            objAlmoxarifado.iCodigo = objLocalizacao.iAlmoxarifado
                    
                            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                            If lErro <> SUCESSO And lErro <> 25056 Then gError 89625
                    
                            'Se não achou o Almoxarifado --> erro
                            If lErro = 25056 Then gError 89626
                    
                            dQuantidadeDist = objLocalizacao.dQuantidade * dFator
                            
                            If dQuantidadeItemPC >= dQuantidadeDist Then
                            
                                dQuantidadeItemPC = dQuantidadeItemPC - dQuantidadeDist
                                
                            Else
                            
                                dQuantidadeDist = dQuantidadeItemPC
                                
                                dQuantidadeItemPC = 0
                                
                            End If
                                                                    
                            'guarda a quantidade do item da nf que ainda falta distribuir
                            dQuantidadeItemNF1 = dQuantidadeItemNF1 - dQuantidadeDist
                                
                            iAchou = 0
                                
                            For iLinha = 1 To objGridDist.iLinhasExistentes
                                                                
                                'se o par item nf/almoxarifado já estiver presente no grid de distribuicao ==> adiciona a quantidade
                                If iItemNF = StrParaInt(GridDist.TextMatrix(iLinha, iGrid_ItemNFDist_Col)) And objAlmoxarifado.sNomeReduzido = objGridDist.objGrid.TextMatrix(iLinha, iGrid_AlmoxDist_Col) Then
                    
                                    GridDist.TextMatrix(iLinha, iGrid_QuantDist_Col) = Formata_Estoque(StrParaDbl(GridDist.TextMatrix(iLinha, iGrid_QuantDist_Col)) + dQuantidadeDist)
                                    
                                    iAchou = 1
                                    Exit For
                                    
                                End If
                                                                
                            Next
                                                                    
                            'se o par item nf/almoxarifado ainda não estava no grid ==>
                            If iAchou = 0 Then
                            
                                objGridDist.iLinhasExistentes = objGridDist.iLinhasExistentes + 1
                            
                                'Coloca os dados da alocação na tela
                                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ItemNFDist_Col) = CStr(iItemNF)
                                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ProdutoAlmoxDist_Col) = objTela.objGridItens.objGrid.TextMatrix(iItemNF, objTela.iGrid_Produto_Col)
                                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_AlmoxDist_Col) = objAlmoxarifado.sNomeReduzido
                                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
                                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantDist_Col) = Formata_Estoque(dQuantidadeDist)
                            
                                '##################################################
                                'Alterado por Wagner 16/11/04
                                'Verifica se o Formata estoque está arredondando o valor da quantidade p\ menos
    '                            If StrParaDbl(Formata_Estoque(dQuantidadeItemNF)) < dQuantidadeItemNF Then
    '
    '                                'Calcula o número de casas decimais do Formato de Estoque
    '                                iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
    '
    '                                dAcrescimo = 0
    '                                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
    '
    '                                dQuantidadeItemNF = dQuantidadeItemNF + dAcrescimo
    '
    '                            End If
                                
                                dQuantidadeItemNF = Arredonda_Estoque(dQuantidadeItemNF)
                                '##################################################
                                
                                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantidadeItemNF)
                                
                            End If
                                                                    
                            'se a quantidade do itemPC está zerada ==> significa que não tem mais quantidade a ser distribuida pelos almoxarifados
                            If dQuantidadeItemPC = 0 Then Exit For
                                                                    
                        Next
                    
                    End If
                    
                End If
                
            Next
        
            'se ainda falta alguma quantidade a distribuir
            If dQuantidadeItemNF1 > 0 Then
        
                iAchou = 0
        
                For iLinha = 1 To objGridDist.iLinhasExistentes
                                                    
                    'ao encontrar o primeiro item da distribuicao que coincide com o item nf ==> adiciona a quantidade que falta ser distribuida
                    If iItemNF = StrParaInt(GridDist.TextMatrix(iLinha, iGrid_ItemNFDist_Col)) Then
        
                        GridDist.TextMatrix(iLinha, iGrid_QuantDist_Col) = Formata_Estoque(StrParaDbl(GridDist.TextMatrix(iLinha, iGrid_QuantDist_Col)) + dQuantidadeItemNF1)
                        iAchou = 1
                        Exit For
                        
                    End If
                                                    
                Next
            
                If iAchou = 0 Then
            
                    'Tenta Alocar o produto no Almoxarifado padrão
                    lErro = AlocaAlmoxarifadoPradrao(iItemNF, dQuantidadeItemNF1, objProduto)
                    If lErro <> SUCESSO And lErro <> 89508 Then gError 89627
            
                End If
                
            End If
    
        End If
    
    Next
    
    Preenche_GridDistribuicaoPC1 = SUCESSO

    Exit Function

Erro_Preenche_GridDistribuicaoPC1:

    Preenche_GridDistribuicaoPC1 = gErr
    
    Select Case gErr
    
        Case 89620, 89621, 89623, 89624, 89625, 89627
        
        Case 89622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 89626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151570)
            
    End Select

End Function


Public Function Rotina_Grid_Enable_Dist(iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim sProduto As String

On Error GoTo Erro_Rotina_Grid_Enable_Dist

    Select Case objControl.Name
    
        Case AlmoxDist.Name, QuantDist.Name
        
            If Len(Trim(objGridDist.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_ItemNFDist_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
                
        Case ProdutoAlmoxDist.Name
            
            iItem = StrParaInt(objGridDist.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_ItemNFDist_Col))
            
            If iItem = 0 Then
            
                objControl.Enabled = False
                Rotina_Grid_Enable_Dist = SUCESSO
                Exit Function
           
            Else
            
                'Alterado por Wagner 18/09/06
                'NF\Com não tem Grade (Erro 438)
                If Not (objTela.gobjnfiscal Is Nothing) Then
            
                    If Not Possui_Grade Or objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade.Count = 0 Then
                        
                        objControl.Enabled = False
                        Rotina_Grid_Enable_Dist = SUCESSO
                        Exit Function
                    
                    Else
                         objControl.Enabled = True
                       
                    End If
                    
                Else
                
                    objControl.Enabled = False
                    Rotina_Grid_Enable_Dist = SUCESSO
                    Exit Function
                        
                End If
                
            End If
            
            ProdutoAlmoxDist.Clear
            
            'Atualizar a combo
            If (Len(Trim(objGridDist.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_ItemNFDist_Col))) > 0) And objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade.Count > 0 Then
                
                objControl.Enabled = True
    
                For iIndice = 0 To objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade.Count
                
                    If objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade.Count = 0 Then
                        
                        sProduto = objTela.objGridItens.objGrid.TextMatrix(objTela.objGridItens.objGrid.Row, objTela.iGrid_Produto_Col)
                    Else
                        If iIndice = 0 Then iIndice = iIndice + 1
                        lErro = Mascara_RetornaProdutoTela(objTela.gobjnfiscal.ColItensNF(iItem).colItensRomaneioGrade(iIndice).sProduto, sProduto)
                        If lErro <> SUCESSO Then gError 86374
                    End If
                    
                    ProdutoAlmoxDist.AddItem sProduto
                Next
                
                For iIndice = 0 To ProdutoAlmoxDist.ListCount - 1
                    If objGridDist.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_ProdutoAlmoxDist_Col) = ProdutoAlmoxDist.List(iIndice) Then
                        ProdutoAlmoxDist.ListIndex = iIndice
                    End If
                Next
                
            Else
                objControl.Enabled = False
            End If
            
    End Select
    
    Rotina_Grid_Enable_Dist = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enable_Dist:

    Rotina_Grid_Enable_Dist = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151571)

    End Select

    Exit Function

End Function

Public Function Distribuicao_Processa() As Long
    Dim objCT As Object
    Set objCT = Me
    Distribuicao_Processa = CF("Distribuicao_ProcessaAux", objCT)
End Function

Public Function Distribuicao_ProcessaAux(ByVal objCT As Object) As Long
Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String
Dim dFator As Double
Dim iLinha As Integer
Dim bFat As Boolean, bMovEst As Boolean

On Error GoTo Erro_Distribuicao_ProcessaAux

    iLinha = objCT.objTela.objGridItens.objGrid.Row

    'Remove as alocações do grid de alocações
    Call objCT.Remove_Distribuicao(iLinha)
    
    'Formata o Produto
    lErro = CF("Produto_Formata", objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 89502

    objProduto.sCodigo = sProduto

    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 89503
    
    'se o produto não está cadastrado ==> erro
    If lErro <> SUCESSO Then gError 89504

    'recolhe a quantidade do grid de itens
    dQuantidade = StrParaDbl(objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_Quantidade_Col))

    bMovEst = True

    If Not (objCT.objTela Is Nothing) Then
        If UCase(objCT.objTela.Name) = "NFISCALFATENTRADA" Then
            Call objCT.objTela.gobjTribTab.Testa_Item_Fat(iLinha, bFat, bMovEst)
        End If
    End If

    'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And bMovEst Then
    
        'recolhe a UM do ItemNF
        sUM = objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_UnidadeMed_Col)
        
        'Faz a conversão da UM da tela para a UM de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 89506
    
        'Converte a quantidade para a UM de estoque
        dQuantidade = dQuantidade * dFator

        'Tenta Alocar o produto no Almoxarifado padrão
        lErro = objCT.AlocaAlmoxarifadoPradrao(iLinha, dQuantidade, objProduto)
        If lErro <> SUCESSO Then gError 89505

    End If

    Distribuicao_ProcessaAux = SUCESSO

    Exit Function

Erro_Distribuicao_ProcessaAux:

    Distribuicao_ProcessaAux = gErr

    Select Case gErr

        Case 89502, 89503, 89505

        Case 89504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151572)

    End Select

    Exit Function

End Function

Public Function AlocaAlmoxarifadoPradrao(iLinhaGridItens As Integer, dQuantAlocar As Double, objProduto As ClassProduto) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemNF As Object
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As Object
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim iCodigo As Integer

On Error GoTo Erro_AlocaAlmoxarifadoPradrao

    '##########################################################
    'Alterado por Wagner 24/11/05
'    'Busca o Almoxarifado padrão
'    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
'    If lErro <> SUCESSO And lErro <> 23796 Then gError 89507
'
'    'Se não encontrou ==> não faz a alocação automática
'    If iAlmoxarifado = 0 Then gError 89508

    lErro = AlocaAlmoxarifadoPadraoTela(giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO Then gError 141068
    '##########################################################

    objAlmoxarifado.iCodigo = iAlmoxarifado
    
    'Lê o Aloxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> 25056 And lErro <> SUCESSO Then gError 89509
    
    If lErro = 25056 Then gError 89510
    
    iLinhas = objGridDist.iLinhasExistentes

    iLinhas = iLinhas + 1
    
    'Coloca no Grid alocações a alocação para o almoxarifado padrão
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_ItemNFDist_Col) = iLinhaGridItens
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_ProdutoAlmoxDist_Col) = objTela.objGridItens.objGrid.TextMatrix(iLinhaGridItens, objTela.iGrid_Produto_Col)
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_AlmoxDist_Col) = objAlmoxarifado.sNomeReduzido
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_QuantDist_Col) = Formata_Estoque(dQuantAlocar)
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantAlocar)
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
    
    'atualiza o numero de linhas existentes no grid de alocações
    objGridDist.iLinhasExistentes = iLinhas

    AlocaAlmoxarifadoPradrao = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao:

    AlocaAlmoxarifadoPradrao = gErr

    Select Case gErr

        Case 89506, 89507, 89508, 89509

        Case 89510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 141068 'Inserido por Wagner

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151573)

    End Select

    Exit Function

End Function

Public Sub Remove_Distribuicao(iItemAtual As Integer)
'Remove do grid de distribuicao as alocações do item passado

Dim iIndice As Integer
Dim iItemLinha As Integer

    iIndice = 1
    
    Do While iIndice <= objGridDist.iLinhasExistentes
        If Len(Trim(objGridDist.objGrid.TextMatrix(iIndice, iGrid_ItemNFDist_Col))) > 0 Then
            'Recolhe o item a qual pertence a alocação
            iItemLinha = objGridDist.objGrid.TextMatrix(iIndice, iGrid_ItemNFDist_Col)
            'Se for o item passado em iItemAtual
            If iItemLinha = iItemAtual Then
                'Exclui do grid de alocações as alocações daquele ItemNF
                Do While CInt(objGridDist.objGrid.TextMatrix(iIndice, iGrid_ItemNFDist_Col)) = iItemAtual
                    Call Grid_Exclui_Linha(objGridDist, iIndice)
                    If Len(Trim(objGridDist.objGrid.TextMatrix(iIndice, iGrid_ItemNFDist_Col))) = 0 Then Exit Do
                Loop
            End If
        End If
        iIndice = iIndice + 1
    Loop

    If Not objRastreamento Is Nothing Then
        'Rastreamento
        Call objRastreamento.Exclusao_Item_Grid(iItemAtual)
    End If
    
End Sub

Public Function Exclusao_Item_GridDist(iItemAtual As Integer) As Long

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Exclusao_Item_GridDist

    For iLinha = objGridDist.iLinhasExistentes To 1 Step -1
        If Len(Trim(objGridDist.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col))) > 0 Then
            If CInt(objGridDist.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col)) = iItemAtual Then
                'exclui a linha em questão
                Call Grid_Exclui_Linha(objGridDist, iLinha)
            End If
        End If
    Next
    
    For iLinha = 1 To objGridDist.iLinhasExistentes
        If Len(Trim(objGridDist.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col))) > 0 Then If CInt(objGridDist.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col)) > iItemAtual Then objGridDist.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col) = CInt(objGridDist.objGrid.TextMatrix(iLinha, iGrid_ItemNFDist_Col)) - 1
    Next
    
    If Not objRastreamento Is Nothing Then
    
        'Rastreamento
        lErro = objRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 89254
    
    End If
    
    Exclusao_Item_GridDist = SUCESSO
    
    Exit Function
    
Erro_Exclusao_Item_GridDist:

    Exclusao_Item_GridDist = gErr

    Select Case gErr

        Case 89254

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151574)

    End Select
    
    Exit Function
    
End Function

Public Sub Limpa_Tela_Distribuicao()
'limpa a parte de Distribuicao

    'Limpa o Grid de Distribuicao
    Call Grid_Limpa(objGridDist)
    
End Sub

Public Sub GridDist_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDist, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDist, objTela.iAlterado)
    End If

End Sub

Public Sub GridDist_EnterCell()

    Call Grid_Entrada_Celula(objGridDist, objTela.iAlterado)

End Sub

Public Sub GridDist_GotFocus()

    Call Grid_Recebe_Foco(objGridDist)

End Sub

Public Sub GridDist_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDist, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDist, objTela.iAlterado)
    End If

End Sub

Public Sub GridDist_LeaveCell()

    Call objTela.Saida_Celula(objGridDist)

End Sub

Public Sub GridDist_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridDist)
    
End Sub

Public Sub GridDist_RowColChange()

    Call Grid_RowColChange(objGridDist)

End Sub

Public Sub GridDist_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer, lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_GridDist_KeyDown

    iLinhasExistentesAnterior = objGridDist.iLinhasExistentes
    iItemAtual = GridDist.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridDist)

    If objGridDist.iLinhasExistentes < iLinhasExistentesAnterior And Not objRastreamento Is Nothing Then

        'Rastreamento
        lErro = objRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 83203
        
    End If

    Exit Sub
    
Erro_GridDist_KeyDown:

    Select Case gErr

        Case 83203

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151575)

    End Select

    Exit Sub
    
End Sub

Public Sub GridDist_Scroll()

    Call Grid_Scroll(objGridDist)

End Sub

Public Function Distribuicao_Processa_Grade() As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String
Dim dFator As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProdutoItem As New ClassProduto

On Error GoTo Erro_Distribuicao_Processa_Grade
    
    objProdutoItem.sCodigo = objTela.gobjnfiscal.ColItensNF(objTela.objGridItens.objGrid.Row).sProduto

    'Lê o produto
    lErro = CF("Produto_Le", objProdutoItem)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 180069
        
    For Each objItemRomaneio In objTela.gobjnfiscal.ColItensNF(objTela.objGridItens.objGrid.Row).colItensRomaneioGrade
    
        objProduto.sCodigo = objItemRomaneio.sProduto
    
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89503
        
        'se o produto não está cadastrado ==> erro
        If lErro <> SUCESSO Then gError 89504
    
        'recolhe a quantidade do grid de itens
        dQuantidade = objItemRomaneio.dQuantidade
    
        'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 Then
        
            'O Kit de Venda já atualiza a quantidade
            If objProdutoItem.iKitVendaComp = DESMARCADO Then
        
                'recolhe a UM do ItemNF
                sUM = objTela.objGridItens.objGrid.TextMatrix(objTela.objGridItens.objGrid.Row, objTela.iGrid_UnidadeMed_Col)
        
                'Faz a conversão da UM da tela para a UM de estoque
                lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89506
            
                'Converte a quantidade para a UM de estoque
                dQuantidade = dQuantidade * dFator
                
            End If
    
            'Tenta Alocar o produto no Almoxarifado padrão
            lErro = AlocaAlmoxarifadoPradrao_Grade(objTela.objGridItens.objGrid.Row, dQuantidade, objProduto, objItemRomaneio)
            If lErro <> SUCESSO Then gError 89505
        
        End If

    Next
    
    Distribuicao_Processa_Grade = SUCESSO

    Exit Function

Erro_Distribuicao_Processa_Grade:

    Distribuicao_Processa_Grade = gErr

    Select Case gErr

        Case 89502, 89503, 89505, 180069

        Case 89504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151576)

    End Select

    Exit Function

End Function

Public Function Atualiza_Grid_Distribuicao(objItemNF As ClassItemNF) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReserva As ClassReservaItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim dFator As Double
Dim dFator2 As Double
Dim objProduto As New ClassProduto
Dim dQuantReservada As Double
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Atualiza_Grid_Distribuicao

    Call Remove_Distribuicao(objItemNF.iItem)

    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
    
        objProduto.sCodigo = objItemRomaneio.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
        'Faz a conversão da unidade do item para a unidade de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 42766

        objProduto.sCodigo = objItemNF.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
        If lErro <> SUCESSO Then gError 42766
        
        If objProduto.iKitVendaComp = MARCADO Then dFator2 = 1 'Já foi feita a conversao
        
        dQuantReservada = 0
        
        For Each objReserva In objItemRomaneio.colLocalizacao
        
            objGridDist.objGrid.TextMatrix(objGridDist.iLinhasExistentes + 1, iGrid_ItemNFDist_Col) = objItemNF.iItem
            
            lErro = Mascara_RetornaProdutoTela(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 86357
 
            If objReserva.iAlmoxarifado > 0 Then
                objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
    
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
                If lErro = 25056 Then gError 46182
                objReserva.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
            
            ElseIf Len(Trim(objReserva.sAlmoxarifado)) > 0 Then
            
                objAlmoxarifado.sNomeReduzido = objReserva.sAlmoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 35939
                If lErro = 25060 Then gError 35940
                objReserva.iAlmoxarifado = objAlmoxarifado.iCodigo
            End If
                
                        
            objGridDist.objGrid.TextMatrix(objGridDist.iLinhasExistentes + 1, iGrid_ProdutoAlmoxDist_Col) = sProdutoMascarado
            objGridDist.objGrid.TextMatrix(objGridDist.iLinhasExistentes + 1, iGrid_UMDist_Col) = objItemRomaneio.sUMEstoque
            objGridDist.objGrid.TextMatrix(objGridDist.iLinhasExistentes + 1, iGrid_AlmoxDist_Col) = objReserva.sAlmoxarifado
            objGridDist.objGrid.TextMatrix(objGridDist.iLinhasExistentes + 1, iGrid_QuantItemNFDist_Col) = Formata_Estoque((objItemRomaneio.dQuantidade - objItemRomaneio.dQuantCancelada) * dFator2)
            objGridDist.objGrid.TextMatrix(objGridDist.iLinhasExistentes + 1, iGrid_QuantDist_Col) = Formata_Estoque(objReserva.dQuantidade * dFator)
            
            objGridDist.iLinhasExistentes = objGridDist.iLinhasExistentes + 1
            
        Next
               
    Next

    Atualiza_Grid_Distribuicao = SUCESSO
     
    Exit Function
    
Erro_Atualiza_Grid_Distribuicao:

    Atualiza_Grid_Distribuicao = gErr

    Select Case gErr
    
        Case 46181, 86357, 42766, 35939, 86357, 42764, 42765
        
        Case 46182, 35940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151577)
            
    End Select
    
    Exit Function

End Function

Public Sub ProdutoAlmoxDist_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoAlmoxDist_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDist)

End Sub

Public Sub ProdutoAlmoxDist_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDist)

End Sub

Public Sub ProdutoAlmoxDist_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDist.objControle = ProdutoAlmoxDist
    lErro = Grid_Campo_Libera_Foco(objGridDist)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ProdutoAlmox(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objProduto2 As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dFator As Double
Dim iItem As Integer
Dim iItem_Atual As Integer
Dim sProdutoMascarado As String 'Inserido por Wagner

On Error GoTo Erro_Saida_Celula_ProdutoAlmox

    Set objGridInt.objControle = ProdutoAlmoxDist

    If ProdutoAlmoxDist <> objGridInt.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_ProdutoAlmoxDist_Col) Then
                        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", ProdutoAlmoxDist.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 89471
                    
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
                    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89472
            
         'ERRO - Wagner
        '######################
        'Não lê produto com Grade - Acerto feito no carregamento da combo de produto
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 89473
                
        'descobre qual o item atual
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col))) > 0 Then
            iItem_Atual = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFDist_Col))
        End If
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iItem_Atual, objTela.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 89471

        'Lê os demais atributos do Produto
        objProduto2.sCodigo = sProdutoFormatado
                    
        lErro = CF("Produto_Le", objProduto2)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89472
                
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoAlmoxDist_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem_Atual, objTela.iGrid_Produto_Col)
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
            
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantDist_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_AlmoxDist_Col) = ""
            
        If objProduto2.iKitVendaComp = MARCADO Then
            dFator = 1
        Else
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTela.objGridItens.objGrid.TextMatrix(iItem_Atual, objTela.iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 89474
        End If
        
        For iItem = 1 To objTela.gobjnfiscal.ColItensNF(iItem_Atual).colItensRomaneioGrade.Count
            If objTela.gobjnfiscal.ColItensNF(iItem_Atual).colItensRomaneioGrade(iItem).sProduto = objProduto.sCodigo Then Exit For
        Next
        
        'ERRO - Wagner
        '######################
        'ColItensRomaneioGrade Vazia
        objGridInt.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_QuantItemNFDist_Col) = Formata_Estoque(objTela.gobjnfiscal.ColItensNF(iItem_Atual).colItensRomaneioGrade(iItem).dQuantidade * dFator)
        
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    objGridInt.objGrid.TextMatrix(objGridDist.objGrid.Row, iGrid_ProdutoAlmoxDist_Col) = ProdutoAlmoxDist.Text
           
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89487

    Saida_Celula_ProdutoAlmox = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoAlmox:

    Saida_Celula_ProdutoAlmox = gErr

    Select Case gErr
    
        Case 89473 'Inserido por Wagner - ERRO
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89487
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151578)

    End Select

    Exit Function

End Function

Function Move_DistribuicaoGrade_Memoria(objItemNF As ClassItemNF) As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim lErro As Long

On Error GoTo Erro_Move_DistribuicaoGrade_Memoria

    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
        Set objItemRomaneio.colLocalizacao = New Collection
    Next

    For iIndice = 1 To objGridDist.iLinhasExistentes

        If StrParaInt(objGridDist.objGrid.TextMatrix(iIndice, iGrid_ItemNFDist_Col)) = objItemNF.iItem Then

            'Formata o Produto para o BD
            lErro = CF("Produto_Formata", objGridDist.objGrid.TextMatrix(iIndice, iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 89471
            
            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                
                If objItemRomaneio.sProduto = sProdutoFormatado Then
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.sAlmoxarifado = objGridDist.objGrid.TextMatrix(iIndice, iGrid_AlmoxDist_Col)
                    objReservaItem.dQuantidade = StrParaDbl(objGridDist.objGrid.TextMatrix(iIndice, iGrid_QuantDist_Col))
                    
                    objItemRomaneio.colLocalizacao.Add objReservaItem
                    
                End If
                       
            Next
        End If
    Next

    Move_DistribuicaoGrade_Memoria = SUCESSO

    Exit Function

Erro_Move_DistribuicaoGrade_Memoria:

    Move_DistribuicaoGrade_Memoria = gErr
    
    Select Case gErr
    
        Case 89471
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151579)
            
    End Select

End Function

Private Function AlocaAlmoxarifadoPradrao_Grade(iLinhaGridItens As Integer, dQuantAlocar As Double, objProduto As ClassProduto, objItemRomaneio As ClassItemRomaneioGrade) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemNF As Object
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As Object
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim iCodigo As Integer
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_AlocaAlmoxarifadoPradrao_Grade

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 89507
    
    'Se não encontrou ==> não faz a alocação automática
    If iAlmoxarifado = 0 Then gError 89508

    objAlmoxarifado.iCodigo = iAlmoxarifado
    
    'Lê o Aloxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> 25056 And lErro <> SUCESSO Then gError 89509
    
    If lErro = 25056 Then gError 89510
    
    Set objReservaItem = New ClassReservaItem
    
    objReservaItem.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
    objReservaItem.dQuantidade = Formata_Estoque(dQuantAlocar)
    objReservaItem.iAlmoxarifado = objAlmoxarifado.iCodigo

    Set objItemRomaneio.colLocalizacao = New Collection
    objItemRomaneio.colLocalizacao.Add objReservaItem
    
    AlocaAlmoxarifadoPradrao_Grade = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao_Grade:

    AlocaAlmoxarifadoPradrao_Grade = gErr

    Select Case gErr

        Case 89506, 89507, 89508, 89509

        Case 89510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151580)

    End Select

    Exit Function

End Function

Private Function Possui_Grade() As Boolean

    Select Case objTela.Name
    
        Case NOME_TELA_PEDIDOVENDA, NOME_TELA_PEDIDOVENDACONSULTA, NOME_TELA_ORCAMENTOVENDA, NOME_TELA_NFISCALFATURAPEDIDO, NOME_TELA_NFISCALPEDIDO, NOME_TELA_NFISCAL, NOME_TELA_NFISCALFATURA, NOME_TELA_RECEBMATERIALC, NOME_TELA_RECEBMATERIALF, NOME_TELA_NFISCALENTRADA, NOME_TELA_NFISCALFATENTRADA, NOME_TELA_NFISCALENTREM, NOME_TELA_NFISCALENTDEV, NOME_TELA_NFISCALREM, NOME_TELA_NFISCALDEV
            Possui_Grade = True
        
        Case Else
            Possui_Grade = False
    
    End Select
    
    Exit Function

End Function

'###########################################################
'Inserido por Wagner
Public Function AlocaAlmoxarifadoPadraoTela(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, iAlmoxarifadoRetorno As Integer) As Long

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlocaAlmoxarifadoPadraoTela

    'Tenta Ler o Almoxarifado padrão para essa tela
    objEstoqueProduto.sProduto = sProduto
    objEstoqueProduto.iAlmoxarifado = iAlmoxarifadoPadraoTela
    
    'Le os estoques desse produto nesse almoxarifado
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 141069
    
    'Não encotrou
    If lErro <> SUCESSO Then
   
        'Busca o Almoxarifado padrão para o produto
        lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoRetorno)
        If lErro <> SUCESSO And lErro <> 23796 Then gError 141070
        
        'Se não encontrou
        If iAlmoxarifadoRetorno = 0 Then gError 141071
        
    Else
    
        iAlmoxarifadoRetorno = iAlmoxarifadoPadraoTela
        
    End If

    Exit Function

Erro_AlocaAlmoxarifadoPadraoTela:

    AlocaAlmoxarifadoPadraoTela = gErr

    Select Case gErr
    
        Case 141069 To 141071

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151581)

    End Select

    Exit Function

End Function
'###########################################################


Public Function Distribuicao_Processa_TEC(ByVal objProduto As ClassProduto, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim sUM As String
Dim dFator As Double

On Error GoTo Erro_Distribuicao_Processa_TEC

    'Remove as alocações do grid de alocações
    Call Remove_Distribuicao(objTela.objGridItens.objGrid.Row)

'    'Formata o Produto
'    lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(objTela.objGridItens.objGrid.Row, objTela.iGrid_Produto_Col), sProduto, iPreenchido)
'    If lErro <> SUCESSO Then gError 89502
'
'    objProduto.sCodigo = sProduto
'
'    'Lê o produto
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 89503
'
'    'se o produto não está cadastrado ==> erro
'    If lErro <> SUCESSO Then gError 89504
'
'    'recolhe a quantidade do grid de itens
'    dQuantidade = StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(objTela.objGridItens.objGrid.Row, objTela.iGrid_Quantidade_Col))

    'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 Then
    
'        'recolhe a UM do ItemNF
'        sUM = objTela.objGridItens.objGrid.TextMatrix(objTela.objGridItens.objGrid.Row, objTela.iGrid_UnidadeMed_Col)
'
'        'Faz a conversão da UM da tela para a UM de estoque
'        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 89506
'
'        'Converte a quantidade para a UM de estoque
'        dQuantidade = dQuantidade * dFator

        'Tenta Alocar o produto no Almoxarifado padrão
        lErro = AlocaAlmoxarifadoPradrao_TEC(objTela.objGridItens.objGrid.Row, dQuantidade, objProduto)
        If lErro <> SUCESSO Then gError 89505
    
    End If

    Distribuicao_Processa_TEC = SUCESSO

    Exit Function

Erro_Distribuicao_Processa_TEC:

    Distribuicao_Processa_TEC = gErr

    Select Case gErr

        Case 89502, 89503, 89505

        Case 89504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151572)

    End Select

    Exit Function

End Function

Private Function AlocaAlmoxarifadoPradrao_TEC(iLinhaGridItens As Integer, dQuantAlocar As Double, objProduto As ClassProduto) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemNF As Object
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As Object
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim iCodigo As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_AlocaAlmoxarifadoPradrao_TEC

    '##########################################################
    'Alterado por Wagner 24/11/05
'    'Busca o Almoxarifado padrão
'    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
'    If lErro <> SUCESSO And lErro <> 23796 Then gError 89507
'
'    'Se não encontrou ==> não faz a alocação automática
'    If iAlmoxarifado = 0 Then gError 89508

    lErro = AlocaAlmoxarifadoPadraoTela(giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO Then gError 199161
    '##########################################################

    objAlmoxarifado.iCodigo = iAlmoxarifado
    
    'Lê o Aloxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> 25056 And lErro <> SUCESSO Then gError 199162
    
    If lErro = 25056 Then gError 199163
    
    iLinhas = objGridDist.iLinhasExistentes

    iLinhas = iLinhas + 1
    
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 199164
    
    
    'Coloca no Grid alocações a alocação para o almoxarifado padrão
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_ItemNFDist_Col) = iLinhaGridItens
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_ProdutoAlmoxDist_Col) = sProdutoMascarado
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_AlmoxDist_Col) = objAlmoxarifado.sNomeReduzido
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_QuantDist_Col) = Formata_Estoque(dQuantAlocar)
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantAlocar)
    objGridDist.objGrid.TextMatrix(iLinhas, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
    
    'atualiza o numero de linhas existentes no grid de alocações
    objGridDist.iLinhasExistentes = iLinhas

    AlocaAlmoxarifadoPradrao_TEC = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao_TEC:

    AlocaAlmoxarifadoPradrao_TEC = gErr

    Select Case gErr

        Case 199161, 199162, 199164

        Case 199163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199165)

    End Select

    Exit Function

End Function

Public Function Preenche_GridDistribuicao_TEC(objNFiscal As Object) As Long
'Preenche o Grid com as Alocações da Nota Fiscal

Dim objItemAloc As Object
Dim iIndice As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridDistribuicao_TEC
    
    'Limpa o grid de alocações
    Call Grid_Limpa(objGridDist)

    
    'Para cada item da NF
    For iIndice = 1 To objNFiscal.ColItensNF.Count
                
        'Para cada alocação do Item de NF
        For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).ColAlocacoes

            objGridDist.iLinhasExistentes = objGridDist.iLinhasExistentes + 1
        
            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 51081
            If lErro <> SUCESSO Then gError 51082
        
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 51083
            
            'Coloca os dados da alocação na tela
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ItemNFDist_Col) = objNFiscal.ColItensNF(iIndice).iItem
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ProdutoAlmoxDist_Col) = objTela.objGridItens.objGrid.TextMatrix(objNFiscal.ColItensNF(iIndice).iItem, objTela.iGrid_Produto_Col)
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_AlmoxDist_Col) = objItemAloc.sAlmoxarifado
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantDist_Col) = Formata_Estoque(objItemAloc.dQuantidade)
            dQuantidade = StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_Quantidade_Col)) * dFator
            
            '#############################################################################
            'Verifica se o Formata estoque está arredondando o valor da quantidade p\ menos
'            If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
'
'                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                dQuantidade = dQuantidade + dAcrescimo
'
'            End If
            dQuantidade = Arredonda_Estoque(dQuantidade)
            '#############################################################################
            
            GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantidade)
            
            
            If objItemAloc.sProduto <> objNFiscal.ColItensNF(iIndice).sProduto Then
            
                objProduto.sCodigo = objItemAloc.sProduto
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 199172
            
                lErro = Mascara_RetornaProdutoTela(objItemAloc.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 199173

                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_ProdutoAlmoxDist_Col) = sProdutoMascarado
                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantDist_Col) = Formata_Estoque(objItemAloc.dQuantidade)
                GridDist.TextMatrix(objGridDist.iLinhasExistentes, iGrid_QuantItemNFDist_Col) = Formata_Estoque(objItemAloc.dQuantidade)
                
            End If
        Next

        
        If objNFiscal.ColItensNF(iIndice).iPossuiGrade Then Call Atualiza_Grid_Distribuicao(objNFiscal.ColItensNF(iIndice))

    Next

    Preenche_GridDistribuicao_TEC = SUCESSO

    Exit Function

Erro_Preenche_GridDistribuicao_TEC:

    Preenche_GridDistribuicao_TEC = gErr
    
    Select Case gErr
    
        Case 51081, 51083, 199172, 199173
        
        Case 51082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoAlmoxDist.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151568)
            
    End Select

End Function

