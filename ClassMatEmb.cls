VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMatEmb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gdValorPesoBrutoSemEmb As Double 'Inserido por Wagner
Dim gdValorPesoLiqSemEmb As Double 'Inserido por Wagner

Public objTela As Object

Public objGridEmb As AdmGrid
Public iGrid_ItemEmb_Col As Integer
Public iGrid_ProdutoEmb_Col As Integer
Public iGrid_Embalagem_Col As Integer
Public iGrid_UMEmb_Col As Integer
Public iGrid_Capacidade_Col As Integer
Public iGrid_QuantEmb_Col As Integer
Public iGrid_QuantProduto_Col As Integer
Public iGrid_PesoLiq_Col As Integer
Public iGrid_PesoBruto_Col As Integer

Public bTela As Boolean

Dim giTipoDocOriginal As Integer

'Janaina
'Declaração utilizada para evento BotaoEmbalagens_Click
Private WithEvents objEventoEmbalagens As AdmEvento
Attribute objEventoEmbalagens.VB_VarHelpID = -1

Private objUserControl As Object

'##################################################
'Inserido por Wagner
Const FORMATO_PESO = "#,##0.00#"
Const NUM_CASAS_DECIMAIS_PESO = 3
'##################################################

Private Sub Class_Initialize()
    Set objGridEmb = New AdmGrid
    bTela = False
End Sub

Public Property Get ItemEmb() As Object
     Set ItemEmb = objUserControl.Controls("ItemEmb")
End Property

Public Property Get ProdutoEmb() As Object
     Set ProdutoEmb = objUserControl.Controls("ProdutoEmb")
End Property

Public Property Get Embalagem() As Object
     Set Embalagem = objUserControl.Controls("Embalagem")
End Property

Public Property Get UMEmb() As Object
     Set UMEmb = objUserControl.Controls("UmEmb")
End Property

Public Property Get Capacidade() As Object
     Set Capacidade = objUserControl.Controls("Capacidade")
End Property

Public Property Get QuantEmb() As Object
     Set QuantEmb = objUserControl.Controls("QuantEmb")
End Property

Public Property Get QuantProduto() As Object
     Set QuantProduto = objUserControl.Controls("QuantProduto")
End Property

Public Property Get PesoLiq() As Object
     Set PesoLiq = objUserControl.Controls("PesoLiq")
End Property

Public Property Get PesoBrutoEmb() As Object
     Set PesoBrutoEmb = objUserControl.Controls("PesoBrutoEmb")
End Property

Public Property Get BotaoEmbalagens() As Object
     Set BotaoEmbalagens = objUserControl.Controls("BotaoEmbalagens")
End Property

Public Property Get CalculoAuto() As Object
     Set CalculoAuto = objUserControl.Controls("CalculoAuto")
End Property

Public Property Get GridEmb() As Object
     Set GridEmb = objUserControl.Controls("GridEmb")
End Property

Public Property Get QuantEmbTotal() As Object
     Set QuantEmbTotal = objUserControl.Controls("QuantEmbTotal")
End Property
'Janaina
Public Property Get DescProduto() As Object
     Set DescProduto = objUserControl.Controls("DescProduto")
End Property
'Janaina
Public Property Get UMProduto() As Object
     Set UMProduto = objUserControl.Controls("UMProduto")
End Property
'Janaina
Public Property Get ProdutoQuant() As Object
     Set ProdutoQuant = objUserControl.Controls("ProdutoQuant")
End Property
''Janaina
'Public Property Get QuantAlocar() As Object
'     Set QuantAlocar = objUserControl.Controls("QuantAlocar")
'End Property

Public Property Get PesoLiqTotal() As Object
     Set PesoLiqTotal = objUserControl.Controls("PesoLiqTotal")
End Property

Public Property Get PesoBrutoTotal() As Object
     Set PesoBrutoTotal = objUserControl.Controls("PesoBrutoTotal")
End Property

Public Function Inicializa_GridEmb() As Long
'Inicializa o Grid de Embalagens

Dim objGridInt As AdmGrid
Dim iIndice As Integer

    Set objGridInt = objGridEmb

    Set objGridInt.objForm = objTela

    If bTela Then
        Set objUserControl = objTela
    Else
        Set objUserControl = objTela.objUserControl
    End If

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    
    If giTipoDocOriginal = DOCORIGEM_NF Then
        objGridInt.colColuna.Add ("Item NF")
    Else
        objGridInt.colColuna.Add ("Item PV")
    End If
    
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Embalagem")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Capacidade")
    objGridInt.colColuna.Add ("Quant.Emb.")
    objGridInt.colColuna.Add ("Quant.Produto")
    objGridInt.colColuna.Add ("Peso Liq.")
    objGridInt.colColuna.Add ("Peso Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemEmb.Name)
    objGridInt.colCampo.Add (ProdutoEmb.Name)
    objGridInt.colCampo.Add (Embalagem.Name)
    objGridInt.colCampo.Add (UMEmb.Name)
    objGridInt.colCampo.Add (Capacidade.Name)
    objGridInt.colCampo.Add (QuantEmb.Name)
    objGridInt.colCampo.Add (QuantProduto.Name)
    objGridInt.colCampo.Add (PesoLiq.Name)
    objGridInt.colCampo.Add (PesoBrutoEmb.Name)

    'Colunas da Grid
    iGrid_ItemEmb_Col = 1
    iGrid_ProdutoEmb_Col = 2
    iGrid_Embalagem_Col = 3
    iGrid_UMEmb_Col = 4
    iGrid_Capacidade_Col = 5
    iGrid_QuantEmb_Col = 6
    iGrid_QuantProduto_Col = 7
    iGrid_PesoLiq_Col = 8
    iGrid_PesoBruto_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridEmb

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_EMBALAGENS + 1

    'Linhas visíveis do grid
    'objGridInt.iLinhasVisiveis = 6
    If GridEmb.Height > 5200 Then
        objGridInt.iLinhasVisiveis = 20
    Else
        objGridInt.iLinhasVisiveis = 6
    End If
    
    'Largura da primeira coluna
    GridEmb.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

'    'Quant.Emb.
'    QuantEmbTotal.Top = GridEmb.Top + GridEmb.Height
'    QuantEmbTotal.Left = GridEmb.Left
'    For iIndice = 0 To iGrid_QuantEmb_Col - 1
'        QuantEmbTotal.Left = QuantEmbTotal.Left + GridEmb.ColWidth(iIndice) + GridEmb.GridLineWidth
'    Next
'
'    QuantEmbTotal.Top = QuantEmbTotal.Top + (QuantEmbTotal.Height / 2) - (QuantEmbTotal.Height / 2)
'    QuantEmbTotal.Left = QuantEmbTotal.Left - QuantEmbTotal.Width
'
'    'Peso Liquido
'    PesoLiqTotal.Top = GridEmb.Top + GridEmb.Height
'    PesoLiqTotal.Left = GridEmb.Left
'    For iIndice = 0 To iGrid_PesoLiq_Col - 1
'        PesoLiqTotal.Left = PesoLiqTotal.Left + GridEmb.ColWidth(iIndice) + GridEmb.GridLineWidth
'    Next
'
'    PesoLiqTotal.Top = PesoLiqTotal.Top + (PesoLiqTotal.Height / 2) - (PesoLiqTotal.Height / 2)
'    PesoLiqTotal.Left = PesoLiqTotal.Left - PesoLiqTotal.Width
'
'    'Peso Bruto
'    PesoBrutoTotal.Top = GridEmb.Top + GridEmb.Height
'    PesoBrutoTotal.Left = GridEmb.Left
'    For iIndice = 0 To iGrid_PesoLiq_Col - 1
'        PesoBrutoTotal.Left = PesoBrutoTotal.Left + GridEmb.ColWidth(iIndice) + GridEmb.GridLineWidth
'    Next
'
'    PesoBrutoTotal.Top = PesoBrutoTotal.Top + (PesoBrutoTotal.Height / 2) - (PesoBrutoTotal.Height / 2)
'    PesoBrutoTotal.Left = PesoBrutoTotal.Left - PesoBrutoTotal.Width

    Inicializa_GridEmb = SUCESSO

    Exit Function

End Function

Public Function Inicializa(Optional ByVal iTipoDocOriginal As Integer = DOCORIGEM_PV) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa

    giTipoDocOriginal = iTipoDocOriginal

    'Inicializa o grid de Embalagens
    lErro = Inicializa_GridEmb()
    If lErro <> SUCESSO Then gError 96130
    
    'Inicializa a mascara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoEmb)
    If lErro <> SUCESSO Then gError 96131

    'Janaina
    'Inicializa o objEventoEmbalagens
    Set objEventoEmbalagens = New AdmEvento
    'Janaina
    
    Inicializa = SUCESSO
    
Exit Function

Erro_Inicializa:

    Select Case gErr

        Case 96130, 96131
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151582)

    End Select

Exit Function

End Function

Public Sub GridEmb_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridEmb, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEmb, objTela.iAlterado)
    End If

End Sub

Public Sub GridEmb_EnterCell()

    Call Grid_Entrada_Celula(objGridEmb, objTela.iAlterado)

End Sub

Public Sub GridEmb_GotFocus()

    Call Grid_Recebe_Foco(objGridEmb)

End Sub

Public Sub GridEmb_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridEmb, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEmb, objTela.iAlterado)
    End If

End Sub

Public Sub GridEmb_LeaveCell()

    Call objTela.Saida_Celula(objGridEmb)

End Sub

Public Sub GridEmb_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridEmb)

End Sub
'Janaina
Public Sub GridEmb_RowColChange()

Dim objProduto As New ClassProduto
Dim iLinhaAnterior As Integer
Dim sProduto As String
Dim sUM As String, sUMEmbalagem As String
Dim dSomaProduto As Double
Dim dQuantAlocar As Double
Dim dFator As Double
Dim iPreenchido As Integer
Dim iItem As Integer
Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_GridEmb_RowColChange

     Call Grid_RowColChange(objGridEmb)
    
    If objGridEmb.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then
    
        'se a linha anterior for diferente entao preenche a decricao do produto selecionado
        If (GridEmb.Row <> iLinhaAnterior) Then
            
            If GridEmb.Row > 0 Then
            
                lErro = CF("Produto_Formata", GridEmb.TextMatrix(GridEmb.Row, iGrid_ProdutoEmb_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then gError 96158
    
                objProduto.sCodigo = sProduto
    
                iItem = StrParaInt(GridEmb.TextMatrix(GridEmb.Row, iGrid_ItemEmb_Col))
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 96157
    
                'Se não achou o Produto --> erro
                If lErro = 28030 Then gError 96159
                    
                DescProduto.Caption = objProduto.sDescricao
                
                '###########################
                'Inserido por Wagner
                If iItem <> 0 Then
                '###########################
                    UMProduto.Caption = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_UnidadeMed_Col)
                    
                    If giTipoDocOriginal = DOCORIGEM_NF Then
                        ProdutoQuant.Caption = Formata_Estoque(StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Quantidade_Col)))
                    Else
                        ProdutoQuant.Caption = Formata_Estoque(Abs(StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_QuantPedida_Col)) - StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_QuantCancel_Col))))
                    End If
                '###########################
                End If
                '###########################
            
'                'Janaina - Testar !!
'                For iLinha = 1 To objGridEmb.iLinhasExistentes
'
'                    If iItem = GridEmb.TextMatrix(iLinha, iGrid_ItemEmb_Col) Then
'                      dSomaProduto = dSomaProduto + GridEmb.TextMatrix(iLinha, iGrid_QuantProduto_Col)
'                    End If
'
'                Next
'
'                sUM = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_UnidadeMed_Col)
'                sUMEmbalagem = GridEmb.TextMatrix(GridEmb.Row, iGrid_UMEmb_Col)
'
'                If sUM <> sUMEmbalagem Then
'
'                    'Converte da UM do GridItens para UM GridEmb
'                    lErro = CF("UM_Conversao_Trans",objProduto.iClasseUM, sUM, sUMEmbalagem, dFator)
'                    If lErro <> SUCESSO Then gError 96117
'
'                End If
'
'                dQuantAlocar = (Abs(StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_QuantPedida_Col)) - StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_QuantCancel_Col)))) * dFator
'                QuantAlocar.Caption = Formata_Estoque(dSomaProduto - dQuantAlocar)
'                'Janaina
'
             End If
            
            'Guarda a Linha corrente
            iLinhaAnterior = GridEmb.Row
                
        End If
    
    End If

    Exit Sub

Erro_GridEmb_RowColChange:

    Select Case gErr
        
        Case 96157, 96158, 96159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 151583)

    End Select

    Exit Sub

End Sub
'Janaina

Public Sub GridEmb_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer, lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_GridEmb_KeyDown

    iLinhasExistentesAnterior = objGridEmb.iLinhasExistentes
    iItemAtual = GridEmb.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridEmb)
    
    'Janaina
    If objGridEmb.iLinhasExistentes < iLinhasExistentesAnterior Then
        
        CalculoAuto.Value = vbUnchecked
    
        '##########################
        'Inserido por Wagner
        'Para poder considerar produtos sem embalagens no grid
        Call objTela.Pesos_Recalcula
        '##########################
    
'        Call AtualizaTotais
'
'        Call Preenche_Pesos
        
    End If
    'Janaina

    Exit Sub

Erro_GridEmb_KeyDown:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151584)

    End Select

    Exit Sub

End Sub

Public Sub GridEmb_Scroll()

    Call Grid_Scroll(objGridEmb)

End Sub

Public Function Preenche_GridEmbalagens(colEmbProdItem As Collection) As Long
'Preenche o Grid de Embalagens a partir dos dados de colEmbProdItem

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objEmbalagensItem As ClassEmbalagensItem
Dim objEmbalagem As New ClassEmbalagem
Dim objEmbProdItem As ClassEmbProdItem
Dim iIndice1 As Integer
Dim iQuantPedidaProd As Integer
Dim dFator As Double
Dim sProdutoEnxuto As String
Dim dPesoEmbalagem As Double

Dim dPesoLiq As Double, dPesoBruto As Double

On Error GoTo Erro_Preenche_GridEmbalagens

    'Limpa o grid de embalagens
    Call Grid_Limpa(objGridEmb)
   
    'Janaina
    CalculoAuto.Value = vbUnchecked
    'Janaina
    
    iIndice1 = 0
                
     'Para cada embalagem do Item
     For Each objEmbProdItem In colEmbProdItem

         iIndice1 = iIndice1 + 1
    
         objEmbalagem.iCodigo = objEmbProdItem.iEmbalagem
         
         'Lê a embalagem à partir do Codigo
         lErro = CF("Embalagem_Le", objEmbalagem)
         If lErro <> SUCESSO And lErro <> 25060 Then gError 35939
        
        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objEmbProdItem.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 35943

        ProdutoEmb.PromptInclude = False
        ProdutoEmb.Text = sProdutoEnxuto
        ProdutoEmb.PromptInclude = True

         GridEmb.TextMatrix(iIndice1, iGrid_ItemEmb_Col) = objEmbProdItem.iItem
         GridEmb.TextMatrix(iIndice1, iGrid_ProdutoEmb_Col) = ProdutoEmb.Text
         GridEmb.TextMatrix(iIndice1, iGrid_Embalagem_Col) = objEmbalagem.sSigla
         GridEmb.TextMatrix(iIndice1, iGrid_UMEmb_Col) = objEmbProdItem.sUMEmbalagem
         GridEmb.TextMatrix(iIndice1, iGrid_Capacidade_Col) = Formata_Estoque(objEmbProdItem.dCapacidade)
         GridEmb.TextMatrix(iIndice1, iGrid_QuantEmb_Col) = Formata_Estoque(objEmbProdItem.dQuantEmb)
         GridEmb.TextMatrix(iIndice1, iGrid_QuantProduto_Col) = Formata_Estoque(objEmbProdItem.dQuantProduto)
         
        lErro = CalculaPesos(objEmbProdItem.dQuantEmb, objEmbProdItem.dPesoBruto - objEmbProdItem.dPesoLiqTotal, objEmbProdItem.sUMPeso, objEmbProdItem.dQuantProduto, objEmbProdItem.dCapacidade, objEmbProdItem.dPesoLiqTotal, dPesoLiq, dPesoBruto)
        If lErro <> SUCESSO Then gError 96117
         
         GridEmb.TextMatrix(iIndice1, iGrid_PesoLiq_Col) = Format(dPesoLiq, FORMATO_PESO) 'Alterado por Wagner
         GridEmb.TextMatrix(iIndice1, iGrid_PesoBruto_Col) = Format(dPesoBruto, FORMATO_PESO) 'Alterado por Wagner

     Next
    
    'Inicializa o número de linhas existentes no grid
    objGridEmb.iLinhasExistentes = iIndice1
    
    Call AtualizaTotais

    Preenche_GridEmbalagens = SUCESSO

    Exit Function

Erro_Preenche_GridEmbalagens:

    Preenche_GridEmbalagens = gErr

    Select Case gErr

        Case 51081, 51083, 96117

        Case 51082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151585)

    End Select

End Function

Public Sub Limpa_Tela_Embalagem()
'Limpa a tela de embalagem

    'Limpa o Grid de Embalagem
    Call Grid_Limpa(objGridEmb)
    
    'Limpa os totais
    QuantEmbTotal.Caption = ""
    PesoLiqTotal.Caption = ""
    PesoBrutoTotal.Caption = ""
    
    'Limpa a descrição do produto
    DescProduto.Caption = ""
    UMProduto.Caption = ""
    ProdutoQuant.Caption = ""
    'QuantAlocar.Caption = ""
    
End Sub

Public Function Move_GridEmb_Memoria(colEmbalagensItens As Collection, iTipoDocOrigem As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objEmbalagensItens As ClassEmbalagensItem
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Move_GridEmb_Memoria

    'Percorre o grid de embalagens
    For iIndice = 1 To objGridEmb.iLinhasExistentes
    
        Set objEmbalagensItens = New ClassEmbalagensItem
        Set objEmbalagem = New ClassEmbalagem
                
        'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
        faz uma nova leitura em busca do codigo
        objEmbalagem.sSigla = GridEmb.TextMatrix(iIndice, iGrid_Embalagem_Col)
        lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 95088 Then gError 95460
        
        'Se nao achou => ERRO
        If lErro = 95088 Then gError 95461
        
        objEmbalagensItens.iEmbalagem = objEmbalagem.iCodigo
        objEmbalagensItens.iItem = CInt(GridEmb.TextMatrix(iIndice, iGrid_ItemEmb_Col))
        objEmbalagensItens.iTipoDocOrigem = iTipoDocOrigem
        objEmbalagensItens.dQuantEmb = StrParaDbl(GridEmb.TextMatrix(iIndice, iGrid_QuantEmb_Col))
        objEmbalagensItens.dQuantProduto = StrParaDbl(GridEmb.TextMatrix(iIndice, iGrid_QuantProduto_Col))
        
        lErro = CF("Produto_Formata", GridEmb.TextMatrix(iIndice, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 96143

        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado

        objEmbalagensItens.sProduto = sProdutoFormatado

        'Adiciona a embalagens do grid na coleção de embalagens
        colEmbalagensItens.Add objEmbalagensItens

    Next
    

    Move_GridEmb_Memoria = SUCESSO

    Exit Function

Erro_Move_GridEmb_Memoria:

    Move_GridEmb_Memoria = gErr

    Select Case gErr

        Case 35939, 96143, 95460
        
        Case 95461
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151586)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ItemEmb(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iItem As Integer
Dim iItem_Atual As Integer

On Error GoTo Erro_Saida_Celula_ItemEmb

    Set objGridInt.objControle = ItemEmb

    'Verifica se o Produto esta preenchido
    If Len(Trim(ItemEmb.Text)) = 0 And Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemEmb_Col))) > 0 Then gError 89468
    
    If Len(Trim(ItemEmb.Text)) > 0 Then

        iItem = CInt(ItemEmb.Text)
        
        If iItem < 1 Or iItem > objTela.objGridItens.iLinhasExistentes Then gError 89469

        'Se o produto não está preenchido ==> erro
        If Len(Trim(objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col))) = 0 Then gError 89470
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 89471
                    
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89472
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 89473
        
        '#########################################################
        'ALTERADO POR WAGNER
        lErro = CF("ProdutoEmbalagem_Le_Produto", objProduto)
        If lErro <> SUCESSO And lErro <> 95016 Then gError 129960
        
        If lErro = 95016 Then gError 129961
        '#########################################################
       
        'descobre qual o item atual
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemEmb_Col))) > 0 Then
            iItem_Atual = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemEmb_Col))
        End If

        'se o item que se está preenchendo é diferente do item atual, ==> limpa os campos de não editaveis
        If iItem <> iItem_Atual Then
        
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantEmb_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantProduto_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Embalagem_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_PesoLiq_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_PesoBruto_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Capacidade_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UMEmb_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Capacidade_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col) = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col)
        End If
        
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            'Janaina
            CalculoAuto.Value = vbUnchecked
            'Janaina
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89493

    Saida_Celula_ItemEmb = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemEmb:

    Saida_Celula_ItemEmb = gErr

    Select Case gErr
        
        Case 89468, 89469
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMDISTRIBUICAO_NAO_ITEMNF", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 89470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID_ITENS", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89471, 89472, 89474, 89493, 129960
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 89473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        '############################################################
        'INSERIDO POR WAGNER
        Case 129961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOEMBALAGEM_INEXISTENTE ", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        '############################################################
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151587)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_QuantEmb(objGridInt As AdmGrid) As Long

Dim lErro As Long, dQuantEmb As Double, dQuantProduto As Double

On Error GoTo Erro_Saida_Celula_QuantEmb
    
    Set objGridInt.objControle = QuantEmb

    'Se a quantidade de embalagens foi preenchida
    If Len(Trim(QuantEmb.clipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Inteiro_Critica(QuantEmb.Text)
        If lErro <> SUCESSO Then gError 89485
       
        dQuantEmb = StrParaDbl(QuantEmb.Text)
            
        'Janaina
        'se alterar a quantidade de embalagens, desmarca a Check
        If dQuantEmb <> StrParaDbl(GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantEmb_Col)) Then
            
            CalculoAuto.Value = vbUnchecked
            
            If dQuantEmb <> 0 Then
            
                GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantProduto_Col) = Formata_Estoque(dQuantEmb * StrParaDbl(GridEmb.TextMatrix(GridEmb.Row, iGrid_Capacidade_Col)))
                GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantEmb_Col) = Formata_Estoque(dQuantEmb)
                
                lErro = CalculaPesosLinha(GridEmb.Row, dQuantEmb, StrParaDbl(GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantProduto_Col)))
                If lErro <> SUCESSO Then gError 81727
            
            End If
            
        End If
        'Janaina
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89487

    Saida_Celula_QuantEmb = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantEmb:

    Saida_Celula_QuantEmb = gErr

    Select Case gErr

        Case 89485, 89487, 81727
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTDIST_MAIOR_QUANTITEMNF", gErr, QuantEmb.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantEmb_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151588)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_QuantProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long, dQuantProduto As Double

On Error GoTo Erro_Saida_Celula_QuantProduto

    Set objGridInt.objControle = QuantProduto

    'Se a quantidade de produto foi preenchida
    If Len(Trim(QuantProduto.clipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantProduto.Text)
        If lErro <> SUCESSO Then gError 89485
       
        QuantProduto.Text = Formata_Estoque(QuantProduto.Text)

        'Janaina
        'se alterar a quantidade de Produtos, desmarca a Check
        If QuantProduto.Text <> GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantProduto_Col) Then
            
            CalculoAuto.Value = vbUnchecked
            
            dQuantProduto = StrParaDbl(QuantProduto.Text)
            
            If dQuantProduto <> 0 Then
            
                GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantEmb_Col) = Arredonda_ParaCima(dQuantProduto / StrParaDbl(GridEmb.TextMatrix(GridEmb.Row, iGrid_Capacidade_Col)))
                GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantProduto_Col) = Formata_Estoque(dQuantProduto)
                
                lErro = CalculaPesosLinha(GridEmb.Row, GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantEmb_Col), dQuantProduto)
                If lErro <> SUCESSO Then gError 81728
            
            End If
        
        End If
        'Janaina

' Testar a quantidade do produto vs capacidade ?????????mario

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89487

    Saida_Celula_QuantProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantProduto:

    Saida_Celula_QuantProduto = gErr

    Select Case gErr

        Case 89485, 89487, 81728
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTDIST_MAIOR_QUANTITEMNF", gErr, QuantEmb.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantProduto_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151589)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_GridEmb() As Long

Dim lErro As Long
Dim objGridInt As AdmGrid

On Error GoTo Erro_Saida_Celula_GridEmb

    Set objGridInt = objGridEmb

    Select Case objGridInt.objGrid.Col
    
        Case iGrid_ItemEmb_Col
            lErro = Saida_Celula_ItemEmb(objGridInt)
            If lErro <> SUCESSO Then gError 89489
    
        Case iGrid_Embalagem_Col
            lErro = Saida_Celula_Embalagem(objGridInt)
            If lErro <> SUCESSO Then gError 89490
        
        Case iGrid_QuantEmb_Col
            lErro = Saida_Celula_QuantEmb(objGridInt)
            If lErro <> SUCESSO Then gError 89491
    
        Case iGrid_QuantProduto_Col
            lErro = Saida_Celula_QuantProduto(objGridInt)
            If lErro <> SUCESSO Then gError 89491
    
    End Select

    Saida_Celula_GridEmb = SUCESSO

    Exit Function

Erro_Saida_Celula_GridEmb:

    Saida_Celula_GridEmb = gErr

    Select Case gErr

        Case 89489, 89490, 89491

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151590)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Embalagem(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEmbalagem As New ClassEmbalagem
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim vbMsg As VbMsgBoxResult, dQuantProduto As Double, dQuantEmb As Double
Dim iLinha As Integer, embText As Object

On Error GoTo Erro_Saida_Celula_Embalagem

    Set objGridInt.objControle = Embalagem

    'Se a Embalagem está preenchido
    If Len(Trim(Embalagem.Text)) > 0 Then
       
        If Embalagem.Text <> GridEmb.TextMatrix(GridEmb.Row, iGrid_ProdutoEmb_Col) Then
        
            CalculoAuto.Value = vbUnchecked
            
            'Le os dados da embalagem
            Set embText = Embalagem
            lErro = CF("TP_Embalagem_Le_Grid", embText, objEmbalagem)
            If lErro <> SUCESSO Then gError 96123
    
            'Janaina
            'Preenche a UMEmbagem e a Capacidade
             objProdutoEmbalagem.iEmbalagem = objEmbalagem.iCodigo
                    
            lErro = CF("Produto_Formata", GridEmb.TextMatrix(GridEmb.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 96153
    
            objProdutoEmbalagem.sProduto = sProdutoFormatado
            
            lErro = CF("ProdutoEmbalagem_Le", objProdutoEmbalagem)
            If lErro <> SUCESSO And lErro <> 96156 Then gError 81728
            
            'Alteracao Daniel em 13/12/2001
            'Se nao está associado => limpa e gera erro
            If lErro = 96156 Then gError 96154
            
            'Preenche o grid
            Embalagem.Text = objEmbalagem.sSigla
            GridEmb.TextMatrix(GridEmb.Row, iGrid_Embalagem_Col) = objEmbalagem.sSigla
            GridEmb.TextMatrix(GridEmb.Row, iGrid_UMEmb_Col) = objProdutoEmbalagem.sUMEmbalagem
            GridEmb.TextMatrix(GridEmb.Row, iGrid_Capacidade_Col) = Formata_Estoque(objProdutoEmbalagem.dCapacidade)
            'Janaina
            
            'Se necessário cria uma nova linha no Grid
            If GridEmb.Row - GridEmb.FixedRows = objGridEmb.iLinhasExistentes Then objGridEmb.iLinhasExistentes = objGridEmb.iLinhasExistentes + 1
            
            'se já há qtde de produto recalcula a qtde de embalagens
            dQuantProduto = StrParaDbl(GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantProduto_Col))
            If dQuantProduto <> 0 Then
            
                dQuantEmb = Arredonda_ParaCima(dQuantProduto / objProdutoEmbalagem.dCapacidade)
            
            Else
            
                dQuantEmb = 0
            
            End If
            
            GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantEmb_Col) = Formata_Estoque(dQuantEmb)
            
            lErro = CalculaPesosLinha(GridEmb.Row, dQuantEmb, dQuantProduto)
            If lErro <> SUCESSO Then gError 81727
        
        End If
    
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96124

    Saida_Celula_Embalagem = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Embalagem:

    Saida_Celula_Embalagem = gErr

    Select Case gErr
       
        Case 96123, 96124, 81727
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 96153, 81728
        
        Case 96154
            Call Rotina_Erro(vbOKOnly, "AVISO_PRODUTOEMBALAGEM_INEXISTENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151591)

    End Select

    Exit Function

End Function

Private Function Soma_Valor_QuantEmb(objGridInt As AdmGrid) As Long
'Atualiza o total da Quant.Emb.

Dim iIndice As Integer
Dim dSomaQuantEmb As Double

    dSomaQuantEmb = 0
    
    'Loop no GridEmb
    For iIndice = 1 To objGridInt.iLinhasExistentes
            
        If StrParaDbl(GridEmb.TextMatrix(iIndice, iGrid_QuantEmb_Col)) > 0 Then
            
            'Acumula Valor em dSomaValor
            dSomaQuantEmb = dSomaQuantEmb + CDbl(GridEmb.TextMatrix(iIndice, iGrid_QuantEmb_Col))
            
        End If
            
    Next

    'Mostra na tela o Valor Total
    QuantEmbTotal.Caption = Format(dSomaQuantEmb, "Standard")

    Soma_Valor_QuantEmb = SUCESSO

    Exit Function
'Janaina
End Function
'Janaina
Private Function Soma_Valor_PesoLiq(objGridInt As AdmGrid) As Long
'Atualiza o total do Peso Liquido

Dim iIndice As Integer
Dim dSomaPesoLiq As Double

    dSomaPesoLiq = 0
    
    'Loop no GridEmb
    For iIndice = 1 To objGridInt.iLinhasExistentes
            
        If StrParaDbl(GridEmb.TextMatrix(iIndice, iGrid_PesoLiq_Col)) > 0 Then
            
            'Acumula Valor em dSomaPesoLiq
            dSomaPesoLiq = dSomaPesoLiq + CDbl(GridEmb.TextMatrix(iIndice, iGrid_PesoLiq_Col))
            
        End If

    Next

    'Mostra na tela o Valor Total
    PesoLiqTotal.Caption = Format(dSomaPesoLiq, FORMATO_PESO)

    Soma_Valor_PesoLiq = SUCESSO

    Exit Function
'Janaina
End Function
'Janaina
Private Function Soma_Valor_PesoBruto(objGridInt As AdmGrid) As Long
'Atualiza o total do Peso Bruto

Dim iIndice As Integer
Dim dSomaPesoBruto As Double

    dSomaPesoBruto = 0
    
    'Loop no GridEmb
    For iIndice = 1 To objGridInt.iLinhasExistentes
               
        If StrParaDbl(GridEmb.TextMatrix(iIndice, iGrid_PesoBruto_Col)) > 0 Then
            
            'Acumula Valor em dSomaPesoBruto
            dSomaPesoBruto = dSomaPesoBruto + CDbl(GridEmb.TextMatrix(iIndice, iGrid_PesoBruto_Col))
            
        End If
            
    Next

    'Mostra na tela o Valor Total
    PesoBrutoTotal.Caption = Format(dSomaPesoBruto, FORMATO_PESO)

    Soma_Valor_PesoBruto = SUCESSO

    Exit Function
'Janaina
End Function
'Janaina
Public Function Calcula_Item_GridEmb(ByVal dQuantAlocar As Double, ByVal iItem As Integer) As Long
'Calculo automatico das embalagens

Dim dQuantEmb As Double
Dim dQuantProduto As Double
Dim iAlteradoEmb As Integer
Dim iIndice1 As Integer
Dim lErro As Long
Dim dFator As Double
Dim dPesoLiq As Double
Dim dPesoBruto As Double, dCapacConv As Double
Dim dPesoEmbalagem As Double, sUMItemPV As String
Dim objProduto As New ClassProduto, sProdutoFormatado As String, iProdutoPreenchido As Integer
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem, sProdMasc As String
Dim bAchou As Boolean 'Inserido por Wagner
Dim iIndice As Integer 'Inserido por Wagner

On Error GoTo Erro_Calcula_Item_GridEmb

    '#####################################
    'Inserido por Wagner
    'Se for o primeiro item do Grid, então zera o Peso para poder recalcular
    If iItem = 1 Then
        PesoBrutoTotal.Caption = Format("0", FORMATO_PESO) 'Alterado por Wagner
        PesoLiqTotal.Caption = Format("0", FORMATO_PESO) 'Alterado por Wagner
        objTela.PesoBruto.Text = Format("0", FORMATO_PESO) 'Alterado por Wagner
        objTela.PesoLiquido.Text = Format("0", FORMATO_PESO) 'Alterado por Wagner
        gdValorPesoBrutoSemEmb = 0
        gdValorPesoLiqSemEmb = 0
    End If

    'Verifica se um produto tem embalagem ou não
    sProdMasc = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col)
            
    'Formata o Produto para o BD
    lErro = CF("Produto_Formata", sProdMasc, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 89471
        
    objProduto.sCodigo = sProdutoFormatado
        
    'Seleciona as embalagens
    lErro = CF("ProdutoEmbalagem_Le_Produto", objProduto)
    If lErro <> SUCESSO And lErro <> 95016 Then gError 132014
    
    'Se o Produto não tem embalagem
    If lErro <> SUCESSO Then
    
        'Incrementa o peso do Item (Com base no Produto)
        lErro = Pesos_Recalcula(dQuantAlocar, iItem)
        If lErro <> SUCESSO Then gError 132010
    
    Else
    'Se o produto tem embalagem
    '#####################################

        'Se o calculo for automatico
        If CalculoAuto.Value = vbChecked Then
        
            'Remove as embalagens do grid de embalagens
            Call Remove_Embalagem(iItem)
            
            iIndice1 = objGridEmb.iLinhasExistentes
           
    '        sProdMasc = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_Produto_Col)
    '
    '        'Formata o Produto para o BD
    '        lErro = CF("Produto_Formata", sProdMasc, sProdutoFormatado, iProdutoPreenchido)
    '        If lErro <> SUCESSO Then gError 89471
    '
            objProdutoEmbalagem.sProduto = objProduto.sCodigo
    
            'Seleciona embalagem padrao
            lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
            If lErro <> SUCESSO Then gError 96115

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 96116
            
            sUMItemPV = objTela.objGridItens.objGrid.TextMatrix(iItem, objTela.iGrid_UnidadeMed_Col)
            
            If sUMItemPV <> objProdutoEmbalagem.sUMEmbalagem Then
                
                'Converte da UM do GridItens para UM GridEmb
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItemPV, objProdutoEmbalagem.sUMEmbalagem, dFator)
                If lErro <> SUCESSO Then gError 96117
        
                dQuantAlocar = dQuantAlocar * dFator
           
            Else
                dFator = 1
            End If
    
            dQuantEmb = Fix(dQuantAlocar / objProdutoEmbalagem.dCapacidade)
            dQuantProduto = Round(objProdutoEmbalagem.dCapacidade * dQuantEmb, 4)
            dQuantAlocar = dQuantAlocar - dQuantProduto
            dQuantAlocar = dQuantAlocar / dFator
            
            If dQuantEmb <> 0 Then
                            
                lErro = Calcula_Item_GridEmb1(iItem, iIndice1, objProdutoEmbalagem, dQuantProduto, dQuantEmb, sProdMasc)
                If lErro <> SUCESSO Then gError 81716
            
            End If
            
            Do While dQuantAlocar > QTDE_ESTOQUE_DELTA
                
                'Le a capacidade maxima onde capacidade <= dQuantAlocar
                lErro = CF("ProdutoEmbalagem_Le_CapacMaxima", objProdutoEmbalagem, dQuantAlocar, sUMItemPV, objProduto, dFator)
                If lErro <> SUCESSO And lErro <> 100008 Then gError 96119
                
                If lErro <> 100008 Then
                    
                    dQuantProduto = dQuantAlocar
                    dQuantAlocar = dQuantAlocar * dFator
                    
                    'Resultado com casa decimais
                    dQuantEmb = dQuantAlocar / objProdutoEmbalagem.dCapacidade
                    
                    'Arrendonda para cima o numero de embalagens
                    If dQuantEmb - Fix(dQuantEmb) <> 0 Then dQuantEmb = Fix(dQuantEmb) + 1
                    
                    dQuantAlocar = dQuantAlocar / dFator
                    dQuantAlocar = dQuantAlocar - dQuantProduto 'aqui tem que ficar zero
                
                    lErro = Calcula_Item_GridEmb1(iItem, iIndice1, objProdutoEmbalagem, dQuantProduto, dQuantEmb, sProdMasc)
                    If lErro <> SUCESSO Then gError 81717
                    
                Else
                
                    'Le a capacidade minima onde capacidade > dQuantAlocar
                    lErro = CF("ProdutoEmbalagem_Le_CapacMinima", objProdutoEmbalagem, dQuantAlocar, sUMItemPV, objProduto, dFator)
                    If lErro <> SUCESSO And lErro <> 100009 Then gError 96120
                    
                    dQuantProduto = dQuantAlocar * dFator
                    dQuantEmb = 1
                    dQuantAlocar = 0
                
                    lErro = Calcula_Item_GridEmb1(iItem, iIndice1, objProdutoEmbalagem, dQuantProduto, dQuantEmb, sProdMasc)
                    If lErro <> SUCESSO Then gError 81718
                
                End If
                
            Loop
            
            '#######################################
            'Inserido por Wagner
            bAchou = True
            
            'Se não é calculado automático
            '#######################################
        Else
              
            '#######################################
            'Inserido por Wagner
            'Verifica se o item está no grid de embalagens
            bAchou = False
            For iIndice = 1 To objGridEmb.iLinhasExistentes
                If (StrParaInt(GridEmb.TextMatrix(iIndice, iGrid_ItemEmb_Col)) = iItem) Then
                    bAchou = True
                    Exit For
                End If
            Next
            '#######################################
        
        End If
                
        '###################################
        'Alterado por Wagner
        'Se está no grid de embalagens
        If bAchou Then
            'Atualiza com base no Grid de Embalagens
            Call AtualizaTotais
        Else
            'Incrementa o peso do Item (Com base no Produto)
            lErro = Pesos_Recalcula(dQuantAlocar, iItem)
            If lErro <> SUCESSO Then gError 132010
        End If
        '###################################
        
    End If 'Inserido por Wagner

    Calcula_Item_GridEmb = SUCESSO

    Exit Function

Erro_Calcula_Item_GridEmb:

    Calcula_Item_GridEmb = gErr

    Select Case gErr
        'Janaina - Tratar mensagens de erro
        Case 96115, 96116, 96117, 96118, 96119, 96120, 81716, 81717, 81718
        
        Case 132010, 132014 'Inserido por Wagner

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151592)

    End Select

    Exit Function
'Janaina
End Function

Public Function Rotina_Grid_Enable_Emb(iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Rotina_Grid_Enable_Emb

    Select Case objControl.Name
    
        Case Embalagem.Name
            
            If Len(Trim(objGridEmb.objGrid.TextMatrix(objGridEmb.objGrid.Row, iGrid_ProdutoEmb_Col))) = 0 Then
                Embalagem.Enabled = False
            Else
                Embalagem.Enabled = True
            End If
                
        Case QuantEmb.Name
            
            If Len(Trim(objGridEmb.objGrid.TextMatrix(objGridEmb.objGrid.Row, iGrid_Embalagem_Col))) > 0 Then
                QuantEmb.Enabled = True
            Else
                QuantEmb.Enabled = False
            End If
        
        Case QuantProduto.Name
            
            If Len(Trim(objGridEmb.objGrid.TextMatrix(objGridEmb.objGrid.Row, iGrid_Embalagem_Col))) Then
                QuantProduto.Enabled = True
            Else
                QuantProduto.Enabled = False
            End If
            
    End Select
    
    Rotina_Grid_Enable_Emb = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enable_Emb:

    Rotina_Grid_Enable_Emb = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151593)

    End Select

    Exit Function
'Janaina
End Function
'Janaina
Public Sub ItemEmb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEmb.objControle = ItemEmb
    lErro = Grid_Campo_Libera_Foco(objGridEmb)
    If lErro <> SUCESSO Then Cancel = True
'Janaina
End Sub
'Janaina
Public Sub ItemEmb_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub
'Janaina
Public Sub ItemEmb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEmb)

End Sub
'Janaina
Public Sub ItemEmb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEmb)

End Sub
'Janaina
Private Sub Remove_Embalagem(iItemAtual As Integer)
'Remove do grid de embalagem os itens passado

Dim iIndice As Integer
Dim iItemLinha As Integer

    iIndice = 1
    
    Do While iIndice <= objGridEmb.iLinhasExistentes
        If Len(Trim(objGridEmb.objGrid.TextMatrix(iIndice, iGrid_ItemEmb_Col))) > 0 Then
            'Recolhe o item a qual pertence a embalagem
            iItemLinha = objGridEmb.objGrid.TextMatrix(iIndice, iGrid_ItemEmb_Col)
            'Se for o item passado em iItemAtual
            If iItemLinha = iItemAtual Then
                'Exclui do grid de Embalagem as embalagem daquele ItemEmb
                Do While CInt(objGridEmb.objGrid.TextMatrix(iIndice, iGrid_ItemEmb_Col)) = iItemAtual
                    Call Grid_Exclui_Linha(objGridEmb, iIndice)
                    If Len(Trim(objGridEmb.objGrid.TextMatrix(iIndice, iGrid_ItemEmb_Col))) = 0 Then Exit Do
                Loop
            End If
        End If
        iIndice = iIndice + 1
    Loop
'Janaina
End Sub
'Janaina
Public Function Exclusao_Item_GridEmb(iItemAtual As Integer) As Long

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Exclusao_Item_GridEmb

    For iLinha = objGridEmb.iLinhasExistentes To 1 Step -1
        If Len(Trim(objGridEmb.objGrid.TextMatrix(iLinha, iGrid_ItemEmb_Col))) > 0 Then
            If CInt(objGridEmb.objGrid.TextMatrix(iLinha, iGrid_ItemEmb_Col)) = iItemAtual Then
                'exclui a linha em questão
                Call Grid_Exclui_Linha(objGridEmb, iLinha)
            End If
        End If
    Next
    
    For iLinha = 1 To objGridEmb.iLinhasExistentes
        If Len(Trim(objGridEmb.objGrid.TextMatrix(iLinha, iGrid_ItemEmb_Col))) > 0 Then If CInt(objGridEmb.objGrid.TextMatrix(iLinha, iGrid_ItemEmb_Col)) > iItemAtual Then objGridEmb.objGrid.TextMatrix(iLinha, iGrid_ItemEmb_Col) = CInt(objGridEmb.objGrid.TextMatrix(iLinha, iGrid_ItemEmb_Col)) - 1
    Next
    
    '##########################################
    'Alterado por Wagner
    'Recalcula o peso tanto dos produtos com embalagem como os sem
    Call objTela.Pesos_Recalcula
    
    'Janaina - Testar
    'Limpa os totais do Grid de Embalagens
'    QuantEmbTotal.Caption = ""
'    PesoLiqTotal.Caption = ""
'    PesoBrutoTotal.Caption = ""
    'Janaina
    '##########################################
    
    Exclusao_Item_GridEmb = SUCESSO
    
    Exit Function
    
Erro_Exclusao_Item_GridEmb:

    Exclusao_Item_GridEmb = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151594)

    End Select
    
    Exit Function
'Janaina
End Function
''''Janaina
'''Public Function Calcula_SaidaCelula_Embalagem() As Long
'''
'''Dim lErro As Long
'''Dim objProdutoEmbalagem As ClassProdutoEmbalagem
'''
'''    'Se o calculo for automatico
'''    If CalculoAuto.Value = vbChecked Then
'''
'''
'''    End If
'''
'''    Exit Function
''''Janaina
'''End Function

'Janaina
Public Function Valida_Grid_Embalagens() As Long

Dim iIndice As Integer, iIndice1 As Integer
Dim dQuantidade As Double, dQuantPedida As Double
Dim iProdutoPreenchido As Integer
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim lErro As Long, sUMEmb As String, sUMItemPV As String
'##############################################
'INSERIDO POR WAGNER
Dim iLinha As Integer
'##############################################

On Error GoTo Erro_Valida_Grid_Embalagens
   
    'Para cada item do grid de itens
    For iIndice = 1 To objTela.objGridItens.iLinhasExistentes
                    
        '??? converter UM se necessario
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 96135

        objProduto.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 96136
                    
        'Critica se é produto de ESTOQUE (c/ reserva ou sem)
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
        
            If giTipoDocOriginal = DOCORIGEM_NF Then
                dQuantPedida = StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_Quantidade_Col))
            Else
                dQuantPedida = Abs(StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_QuantPedida_Col)) - StrParaDbl(objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_QuantCancel_Col)))
            End If
            sUMItemPV = objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_UnidadeMed_Col)

'        If Len(Trim(objTela.objGridItens.objGrid.TextMatrix(iIndice, objTela.iGrid_QuantPedida_Col))) > 0 Then
            
            dQuantidade = 0
            
            'para cada linha do grid de embalagens
            For iIndice1 = 1 To objGridEmb.iLinhasExistentes
                
                If objTela.objGridItens.objGrid.TextMatrix(iIndice, 0) = GridEmb.TextMatrix(iIndice1, iGrid_ItemEmb_Col) Then
                    
                    
                    '##############################################################
                    'INSERIDO POR WAGNER
                    iLinha = iIndice1
                    
                    If Len(Trim(GridEmb.TextMatrix(iIndice, iGrid_Embalagem_Col))) = 0 Then gError 129962
                    If Len(Trim(GridEmb.TextMatrix(iIndice, iGrid_UMEmb_Col))) = 0 Then gError 129963
                    If StrParaDbl(GridEmb.TextMatrix(iIndice, iGrid_QuantEmb_Col)) = 0 Then gError 129964
                    '##############################################################
                    
                    
                    sUMEmb = GridEmb.TextMatrix(iIndice1, iGrid_UMEmb_Col)
                    
                    'Converte da UM do GridItens para UM GridEmb
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMEmb, sUMItemPV, dFator)
                    If lErro <> SUCESSO Then gError 96117
                                             
                    dQuantidade = dQuantidade + (dFator * GridEmb.TextMatrix(iIndice1, iGrid_QuantProduto_Col))
                    
                End If
                
            Next
            
            If Abs(dQuantPedida - dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 96133
        
        End If

    Next

    Valida_Grid_Embalagens = SUCESSO

    Exit Function

Erro_Valida_Grid_Embalagens:

    Valida_Grid_Embalagens = gErr

    Select Case gErr
    
        Case 96133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTITENS_DIFERENTE_QUANTEMB", gErr)
        
        Case 96135, 96136
        
        '##############################################
        'INSERIDO POR WAGNER
        Case 129962
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 129963
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTEMB_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 129964
            Call Rotina_Erro(vbOKOnly, "ERRO_UMEMBALAGEM_NAO_PREENCHIDA", gErr, iLinha)
        '##############################################
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151595)

    End Select

    Exit Function
'Janaina
End Function
'Janaina
Public Function Preenche_Pesos() As Long
'Preenche o peso liquido e o peso bruto no tab de complemento,
'caso as embalagens estejam preenchidas p/todos os produto inventariados

Dim iIndice As Integer, iIndice1 As Integer
Dim bAchou As Boolean
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_Pesos

'###################################################
'ALTERADO POR WAGNER
'    bAchou = True
'
'    'Para cada item do grid de itens
'    For iIndice1 = 1 To objTela.objGridItens.iLinhasExistentes
'
'        'Formata o Produto para o BD
'        lErro = CF("Produto_Formata", objTela.objGridItens.objGrid.TextMatrix(iIndice1, objTela.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 96141
'
'        objProduto.sCodigo = sProdutoFormatado
'
'        lErro = CF("Produto_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 96142
'
'        'Critica se é produto de ESTOQUE (c/ reserva ou sem)
'        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'            iIndice = 1
'
'            Do While bAchou = True And iIndice <= objGridEmb.iLinhasExistentes
'
'                If objTela.objGridItens.objGrid.TextMatrix(iIndice1, objTela.iGrid_ItemProduto_Col) = GridEmb.TextMatrix(iIndice, iGrid_ItemEmb_Col) Then
'
'                    If Len(Trim(GridEmb.TextMatrix(iIndice, iGrid_Embalagem_Col))) > 0 Then
'                        bAchou = True
'                    Else
'                        bAchou = False
'                    End If
'
'                End If
'
'                iIndice = iIndice + 1
'
'            Loop
'
'        End If
'
'        If bAchou = False Then Exit For
'
'    Next
    
'     If bAchou = True Then

    'Considera o peso = a Peso Produtos com Embalagens + Peso Produtos Sem Embalagens
    objTela.PesoBruto.Text = CStr(gdValorPesoBrutoSemEmb + StrParaDbl(PesoBrutoTotal.Caption)) 'Alterado por Wagner
    objTela.PesoLiquido.Text = CStr(gdValorPesoLiqSemEmb + StrParaDbl(PesoLiqTotal.Caption)) 'Alterado por Wagner
    objTela.VolumeQuant.Text = CStr(StrParaLong(QuantEmbTotal.Caption))
'     End If
'###################################################
        
   Preenche_Pesos = SUCESSO

    Exit Function

Erro_Preenche_Pesos:

    Preenche_Pesos = gErr

    Select Case gErr
       
        Case 96141, 96142
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151596)

    End Select

    Exit Function
'Janaina
End Function

Private Function Calcula_Item_GridEmb1(ByVal iItem As Integer, iLinha As Integer, objProdutoEmbalagem As ClassProdutoEmbalagem, ByVal dQuantProduto As Double, ByVal dQuantEmb As Double, sProdMasc As String) As Long
'Calculo automatico das embalagens

Dim lErro As Long
Dim dFator As Double
Dim dPesoLiq As Double
Dim dPesoBruto As Double, iLinhaExistente As Integer
Dim dPesoEmbalagem As Double, iIndice As Integer
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_Calcula_Item_GridEmb1

    'verifica se o par produto-embalagem já está no grid
    
    'Le a sigla da embalagem
    objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
    
    lErro = CF("Embalagem_Le", objEmbalagem)
    If lErro <> SUCESSO Then gError 96118
            
    iLinhaExistente = 0
    
    For iIndice = 1 To objGridEmb.iLinhasExistentes
    
        If StrParaInt(GridEmb.TextMatrix(iIndice, iGrid_ItemEmb_Col)) = iItem And _
           GridEmb.TextMatrix(iIndice, iGrid_Embalagem_Col) = objEmbalagem.sSigla Then
           
            iLinhaExistente = iIndice
            Exit For
            
        End If
    
    Next
    
    'se nao está no grid
    If iLinhaExistente = 0 Then
    
        iLinha = iLinha + 1
        objGridEmb.iLinhasExistentes = iLinha
        
        lErro = CalculaPesos(dQuantEmb, objProdutoEmbalagem.dPesoBruto - objProdutoEmbalagem.dPesoLiqTotal, objProdutoEmbalagem.sUMPeso, dQuantProduto, objProdutoEmbalagem.dCapacidade, objProdutoEmbalagem.dPesoLiqTotal, dPesoLiq, dPesoBruto)
        If lErro <> SUCESSO Then gError 96117
        
        GridEmb.TextMatrix(iLinha, iGrid_ItemEmb_Col) = iItem
        GridEmb.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = sProdMasc
        GridEmb.TextMatrix(iLinha, iGrid_Embalagem_Col) = objEmbalagem.sSigla
        GridEmb.TextMatrix(iLinha, iGrid_UMEmb_Col) = objProdutoEmbalagem.sUMEmbalagem
        GridEmb.TextMatrix(iLinha, iGrid_Capacidade_Col) = objProdutoEmbalagem.dCapacidade
        GridEmb.TextMatrix(iLinha, iGrid_QuantEmb_Col) = dQuantEmb
        GridEmb.TextMatrix(iLinha, iGrid_QuantProduto_Col) = Round(dQuantProduto, 4)
        GridEmb.TextMatrix(iLinha, iGrid_PesoLiq_Col) = Format(dPesoLiq, FORMATO_PESO) 'Alterado por Wagner
        GridEmb.TextMatrix(iLinha, iGrid_PesoBruto_Col) = Format(dPesoBruto, FORMATO_PESO) 'Alterado por Wagner
    
    Else
    
        'é preciso incorporar qtdes e pesos a linhaexistente
    
        lErro = CalculaPesos(dQuantEmb, objProdutoEmbalagem.dPesoBruto - objProdutoEmbalagem.dPesoLiqTotal, objProdutoEmbalagem.sUMPeso, dQuantProduto, objProdutoEmbalagem.dCapacidade, objProdutoEmbalagem.dPesoLiqTotal, dPesoLiq, dPesoBruto)
        If lErro <> SUCESSO Then gError 96117
        
        GridEmb.TextMatrix(iLinhaExistente, iGrid_QuantEmb_Col) = dQuantEmb + StrParaDbl(GridEmb.TextMatrix(iLinhaExistente, iGrid_QuantEmb_Col))
        GridEmb.TextMatrix(iLinhaExistente, iGrid_QuantProduto_Col) = Round(StrParaDbl(GridEmb.TextMatrix(iLinhaExistente, iGrid_QuantProduto_Col)) + dQuantProduto, 4)
        GridEmb.TextMatrix(iLinha, iGrid_PesoLiq_Col) = Format(StrParaDbl(GridEmb.TextMatrix(iLinha, iGrid_PesoLiq_Col) + dPesoLiq), FORMATO_PESO) 'Alterado por Wagner
        GridEmb.TextMatrix(iLinha, iGrid_PesoBruto_Col) = Format(StrParaDbl(GridEmb.TextMatrix(iLinha, iGrid_PesoBruto_Col)) + dPesoBruto, FORMATO_PESO) 'Alterado por Wagner
    
    End If
    
    Calcula_Item_GridEmb1 = SUCESSO

    Exit Function

Erro_Calcula_Item_GridEmb1:

    Calcula_Item_GridEmb1 = gErr

    Select Case gErr

        Case 96117, 96118
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151597)

    End Select

    Exit Function

End Function

Public Sub BotaoEmbalagens_Click()

Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim sSQL As String

On Error GoTo Erro_BotaoEmbalagens_Click

    'Verifica se o produto está preenchido
    If Len(Trim(GridEmb.TextMatrix(GridEmb.Row, iGrid_ProdutoEmb_Col))) < 0 Then gError 96144

    'Verifica se há uma linha selecionada
    If GridEmb.Row = 0 Then gError 96145
        
    lErro = CF("Produto_Formata", GridEmb.TextMatrix(GridEmb.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 96160

    sSQL = "Produto = '" & sProdutoFormatado & "'"
    
   'chama a tela de browser
    Call Chama_Tela("ProdutoEmbalagemLista", colSelecao, objProdutoEmbalagem, objEventoEmbalagens, sSQL)
    
    Exit Sub
    
Erro_BotaoEmbalagens_Click:
    
    Select Case gErr
    
        Case 96145
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 96144
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 96160
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151598)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEmbalagens_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim iLinha As Integer
Dim iCodigoAux As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer, dQuantProduto As Double, dQuantEmb As Double

On Error GoTo Erro_objEventoEmbalagens_evSelecao
           
    'Define o tipo de obj recebido (Tipo ProdutoEmbalagem)
    Set objProdutoEmbalagem = obj1

    'Verifica se há alguma relacao Produto X Embalagem repetida no grid
    For iLinha = 1 To objGridEmb.iLinhasExistentes
        
        If iLinha <> GridEmb.Row Then
                            
            'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
            faz uma nova leitura em busca do codigo
            objEmbalagem.sSigla = GridEmb.TextMatrix(iLinha, iGrid_Embalagem_Col)
            lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
            If lErro <> SUCESSO And lErro <> 95088 Then gError 95468
            
            iCodigoAux = objEmbalagem.iCodigo
            
            lErro = CF("Produto_Formata", GridEmb.TextMatrix(iLinha, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 96152

            If iCodigoAux = objProdutoEmbalagem.iEmbalagem And sProdutoFormatado = objProdutoEmbalagem.sProduto Then gError 96146
                
        End If
                       
    Next
    
    objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
     
    'Lê a embalagem à partir do Codigo
    lErro = CF("Embalagem_Le", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 25060 Then gError 96155

    lErro = CF("Produto_Formata", GridEmb.TextMatrix(GridEmb.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 96152

    objProdutoEmbalagem.sProduto = sProdutoFormatado
            
    lErro = CF("ProdutoEmbalagem_Le", objProdutoEmbalagem)
    If lErro <> SUCESSO And lErro <> 96156 Then gError 96151
        
    'Preenche o grid de embalagens
    GridEmb.TextMatrix(GridEmb.Row, iGrid_Embalagem_Col) = objEmbalagem.sSigla
    GridEmb.TextMatrix(GridEmb.Row, iGrid_UMEmb_Col) = objProdutoEmbalagem.sUMEmbalagem
    GridEmb.TextMatrix(GridEmb.Row, iGrid_Capacidade_Col) = Formata_Estoque(objProdutoEmbalagem.dCapacidade)
    
    Embalagem.Text = objEmbalagem.sSigla

    'Cria mais uma linha no grid
    If GridEmb.Row - GridEmb.FixedRows = objGridEmb.iLinhasExistentes Then objGridEmb.iLinhasExistentes = objGridEmb.iLinhasExistentes + 1
    
    'se já há qtde de produto recalcula a qtde de embalagens
    dQuantProduto = StrParaDbl(GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantProduto_Col))
    If dQuantProduto <> 0 Then
    
        dQuantEmb = Arredonda_ParaCima(dQuantProduto / objProdutoEmbalagem.dCapacidade)
    
    Else
    
        dQuantEmb = 0
    
    End If
    
    GridEmb.TextMatrix(GridEmb.Row, iGrid_QuantEmb_Col) = Formata_Estoque(dQuantEmb)
    
    lErro = CalculaPesosLinha(GridEmb.Row, dQuantEmb, dQuantProduto)
    If lErro <> SUCESSO Then gError 81727
    
    objTela.Show
    
    Exit Sub
    
Erro_objEventoEmbalagens_evSelecao:

    Select Case gErr

        Case 96146
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_REPETIDA", gErr, objEmbalagem.iCodigo, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridEmb)
        
        Case 81727, 96151, 96152, 96155
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151599)
              
    End Select
    
    Exit Sub

End Sub

Public Sub Embalagem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEmb.objControle = Embalagem
    lErro = Grid_Campo_Libera_Foco(objGridEmb)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Embalagem_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Embalagem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEmb)

End Sub

Public Sub Embalagem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEmb)

End Sub


Public Sub QuantEmb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEmb.objControle = QuantEmb
    lErro = Grid_Campo_Libera_Foco(objGridEmb)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantEmb_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantEmb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEmb)

End Sub

Public Sub QuantEmb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEmb)

End Sub

Public Sub QuantProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEmb.objControle = QuantProduto
    lErro = Grid_Campo_Libera_Foco(objGridEmb)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantProduto_Change()

    objTela.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEmb)

End Sub

Public Sub QuantProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEmb)

End Sub

Private Function CalculaPesos(dQuantEmb As Double, dPesoEmb As Double, sUMPesoEmb As String, dQuantProduto As Double, dCapacidade As Double, dPesoLiqTotalEmb As Double, dPesoLiq As Double, dPesoBruto As Double) As Long
'Calcula o peso líquido e o bruto em funcao dos outros parametros
'supoe que dQuantProduto e dCapacidade estejam na mesma unidade

Dim lErro As Long, dFator As Double

On Error GoTo Erro_CalculaPesos

    'obter fator de conversao da um de peso da embalagem p/kg
    lErro = CF("UM_Conversao_Trans", UM_PESO_CLASSE, sUMPesoEmb, "kg", dFator)
    If lErro <> SUCESSO Then gError 81723

    dPesoLiq = Round(dPesoLiqTotalEmb * dQuantProduto * dFator / dCapacidade, NUM_CASAS_DECIMAIS_PESO) 'Alterado por Wagner
    dPesoBruto = Round((dPesoLiq + (dPesoEmb * dFator * dQuantEmb)), NUM_CASAS_DECIMAIS_PESO) 'Alterado por Wagner
    
    CalculaPesos = SUCESSO
     
    Exit Function
    
Erro_CalculaPesos:

    CalculaPesos = gErr
     
    Select Case gErr
          
        Case 81723
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151600)
     
    End Select
     
    Exit Function

End Function

Function CalculaPesosLinha(iLinha As Integer, dQuantEmb As Double, dQuantProduto As Double) As Long
'Preenche os campos peso bruto e liquido

Dim lErro As Long, dPesoEmb As Double, sUMPesoEmb As String, dCapacidade As Double, dPesoLiqTotalEmb As Double, dPesoLiq As Double, dPesoBruto As Double
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_CalculaPesosLinha

    objEmbalagem.sSigla = GridEmb.TextMatrix(iLinha, iGrid_Embalagem_Col)
    lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 95088 Then gError 95460
    
    'Se nao achou => ERRO
    If lErro = 95088 Then gError 95461
    
    objProdutoEmbalagem.iEmbalagem = objEmbalagem.iCodigo
    
    lErro = CF("Produto_Formata", GridEmb.TextMatrix(iLinha, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 81724
    
    objProdutoEmbalagem.sProduto = sProdutoFormatado
    
    lErro = CF("ProdutoEmbalagem_Le", objProdutoEmbalagem)
    If lErro <> SUCESSO Then gError 81725
    
    lErro = CalculaPesos(dQuantEmb, objProdutoEmbalagem.dPesoEmbalagem, objProdutoEmbalagem.sUMPeso, dQuantProduto, objProdutoEmbalagem.dCapacidade, objProdutoEmbalagem.dPesoLiqTotal, dPesoLiq, dPesoBruto)
    If lErro <> SUCESSO Then gError 81726
    
    GridEmb.TextMatrix(iLinha, iGrid_PesoLiq_Col) = Format(dPesoLiq, FORMATO_PESO)
    GridEmb.TextMatrix(iLinha, iGrid_PesoBruto_Col) = Format(dPesoBruto, FORMATO_PESO)
    
    '##########################
    'Inserido por Wagner
    'Para poder considerar produtos sem embalagens no grid
    Call objTela.Pesos_Recalcula

'    Call AtualizaTotais
'
'    Call Preenche_Pesos
    '##########################
    
    CalculaPesosLinha = SUCESSO
     
    Exit Function
    
Erro_CalculaPesosLinha:

    CalculaPesosLinha = gErr
     
    Select Case gErr
                    
        Case 81724, 81725, 81726, 95461
                
        Case 95460
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.sSigla)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151601)
     
    End Select
     
    Exit Function

End Function

Sub AtualizaTotais()

    Call Soma_Valor_PesoBruto(objGridEmb)
    Call Soma_Valor_PesoLiq(objGridEmb)
    Call Soma_Valor_QuantEmb(objGridEmb)

End Sub

'##############################################
'Inserido por Wagner
Function Pesos_Recalcula(ByVal dQuantidade As Double, ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim dFator As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dPesoBruto As Double
Dim dPesoLiquido As Double

On Error GoTo Erro_Pesos_Recalcula
        
    'Se o produto e a UM estiverem preenchidos
    If Len(Trim(objTela.GridItens.TextMatrix(iIndice, objTela.iGrid_Produto_Col))) > 0 And Len(Trim(objTela.GridItens.TextMatrix(iIndice, objTela.iGrid_UnidadeMed_Col))) > 0 Then
        
        'Formata o produto
        lErro = CF("Produto_Formata", objTela.GridItens.TextMatrix(iIndice, objTela.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132011
        
        'Lê o produto
        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 132012

        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 132013
                    
        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTela.GridItens.TextMatrix(iIndice, objTela.iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 132014
        
        If objProduto.dPesoBruto < objProduto.dPesoLiq Then objProduto.dPesoBruto = objProduto.dPesoLiq
                                       
        'Acumula o peso Bruto e o peso líquido
        dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * dQuantidade
        dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * dQuantidade
                    
    End If
        
    'Coloca os Pesos na variável global (ÚNICO LUGAR ONDE ESSAS VAR SÃO ALTERADAS)
    gdValorPesoBrutoSemEmb = gdValorPesoBrutoSemEmb + dPesoBruto
    gdValorPesoLiqSemEmb = gdValorPesoLiqSemEmb + dPesoLiquido
    
    Pesos_Recalcula = SUCESSO
    
    Exit Function
    
Erro_Pesos_Recalcula:

    Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case 132011, 132012, 132014
        
        Case 132013
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151602)
    
    End Select
    
    Exit Function
    
End Function
'#######################################################################
