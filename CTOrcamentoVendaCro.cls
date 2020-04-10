VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcVendaCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela
Public Property Get CondPagtoItem(ByVal objCT As Object) As Object
     Set CondPagtoItem = objCT.objUserControl.Controls("CondPagtoItem")
End Property

Public Property Get Concorrente(ByVal objCT As Object) As Object
     Set Concorrente = objCT.objUserControl.Controls("Concorrente")
End Property

Public Property Get Escolhido(ByVal objCT As Object) As Object
     Set Escolhido = objCT.objUserControl.Controls("Escolhido")
End Property
'--- fim dos properties get dos controles da tela

Public Sub Escolhido_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Escolhido_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub Escolhido_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub Escolhido_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = Escolhido(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Concorrente_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Concorrente_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub Concorrente_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub Concorrente_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = Concorrente(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub CondPagtoItem_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CondPagtoItem_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub CondPagtoItem_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub CondPagtoItem_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = CondPagtoItem(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function OV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.gobjInfoUsu.iGrid_Concorrente_Col
        
            lErro = Saida_Celula_Concorrente(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 141392

        Case objCT.gobjInfoUsu.iGrid_CondPagto_Col
        
            lErro = Saida_Celula_CondPagto(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 141393

        Case objCT.gobjInfoUsu.iGrid_Escolhido_Col
        
            lErro = Saida_Celula_Escolhido(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 141394

    End Select

    OV_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_OV_Saida_Celula_GridItens:

    OV_Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 141392 To 141394

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Function Saida_Celula_Concorrente(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Concorrente

    Set objGridInt.objControle = Concorrente(objCT)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141395
    
    Saida_Celula_Concorrente = SUCESSO
     
    Exit Function
    
Erro_Saida_Celula_Concorrente:

    Saida_Celula_Concorrente = gErr
     
    Select Case gErr
          
        Case 141395
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
     
    End Select
     
    Exit Function

End Function

Function Saida_Celula_Escolhido(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Escolhido

    Set objGridInt.objControle = Escolhido(objCT)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141396
    
    Saida_Celula_Escolhido = SUCESSO
     
    Exit Function
    
Erro_Saida_Celula_Escolhido:

    Saida_Celula_Escolhido = gErr
     
    Select Case gErr
          
        Case 141396
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
     
    End Select
     
    Exit Function

End Function

Private Function Saida_Celula_CondPagto(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_CondPagto

    Set objGridInt.objControle = CondPagtoItem(objCT)

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondPagtoItem(objCT).Text)) <> 0 Then

        'Verifica se é uma Condicaopagamento selecionada
        If CondPagtoItem(objCT).Text <> CondPagtoItem(objCT).List(CondPagtoItem(objCT).ListIndex) Then
    
            'Tenta selecionar na combo
            lErro = Combo_Seleciona_Grid(CondPagtoItem(objCT), iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 141397
            
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                objCondicaoPagto.iCodigo = iCodigo
        
                'Tenta ler CondicaoPagto com esse código no BD
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 141398
                
                'Não encontrou CondicaoPagto no BD
                If lErro <> SUCESSO Then gError 141399
        
                'Encontrou CondicaoPagto no BD e não é de Recebimento
                If objCondicaoPagto.iEmRecebimento = 0 Then gError 141400
        
                'Coloca no Text da Combo
                CondPagtoItem(objCT).Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
        
            End If
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then gError 141401
            
            objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_CondPagto_Col) = CondPagtoItem(objCT).Text
        
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141402

    Saida_Celula_CondPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_CondPagto:

    Saida_Celula_CondPagto = gErr

    Select Case gErr
    
        Case 141397, 141398, 141402
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 141399

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

            End If
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 141400
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 141401
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondPagtoItem(objCT).Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function OV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objItemOVCro As New ClassItemOVCro

On Error GoTo Erro_OV_Move_GridItens_Memoria

    objItemOVCro.iEscolhido = StrParaInt(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Escolhido_Col))
    objItemOVCro.iCondPagto = Codigo_Extrai(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_CondPagto_Col))
    objItemOVCro.sConcorrente = objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Concorrente_Col)

    Set objItemOV.objInfoUsu = objItemOVCro

    OV_Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_OV_Move_GridItens_Memoria:

    OV_Move_GridItens_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function OV_Traz_GridItens_Tela(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objItemOVCro As ClassItemOVCro

On Error GoTo Erro_OV_Traz_GridItens_Tela

    Set objItemOVCro = objItemOV.objInfoUsu

    If objItemOVCro.iCondPagto <> 0 Then
        CondPagtoItem(objCT).Text = objItemOVCro.iCondPagto
        lErro = Combo_Seleciona_Grid(CondPagtoItem(objCT), objItemOVCro.iCondPagto)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129564
        
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_CondPagto_Col) = CondPagtoItem(objCT).Text
    End If
            
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Escolhido_Col) = objItemOVCro.iEscolhido
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Concorrente_Col) = objItemOVCro.sConcorrente

    Call Grid_Refresh_Checkbox(objCT.objGridItens)

    OV_Traz_GridItens_Tela = SUCESSO

    Exit Function

Erro_OV_Traz_GridItens_Tela:

    OV_Traz_GridItens_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function OV_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objItemOVCro As ClassItemOVCro

On Error GoTo Erro_OV_Form_Load

    Call Carrega_CondicaoPagamento(objCT)

    OV_Form_Load = SUCESSO

    Exit Function

Erro_OV_Form_Load:

    OV_Form_Load = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento(ByVal objCT As Object) As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 141402

   For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        CondPagtoItem(objCT).AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        CondPagtoItem(objCT).ItemData(CondPagtoItem(objCT).NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 141402

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function OV_Inicializa_Grid_Itens(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens
Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = objCT
    Set objUserControl = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Escolhido")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    
    'Inserido por Wagner
    '########################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    '########################################
    
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens4", objGridInt)
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens1", objGridInt)
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Data Entrega")
    
    '###############################
    'Inserido por Wagner 09/12/05
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Motivo Perda")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Concorrente")
    '###############################

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Escolhido(objCT).Name)
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    
    'Inserido por Wagner
    '############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (objCT.VersaoKit.Name)
        objGridInt.colCampo.Add (objCT.VersaoKitBase.Name)
        
    Else
    
        objCT.VersaoKit.Left = POSICAO_FORA_TELA
        objCT.VersaoKit.TabStop = False
    
        objCT.VersaoKitBase.Left = POSICAO_FORA_TELA
        objCT.VersaoKitBase.TabStop = False

    End If
    '############################################

    Call CF("Orcamento_Inicializa_Grid_Itens5", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens2", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
    
    '###############################
    'Inserido por Wagner 09/12/05
    objGridInt.colCampo.Add (objCT.StatusItem.Name)
    objGridInt.colCampo.Add (objCT.MotivoPerdaItem.Name)
    objGridInt.colCampo.Add (CondPagtoItem(objCT).Name)
    objGridInt.colCampo.Add (objCT.Observacao.Name)
    objGridInt.colCampo.Add (Concorrente(objCT).Name)
    '###############################
    
    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    
    'cromaton
    objCT.gobjInfoUsu.iGrid_Escolhido_Col = 1
    iIncremento = 1
    
    objCT.iGrid_Produto_Col = 2
    objCT.iGrid_DescProduto_Col = 3
    
    'Inserido por Wagner
    '#############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objCT.iGrid_VersaoKit_Col = 4
        objCT.iGrid_VersaoKitBase_Col = 5
        iIncremento = iIncremento + 2
        
    End If
    '#############################################
    
    Call CF("Orcamento_Inicializa_Grid_Itens6", iIncremento)
    objCT.iGrid_UnidadeMed_Col = 3 + iIncremento
    objCT.iGrid_Quantidade_Col = 4 + iIncremento
    objCT.iGrid_PrecoUnitario_Col = 5 + iIncremento
    Call CF("Orcamento_Inicializa_Grid_Itens3", iIncremento)
    objCT.iGrid_PercDesc_Col = 6 + iIncremento
    objCT.iGrid_Desconto_Col = 7 + iIncremento
    objCT.iGrid_PrecoTotal_Col = 8 + iIncremento
    objCT.iGrid_DataEntrega_Col = 9 + iIncremento
    
    '###############################
    'Inserido por Wagner 09/12/05
    objCT.iGrid_Status_Col = 10 + iIncremento
    objCT.iGrid_MotivoPerda_Col = 11 + iIncremento
    
    objCT.gobjInfoUsu.iGrid_CondPagto_Col = 12 + iIncremento
    iIncremento = iIncremento + 1
    
    objCT.iGrid_Observacao_Col = 12 + iIncremento
    objCT.gobjInfoUsu.iGrid_Concorrente_Col = 13 + iIncremento
    '###############################
    
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

    OV_Inicializa_Grid_Itens = SUCESSO

    Exit Function
    
End Function

Public Function OV_Produto_Valida(ByVal objCT As Object) As Long
    
    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_CondPagto_Col) = objCT.CondicaoPagamento.Text
    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Escolhido_Col) = MARCADO
    
    Call Grid_Refresh_Checkbox(objCT.objGridItens)
    
    OV_Produto_Valida = SUCESSO
    
End Function

Public Function OV_Valida_Itens(ByVal objCT As Object, ByVal objOV As ClassOrcamentoVenda) As Long
    
Dim lErro As Long
Dim objItemOV As ClassItemOV
Dim objItemOVCro As ClassItemOVCro
Dim objItemOVAux As ClassItemOV
Dim objItemOVCroAux As ClassItemOVCro
Dim iIndice1 As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_OV_Valida_Itens

    For Each objItemOV In objOV.colItens
    
        Set objItemOVCro = objItemOV.objInfoUsu
    
        iIndice1 = iIndice1 + 1
        
        'Se esse item tiver sido escolhido
        If objItemOVCro.iEscolhido = MARCADO Then
        
            iIndice2 = 0
              
            For Each objItemOVAux In objOV.colItens
             
                Set objItemOVCroAux = objItemOVAux.objInfoUsu
             
                iIndice2 = iIndice2 + 1
                 
                'Se não se trata do mesmo item e o item estiver marcado
                If iIndice1 <> iIndice2 And objItemOVCroAux.iEscolhido = MARCADO Then
                 
                    'Se o Produto estiver repetido
                    If objItemOV.sProduto = objItemOVAux.sProduto Then gError 141410
                    
                    'Se a condição de pagamento for diferente
                    If objItemOVCro.iCondPagto <> objItemOVCroAux.iCondPagto Then gError 141411
                    
                End If
                 
            Next
    
        End If
    
    Next

    OV_Valida_Itens = SUCESSO

    Exit Function

Erro_OV_Valida_Itens:

    OV_Valida_Itens = gErr

    Select Case gErr

        Case 141410
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_ESCOLHIDO_REPETIDO", gErr, iIndice1, iIndice2)
        
        Case 141411
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDPAGTO_ESCOLHIDA_DIFERENTE", gErr, iIndice1, iIndice2)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
    
End Function

Public Function OV_Rotina_Grid_Enable(ByVal objCT As Object, ByVal sNomeControle As String, ByVal iProdutoPreenchido As Integer) As Long
    
Dim lErro As Long

On Error GoTo Erro_OV_Rotina_Grid_Enable

    Select Case sNomeControle
    
        Case Concorrente(objCT).Name
        
            If iProdutoPreenchido = MARCADO Then
                Concorrente(objCT).Enabled = True
            Else
                Concorrente(objCT).Enabled = False
            End If
                
        Case CondPagtoItem(objCT).Name
                
            If iProdutoPreenchido = MARCADO Then
                CondPagtoItem(objCT).Enabled = True
            Else
                CondPagtoItem(objCT).Enabled = False
            End If
            
        Case Escolhido(objCT).Name
            
            If iProdutoPreenchido = MARCADO Then
                Escolhido(objCT).Enabled = True
            Else
                Escolhido(objCT).Enabled = False
            End If
            
    End Select

    OV_Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_OV_Rotina_Grid_Enable:

    OV_Rotina_Grid_Enable = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
    
End Function
