VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcVendaDan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoOS As AdmEvento
Attribute objEventoOS.VB_VarHelpID = -1

Private gobjCT As Object

'--- inicio dos properties get dos controles da tela
Public Property Get Dan_OS(ByVal objCT As Object) As Object
     Set Dan_OS = objCT.objUserControl.Controls("Dan_OS")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Dan_OS_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Dan_OS_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub Dan_OS_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub Dan_OS_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = Dan_OS(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function OV_FormPreco_ValidaDados(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    
Dim lErro As Long
Dim objItemOV As ClassItemOV
Dim objUsuario As New ClassDicUsuario

On Error GoTo Erro_OV_FormPreco_ValidaDados

    If objOrcamentoVenda.iStatusComercial = OV_STATUS_COMERCIAL_REVISADO Then
    
        For Each objItemOV In objOrcamentoVenda.colItens
        
            If objItemOV.dPCPrecoUnitCalc - objItemOV.dPrecoUnitario > DELTA_VALORMONETARIO Then gError 184751
        
        Next
        
    End If
    
    If objOrcamentoVenda.iStatusComercial = OV_STATUS_COMERCIAL_LIBERADO Then
    
        objUsuario.sCodUsuario = gsUsuario
    
        lErro = DicUsuario_Le(objUsuario)
        If lErro <> SUCESSO Then gError 184752
        
        If UCase(objUsuario.sCodGrupo) <> "SUPERVISOR" Then gError 184753
        
    End If
    
    OV_FormPreco_ValidaDados = SUCESSO
    
    Exit Function
    
Erro_OV_FormPreco_ValidaDados:

    OV_FormPreco_ValidaDados = gErr

    Select Case gErr

        Case 184751
            Call Rotina_Erro(vbOKOnly, "ERRO_OV_PRECO_MENOR_CALC", gErr)
        
        Case 184752
        
        Case 184753
            Call Rotina_Erro(vbOKOnly, "ERRO_OV_LIBERACAO_POR_SUPERVISOR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184750)

    End Select
    
    Exit Function
    
End Function

Public Function OV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

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
    
    'Inserido por Wagner
    '########################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    '########################################
    
    'Dan-Presse
    objGridInt.colColuna.Add ("O.S.")
        
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens4", objGridInt)
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens1", objGridInt)
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")
    
    '###############################
    'Inserido por Wagner 09/12/05
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Motivo Perda")
    objGridInt.colColuna.Add ("Observação")
    '###############################

    'Controles que participam do Grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colCampo.Add (objCT.SomaItem.Name)
    End If
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    
    'Inserido por Wagner
    '############################################
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (objCT.VersaoKit.Name)
        objGridInt.colCampo.Add (objCT.VersaoKitBase.Name)
        
    Else
    
        objCT.VersaoKit.left = POSICAO_FORA_TELA
        objCT.VersaoKit.TabStop = False
    
        objCT.VersaoKitBase.left = POSICAO_FORA_TELA
        objCT.VersaoKitBase.TabStop = False

    End If
    '############################################

    'Dan-Presse
    objGridInt.colCampo.Add (Dan_OS(objCT).Name)
    
    Call CF("Orcamento_Inicializa_Grid_Itens5", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens2", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
    
    '###############################
    'Inserido por Wagner 09/12/05
    objGridInt.colCampo.Add (objCT.StatusItem.Name)
    objGridInt.colCampo.Add (objCT.MotivoPerdaItem.Name)
    objGridInt.colCampo.Add (objCT.Observacao.Name)
    '###############################
    
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
    
    'Dan-Presse
    objCT.gobjInfoUsu.iGrid_Dan_OS_Col = 3 + iIncremento
    iIncremento = iIncremento + 1
    
    Call CF("Orcamento_Inicializa_Grid_Itens6", iIncremento)
    objCT.iGrid_UnidadeMed_Col = 3 + iIncremento
    objCT.iGrid_Quantidade_Col = 4 + iIncremento
    objCT.iGrid_PrecoUnitario_Col = 5 + iIncremento
    Call CF("Orcamento_Inicializa_Grid_Itens3", iIncremento)
    objCT.iGrid_PercDesc_Col = 6 + iIncremento
    objCT.iGrid_Desconto_Col = 7 + iIncremento
    objCT.iGrid_PrecoTotal_Col = 8 + iIncremento
    objCT.iGrid_PrecoTotalB_Col = 9 + iIncremento
    objCT.iGrid_DataEntrega_Col = 10 + iIncremento
    
    '###############################
    'Inserido por Wagner 09/12/05
    objCT.iGrid_Status_Col = 11 + iIncremento
    objCT.iGrid_MotivoPerda_Col = 12 + iIncremento
    objCT.iGrid_Observacao_Col = 13 + iIncremento
    '###############################
    
    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.iLinhasVisiveis = 4
    Else
        objGridInt.iLinhasVisiveis = 5
    End If

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

Public Sub Botao_Dan_OS_Click(ByVal objCT As Object)

Dim lErro As Long, iLinha As Integer
Dim objDan_OS As New ClassDan_OS, sProduto1 As String, sProduto As String
Dim colSelecao As New Collection, iPreenchido As Integer, sSelecao As String
Dim objCliente As New ClassCliente

On Error GoTo Erro_Botao_Dan_OS_Click

    If objCT.GridItens.Row = 0 Then gError 141743

    Set objEventoOS = New AdmEvento
    
    Set gobjCT = objCT

    iLinha = objCT.GridItens.Row
    sProduto1 = objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col)
    
    'Verifica se o Produto está preenchido
    If Len(Trim(sProduto1)) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 184856
        
        If iPreenchido <> PRODUTO_VAZIO Then
        
            sSelecao = "EXISTS (SELECT 1 FROM Dan_ItensOS WHERE Produto = ? AND Dan_ItensOS.OS = Dan_OSCli.OS)"
            colSelecao.Add sProduto
            
        End If
        
    End If
    
    'Verifica se o Cliente foi preenchido
    If Len(Trim(objCT.Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = objCT.Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 84239 '26779

        If lErro = SUCESSO Then
            If sSelecao <> "" Then
                sSelecao = sSelecao & " AND "
            End If
            sSelecao = sSelecao & "Dan_OSCli.Cliente = ?"
            colSelecao.Add objCliente.lCodigo
        End If
            
    End If

    Call Chama_Tela("Dan_OSCliLista", colSelecao, objDan_OS, objEventoOS, sSelecao)
        
    Exit Sub

Erro_Botao_Dan_OS_Click:

    Select Case gErr

        Case 141743
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184831)

    End Select

    Exit Sub

End Sub

Public Function OV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long
    
    objItemOV.sVersaoKitBase = objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Dan_OS_Col)
    
    OV_Move_GridItens_Memoria = SUCESSO
    
End Function

Public Function OV_Traz_GridItens_Tela(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long
    
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Dan_OS_Col) = objItemOV.sVersaoKitBase
    
    OV_Traz_GridItens_Tela = SUCESSO
End Function

Public Function OV_Rotina_Grid_Enable(ByVal objCT As Object, ByVal sNomeControle As String, ByVal iProdutoPreenchido As Integer) As Long
    
Dim lErro As Long
    
On Error GoTo Erro_OV_Rotina_Grid_Enable

    Select Case sNomeControle
    
        Case Dan_OS(objCT).Name
        
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                Dan_OS(objCT).Enabled = False
            Else
                Dan_OS(objCT).Enabled = True
            End If
    
    End Select
    
    OV_Rotina_Grid_Enable = SUCESSO
    
    Exit Function
    
Erro_OV_Rotina_Grid_Enable:

    OV_Rotina_Grid_Enable = gErr

    Select Case gErr

        Case 184850
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184849)

    End Select
    
    Exit Function

End Function

Public Function OV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.gobjInfoUsu.iGrid_Dan_OS_Col
        
            lErro = Saida_Celula_Dan_OS(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 141392

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

Private Function Saida_Celula_Dan_OS(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long, dQtde As Double, iLinha As Integer, objDan_ItensOS As New ClassDan_ItensOS
Dim iPreenchido As Integer, sProduto As String, sProduto1 As String, sOS As String, sOSAnterior As String
Dim dQuantAnterior As Double

On Error GoTo Erro_Saida_Celula_Dan_OS

    Set objGridInt.objControle = Dan_OS(objCT)
    
    iLinha = objCT.GridItens.Row
    
    sOSAnterior = objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Dan_OS_Col)
    sOS = Trim(Dan_OS(objCT).Text)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141395
    
    If sOS <> sOSAnterior And Len(sOS) <> 0 Then
    
        lErro = OS_Atualiza_Linha(objCT, iLinha, sOS)
        If lErro <> SUCESSO Then gError 141395
            
    End If
    
    Saida_Celula_Dan_OS = SUCESSO
     
    Exit Function
    
Erro_Saida_Celula_Dan_OS:

    Saida_Celula_Dan_OS = gErr
     
    Select Case gErr
          
        Case 141395, 184856, 184857
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
     
    End Select
     
    Exit Function

End Function

Public Function OV_Form_UnLoad(ByVal objCT As Object, Cancel As Integer) As Long

    Set objEventoOS = Nothing

    OV_Form_UnLoad = SUCESSO
    
End Function
    
Public Function OV_Form_Load(ByVal objCT As Object) As Long
    
    Set objEventoOS = New AdmEvento
    
    OV_Form_Load = SUCESSO

End Function

Public Sub objEventoOS_evSelecao(obj1 As Object)

Dim lErro As Long, iLinha As Integer
Dim objCT As Object
Dim objDan_OS As ClassDan_OS

On Error GoTo Erro_objEventoOS_evSelecao

    Set objDan_OS = obj1
    
    Set objCT = gobjCT

    iLinha = objCT.GridItens.Row

    'Verifica se tem alguma linha selecionada no Grid
    If iLinha <> 0 Then
        
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Dan_OS_Col) = objDan_OS.sOS
    
        lErro = OS_Atualiza_Linha(objCT, iLinha, objDan_OS.sOS)
        If lErro <> SUCESSO Then gError 141395
    
    End If
   
    objCT.Show

    Exit Sub

Erro_objEventoOS_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function OS_Atualiza_Linha(ByVal objCT As Object, ByVal iLinha As Integer, ByVal sOS As String) As Long

Dim lErro As Long, dQtde As Double, objDan_ItensOS As New ClassDan_ItensOS
Dim iPreenchido As Integer, sProduto As String, sProduto1 As String, sOSAnterior As String
Dim dQuantAnterior As Double

On Error GoTo Erro_OS_Atualiza_Linha

    'pegar a qtde default do produto da tabela Dan_ItensOS
    
    sProduto1 = objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col)
    
    'Verifica se o Produto está preenchido
    If Len(Trim(sProduto1)) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 184856
        
        If iPreenchido <> PRODUTO_VAZIO Then
        
            objDan_ItensOS.sOS = sOS
            objDan_ItensOS.sProduto = sProduto
            
            lErro = CF("Dan_ItemOS_Le_OSProd", objDan_ItensOS)
            If lErro <> SUCESSO And lErro <> 184855 Then gError 184857
            
            If lErro = SUCESSO Then
                
                dQuantAnterior = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Quantidade_Col))
                
                If Abs(dQuantAnterior - objDan_ItensOS.dQuantidade) > 0.00001 Then
                
                    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Quantidade_Col) = Formata_Estoque(objDan_ItensOS.dQuantidade)
                    
                    '###########################################
                    'Inserido por Wagner 23/03/2006
                    lErro = CF2(Me, "OV_Calcula_PesoKg", iLinha)
                    If lErro <> SUCESSO Then gError 177171
                    '###########################################
                
                    'Recalcula preço do ítem e valor total da nota
                    Call objCT.PrecoTotal_Calcula(iLinha)
                    lErro = objCT.ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 84162 '51037
                    lErro = objCT.FormPreco_Calcula()
                    If lErro <> SUCESSO Then gError 84162 '51037
                
                End If
                
            End If
            
        End If
        
    End If

    OS_Atualiza_Linha = SUCESSO
    
    Exit Function
    
Erro_OS_Atualiza_Linha:

    OS_Atualiza_Linha = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184859)

    End Select
    
    Exit Function

End Function

