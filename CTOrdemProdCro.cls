VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrdemProdCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MAQUINA_UM = 0
Const MAQUINA_DOIS = 1

Const FRAME_CARGA = 3

Private WithEvents objEventoProdutoBase As AdmEvento
Attribute objEventoProdutoBase.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoMaquinaCro As AdmEvento
Attribute objEventoMaquinaCro.VB_VarHelpID = -1

Private gobjCT As Object

Public Property Get Fornecedor(ByVal objCT As Object) As Object
     Set Fornecedor = objCT.objUserControl.Controls("Fornecedor")
End Property

Public Property Get Maquina2(ByVal objCT As Object) As Object
     Set Maquina2 = objCT.objUserControl.Controls("Maquina2")
End Property

Public Property Get Item(ByVal objCT As Object) As Object
     Set Item = objCT.objUserControl.Controls("Item")
End Property

Public Property Get ProdutoBase(ByVal objCT As Object) As Object
     Set ProdutoBase = objCT.objUserControl.Controls("ProdutoBase")
End Property

Public Property Get Descricao(ByVal objCT As Object) As Object
     Set Descricao = objCT.objUserControl.Controls("Descricao")
End Property

Public Property Get KgProdBase(ByVal objCT As Object) As Object
     Set KgProdBase = objCT.objUserControl.Controls("KgProdBase")
End Property

Public Property Get KgCarga(ByVal objCT As Object) As Object
     Set KgCarga = objCT.objUserControl.Controls("KgCarga")
End Property

Public Property Get QtdCarga(ByVal objCT As Object) As Object
     Set QtdCarga = objCT.objUserControl.Controls("QtdCarga")
End Property

Public Property Get QtdTotal(ByVal objCT As Object) As Object
     Set QtdTotal = objCT.objUserControl.Controls("QtdTotal")
End Property

Public Property Get Mistura(ByVal objCT As Object) As Object
     Set Mistura = objCT.objUserControl.Controls("Mistura")
End Property

Public Property Get TempoMistura(ByVal objCT As Object) As Object
     Set TempoMistura = objCT.objUserControl.Controls("TempoMistura")
End Property

Public Property Get CargaDrays(ByVal objCT As Object) As Object
     Set CargaDrays = objCT.objUserControl.Controls("CargaDrays")
End Property

Public Property Get Embalagem(ByVal objCT As Object) As Object
     Set Embalagem = objCT.objUserControl.Controls("Embalagem")
End Property

Public Property Get GridInsumo(ByVal objCT As Object) As Object
     Set GridInsumo = objCT.objUserControl.Controls("GridInsumo")
End Property

Public Property Get CodInsumo(ByVal objCT As Object) As Object
     Set CodInsumo = objCT.objUserControl.Controls("CodInsumo")
End Property

Public Property Get DescInsumo(ByVal objCT As Object) As Object
     Set DescInsumo = objCT.objUserControl.Controls("DescInsumo")
End Property

Public Property Get KgCargaInsumo(ByVal objCT As Object) As Object
     Set KgCargaInsumo = objCT.objUserControl.Controls("KgCargaInsumo")
End Property

Public Property Get Lote(ByVal objCT As Object) As Object
     Set Lote = objCT.objUserControl.Controls("Lote")
End Property

Public Property Get GridCarga(ByVal objCT As Object) As Object
     Set GridCarga = objCT.objUserControl.Controls("GridCarga")
End Property

Public Property Get Agua(ByVal objCT As Object) As Object
     Set Agua = objCT.objUserControl.Controls("Agua")
End Property

Public Property Get Oleo(ByVal objCT As Object) As Object
     Set Oleo = objCT.objUserControl.Controls("Oleo")
End Property

Public Property Get Amperagem(ByVal objCT As Object) As Object
     Set Amperagem = objCT.objUserControl.Controls("Amperagem")
End Property

Public Property Get VRDe(ByVal objCT As Object) As Object
     Set VRDe = objCT.objUserControl.Controls("VRDe")
End Property

Public Property Get VRAte(ByVal objCT As Object) As Object
     Set VRAte = objCT.objUserControl.Controls("VRAte")
End Property

Public Property Get Gate(ByVal objCT As Object) As Object
     Set Gate = objCT.objUserControl.Controls("Gate")
End Property

Public Property Get TCZonaDe(ByVal iIndex As Integer, ByVal objCT As Object) As Object
     Set TCZonaDe = objCT.objUserControl.Controls("TCZonaDe")(iIndex)
End Property

Public Property Get TCZonaAte(ByVal iIndex As Integer, ByVal objCT As Object) As Object
     Set TCZonaAte = objCT.objUserControl.Controls("TCZonaAte")(iIndex)
End Property

Public Property Get Tela(ByVal objCT As Object) As Object
     Set Tela = objCT.objUserControl.Controls("Tela")
End Property

Public Property Get TEZonaDe(ByVal iIndex As Integer, ByVal objCT As Object) As Object
     Set TEZonaDe = objCT.objUserControl.Controls("TEZonaDe")(iIndex)
End Property

Public Property Get TEZonaAte(ByVal iIndex As Integer, ByVal objCT As Object) As Object
     Set TEZonaAte = objCT.objUserControl.Controls("TEZonaAte")(iIndex)
End Property

Public Property Get VVDe(ByVal objCT As Object) As Object
     Set VVDe = objCT.objUserControl.Controls("VVDe")
End Property

Public Property Get VVAte(ByVal objCT As Object) As Object
     Set VVAte = objCT.objUserControl.Controls("VVAte")
End Property

Public Property Get Observacao(ByVal objCT As Object) As Object
     Set Observacao = objCT.objUserControl.Controls("ObservacaoCro")
End Property

Public Property Get Amostras(ByVal objCT As Object) As Object
     Set Amostras = objCT.objUserControl.Controls("Amostras")
End Property

Public Property Get OpcaoAmbos(ByVal objCT As Object) As Object
     Set OpcaoAmbos = objCT.objUserControl.Controls("OpcaoAmbos")
End Property

Public Property Get OpcaoSimples(ByVal objCT As Object) As Object
     Set OpcaoSimples = objCT.objUserControl.Controls("OpcaoSimples")
End Property

Public Property Get OpcaoDetalhado(ByVal objCT As Object) As Object
     Set OpcaoDetalhado = objCT.objUserControl.Controls("OpcaoDetalhado")
End Property

Public Function OP_Form_UnLoad(ByVal objCT As Object, Cancel As Integer) As Long

    Set objCT.gobjInfoUsu.objGridCarga = Nothing
    Set objCT.gobjInfoUsu.objGridInsumo = Nothing
    Set objCT.gobjInfoUsu.gcolItemOPCarga = Nothing
    
    Set objEventoProdutoBase = Nothing
    Set objEventoMaquinaCro = Nothing
    Set objEventoLote = Nothing
    Set objEventoFornecedor = Nothing
    
    OP_Form_UnLoad = SUCESSO
    
End Function

Function OP_Inicializa_GridMovimentos(ByVal objCT As Object) As Long

Dim iIndice As Integer, objGridInt As AdmGrid

    Set objGridInt = New AdmGrid
    
    Set objCT.objGrid = objGridInt

    'tela em questão
    Set objGridInt.objForm = objCT

    'titulos do grid
    With objGridInt
        .colColuna.Add ("")
        .colColuna.Add ("Produto")
        .colColuna.Add ("Versão")
        .colColuna.Add ("Descrição")
        .colColuna.Add ("U.M.")
        .colColuna.Add ("Quantidade")
        .colColuna.Add ("Almoxarifado")
        .colColuna.Add ("Benef.")
        .colColuna.Add ("Ccl")
        .colColuna.Add ("Previsão Início")
        .colColuna.Add ("Previsão Fim")
        .colColuna.Add ("Situação")
        .colColuna.Add ("Destinação")
        .colColuna.Add ("Pedido de Venda")
        .colColuna.Add ("Filial do Pedido")
        .colColuna.Add ("Cliente")
        .colColuna.Add ("Filial do Cliente")
        .colColuna.Add ("Prioridade")
        .colColuna.Add ("Maquina 1")
        '#########################################
        'Inserido por Wagner - CROMATON 03/11/04
        .colColuna.Add ("Maquina 2")
        '#########################################
        .colColuna.Add ("Produz Logo")

        'Controles que participam do Grid
        .colCampo.Add (objCT.Produto.Name)
        .colCampo.Add (objCT.Versao.Name)
        .colCampo.Add (objCT.DescricaoItem.Name)
        .colCampo.Add (objCT.UnidadeMed.Name)
        .colCampo.Add (objCT.Quantidade.Name)
        .colCampo.Add (objCT.Almoxarifado.Name)
        .colCampo.Add (objCT.Benef.Name)
        .colCampo.Add (objCT.Ccl.Name)
        .colCampo.Add (objCT.DataPrevisaoInicio.Name)
        .colCampo.Add (objCT.DataPrevisaoFim.Name)
        .colCampo.Add (objCT.Situacao.Name)
        .colCampo.Add (objCT.Destinacao.Name)
        .colCampo.Add (objCT.PedidoDeVendaId.Name)
        .colCampo.Add (objCT.ComboFilialPedido.Name)
        .colCampo.Add (objCT.Cliente.Name)
        .colCampo.Add (objCT.FilialCliente.Name)
        .colCampo.Add (objCT.Prioridade.Name)
        .colCampo.Add (objCT.Maquina.Name)
        '#########################################
        'Inserido por Wagner - CROMATON 03/11/04
        .colCampo.Add (Maquina2(objCT).Name)
        '#########################################
        .colCampo.Add (objCT.ProduzLogo.Name)
    End With

    'Colunas do Grid
    With objCT
        .iGrid_Produto_Col = 1
        .iGrid_Versao_Col = 2
        .iGrid_DescricaoItem_Col = 3
        .iGrid_UnidadeMed_Col = 4
        .iGrid_Quantidade_Col = 5
        .iGrid_Almoxarifado_Col = 6
        .iGrid_Benef_Col = 7
        .iGrid_Ccl_Col = 8
        .iGrid_DataPrevInicio_Col = 9
        .iGrid_DataPrevFim_Col = 10
        .iGrid_Situacao_Col = 11
        .iGrid_Destinacao_Col = 12
        .iGrid_PedidoDeVenda_Col = 13
        .iGrid_FilialPedido_Col = 14
        .iGrid_Cliente_Col = 15
        .iGrid_FilialCliente_Col = 16
        .iGrid_Prioridade_Col = 17
        .iGrid_Maquina_Col = 18
        '#########################################
        'Inserido por Wagner - CROMATON 03/11/04
        objCT.gobjInfoUsu.iGrid_Maquina2_Col = 19
        '#########################################
        .iGrid_ProduzLogo_Col = 20
    End With
    
    objGridInt.objGrid = objCT.GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    objCT.GridMovimentos.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGridInt)

    OP_Inicializa_GridMovimentos = SUCESSO

End Function

Public Function OP_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    
'Inserido por Wagner
Dim lErro As Long
Dim bItemPreenchido As Boolean
Dim sProdutoBaseFormatado As String
Dim iProdutoBasePreenchido As Integer
Dim sCodProdutoBase As String
Dim bQtdPreenchida As Boolean
Dim objInsumo As New ClassProduto
Dim iInsumoPreenchido As Integer
Dim sInsumoFormatado As String
Dim sCodInsumo As String
    
On Error GoTo Erro_OP_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then
    
        '#################################################
        'Inserido por Wagner 13/02/2006
        'Verifica se produto está preenchido
        If GridInsumo(objCT).Row > 0 Then
            sCodInsumo = GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Insumo_Col)
        End If
        
        lErro = CF("Produto_Formata", sCodInsumo, sInsumoFormatado, iInsumoPreenchido)
        If lErro <> SUCESSO Then gError 141769
    
        If iInsumoPreenchido = PRODUTO_PREENCHIDO Then
        
            objInsumo.sCodigo = sInsumoFormatado
    
            lErro = CF("Produto_Le", objInsumo)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 141770
        
        End If
        '#################################################
        
        '#########################################
        'Inserido por Wagner - CROMATON 03/11/04
        'Verifica se o item da carga está preenchido
        If Len(Trim(GridCarga(objCT).TextMatrix(GridCarga(objCT).Row, objCT.gobjInfoUsu.iGrid_Item_Col))) > 0 Then
            bItemPreenchido = True
        Else
            bItemPreenchido = False
        End If
        
        If Len(Trim(objCT.GridMovimentos.TextMatrix(GridCarga(objCT).Row, objCT.iGrid_Quantidade_Col))) > 0 Then
            bQtdPreenchida = True
        Else
            bQtdPreenchida = False
        End If
        
        'Verifica se produto está preenchido
        If GridCarga(objCT).Row > 0 Then
            sCodProdutoBase = GridCarga(objCT).TextMatrix(GridCarga(objCT).Row, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col)
        End If
        
        lErro = CF("Produto_Formata", sCodProdutoBase, sProdutoBaseFormatado, iProdutoBasePreenchido)
        If lErro <> SUCESSO Then gError 131264
        '#########################################
    
        'Pesquisa o controle da coluna em questão
        Select Case objControl.Name
        
            '#########################################
            'Inserido por Wagner - CROMATON 08/11/04
            Case ProdutoBase(objCT).Name
    
                objControl.Enabled = True
            
            Case KgProdBase(objCT).Name
                
                If iProdutoBasePreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = bQtdPreenchida
                Else
                    objControl.Enabled = False
                End If
            
            Case QtdCarga(objCT).Name
    
                If iProdutoBasePreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = bItemPreenchido
                Else
                    objControl.Enabled = False
                End If
                           
            Case KgCarga(objCT).Name, Item(objCT).Name, Descricao(objCT).Name, QtdTotal(objCT).Name
    
                objControl.Enabled = False
                
            'Case Fornecedor.Name
            
                'If objInsumo.iRastro = PRODUTO_RASTRO_NENHUM Then
                '    objControl.Enabled = False
                'Else
                '    objControl.Enabled = True
                'End If
                
            Case Lote(objCT).Name
                'If Len(Trim(GridInsumo.TextMatrix(GridInsumo.Row, iGrid_Fornecedor_Col))) = 0 Then
                '    objControl.Enabled = False
                'Else
                    If objInsumo.iRastro = PRODUTO_RASTRO_NENHUM Then
                        objControl.Enabled = False
                    Else
                        objControl.Enabled = True
                    End If
                'End If
                
            Case DescInsumo(objCT).Name, CodInsumo(objCT).Name, KgCargaInsumo(objCT).Name, Fornecedor(objCT).Name
            
                objControl.Enabled = False
            '#########################################
    
        End Select
    
    End If
    
    OP_Rotina_Grid_Enable = SUCESSO
    
    Exit Function

Erro_OP_Rotina_Grid_Enable:

    OP_Rotina_Grid_Enable = gErr
    
    Select Case gErr

        Case 131264, 141769, 141770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163800)

    End Select

    Exit Function

End Function

Public Sub BotaoCarga_Click(ByVal objCT As Object)
    
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoCarga_Click
    
    For iIndice = 1 To objCT.gobjInfoUsu.gcolItemOPCarga.Count
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iIndice).iItem = iIndice
    Next
    
    objCT.iAlterado = REGISTRO_ALTERADO
    
    lErro = Preenche_GridCarga(objCT)
    If lErro <> SUCESSO Then gError 131211
    
    Exit Sub

Erro_BotaoCarga_Click:

   Select Case gErr
   
        Case 131211
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Sub
    
End Sub

Public Sub BotaoLimpaCarga_Click(ByVal objCT As Object)

Dim lErro As Long
Dim iIndice As Integer
Dim iLinhasExistentesAnt As Integer

On Error GoTo Erro_BotaoLimpaCarga_Click

    'iLinhasExistentesAnt = objGridCarga.iLinhasExistentes

    Call Limpa_GridCarga(objCT)
    Call Limpa_GridInsumo(objCT)
    
    'objGridCarga.iLinhasExistentes = iLinhasExistentesAnt
    
    For iIndice = 1 To objCT.gobjInfoUsu.gcolItemOPCarga.Count
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iIndice).iItem = 0
    Next

    Exit Sub
    
Erro_BotaoLimpaCarga_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub BotaoMaquinas_Click(ByVal objCT As Object)
'Chama o Browser de Maquinas...

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Dim sMaquina As String
'#########################################

On Error GoTo Erro_BotaoMaquinas_Click

    'Verifica se tem alguma linha selecionada no Grid
    If objCT.GridMovimentos.Row = 0 Then gError 106320

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Identifica qual equipamento está sendo alterado
    If objCT.gobjInfoUsu.iMaquina = MAQUINA_UM Then
        sMaquina = objCT.Maquina.Text
    Else
        sMaquina = Maquina2(objCT).Text
    End If
    '#########################################

    'Se o equipamento foi preenchido => armazena no obj
    If Len(Trim(sMaquina)) > 0 Then
    
        If IsNumeric(sMaquina) Then
        
            'Alterado por Wagner
            If objCT.gobjInfoUsu.iMaquina = MAQUINA_UM Then
                objMaquina.iCodigo = StrParaInt(objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Maquina_Col))
            Else
                objMaquina.iCodigo = StrParaInt(objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.gobjInfoUsu.iGrid_Maquina2_Col))
            End If
        
        Else
            
            'Alterado por Wagner
            If objCT.gobjInfoUsu.iMaquina = MAQUINA_UM Then
                objMaquina.sNomeReduzido = CStr(objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Maquina_Col))
            Else
                objMaquina.sNomeReduzido = CStr(objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.gobjInfoUsu.iGrid_Maquina2_Col))
            End If
        
        End If
    
    End If
    
    Set objEventoMaquinaCro = New AdmEvento

    Set gobjCT = objCT
    'Lista de Equipamentos
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquinaCro)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr

        Case 106320
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 106433
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NOME_INEXISTENTE", gErr, sMaquina)
            
        Case 106434
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_CODIGO_INEXISTENTE", gErr, sMaquina, giFilialEmpresa)

        Case 55325, 106432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Sub

Public Function OP_GridMovimentos_KeyDown(ByVal objCT As Object, ByVal iLinhaAnterior As Integer) As Long
            
    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Mantén a integridade entre gcolItemOP e gcolItemOPCarga
    objCT.gobjInfoUsu.gcolItemOPCarga.Remove (iLinhaAnterior)
    '#########################################
    
    OP_GridMovimentos_KeyDown = SUCESSO
    
End Function

Private Sub objEventoMaquinaCro_evSelecao(obj1 As Object)

Dim objMaquinas As ClassMaquinas
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long, objCT As Object

On Error GoTo Erro_objEventoMaquina_evSelecao

    Set objMaquinas = obj1
    Set objCT = gobjCT
    
    sCodProduto = objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106426

    'verifica se o produto esta preenchido...
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        'Alterado por Wagner
        'Identifica qual equipamento deve ser preenchido
        If objCT.gobjInfoUsu.iMaquina = MAQUINA_UM Then
            objCT.Maquina.Text = objMaquinas.sNomeReduzido
            objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Maquina_Col) = objMaquinas.sNomeReduzido
        Else
            Maquina2(objCT).Text = objMaquinas.sNomeReduzido
            objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.gobjInfoUsu.iGrid_Maquina2_Col) = objMaquinas.sNomeReduzido
        End If
    
    End If
    
    objCT.Show
    
    Exit Sub
    
Erro_objEventoMaquina_evSelecao:

    Select Case gErr
        
        Case 106426
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select

End Sub

Function OP_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_OP_Form_Load

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    Set objEventoProdutoBase = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoLote = New AdmEvento
    
    objCT.gobjInfoUsu.iMaquina = MAQUINA_UM

    Call CarregaCombo_Mistura(Mistura(objCT))
    Call CarregaCombo_CargaDrays(CargaDrays(objCT))
'    Call CarregaCombo_Embalagem(Embalagem(objCT))

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoBase(objCT))
    If lErro <> SUCESSO Then gError 131071

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodInsumo(objCT))
    If lErro <> SUCESSO Then gError 131072
    '#########################################

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'inicializa Grid
    lErro = Inicializa_GridCarga(objCT)
    If lErro <> SUCESSO Then gError 131073
    
    'inicializa Grid
    lErro = Inicializa_GridInsumo(objCT)
    If lErro <> SUCESSO Then gError 131074
    '#########################################
    
    'Inserido por Jorge Specian - 11/07/2005 - Tamanho do campo Codigo configurável
    '---------------------------------------
    objCT.Codigo.MaxLength = STRING_ORDEM_DE_PRODUCAO
    objCT.CodOPGerada.MaxLength = STRING_ORDEM_DE_PRODUCAO
    '---------------------------------------
    
    OP_Form_Load = SUCESSO

    Exit Function

Erro_OP_Form_Load:

    OP_Form_Load = gErr

    Select Case gErr

        Case 131071 To 131074
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Function OP_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OP_Saida_Celula

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    If objGridInt.objGrid Is GridCarga(objCT) Then

        Select Case GridCarga(objCT).Col
      
            Case objCT.gobjInfoUsu.iGrid_ProdutoBase_Col
      
                lErro = Saida_Celula_ProdutoBase(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 131077
                  
            Case objCT.gobjInfoUsu.iGrid_KgProdBase_Col
      
                lErro = Saida_Celula_KgProdBase(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 131078
                                      
             Case objCT.gobjInfoUsu.iGrid_QtdCarga_Col
      
                lErro = Saida_Celula_QtdCarga(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 131080
      
        End Select
      
    End If
      
    If objGridInt.objGrid Is GridInsumo(objCT) Then

        Select Case GridInsumo(objCT).Col
                  
            Case objCT.gobjInfoUsu.iGrid_Fornecedor_Col
      
                lErro = Saida_Celula_Fornecedor(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 131083
      
            Case objCT.gobjInfoUsu.iGrid_Lote_Col
      
                lErro = Saida_Celula_Lote(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 131083
    
        End Select
      
    End If
    
    If objGridInt.objGrid Is objCT.GridMovimentos Then
    
        Select Case objCT.GridMovimentos.Col
                  
            Case objCT.gobjInfoUsu.iGrid_Maquina2_Col
      
                lErro = Saida_Celula_Maquina2(objCT, objGridInt)
                If lErro <> SUCESSO Then gError 131083
    
        End Select
      
    End If
    '#########################################
    
    OP_Saida_Celula = SUCESSO

    Exit Function

Erro_OP_Saida_Celula:

    OP_Saida_Celula = gErr

    Select Case gErr

        Case 131077, 131078, 131080, 131083
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Public Function OP_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long

    If KeyCode = KEYCODE_BROWSER Then
        If objCT.ActiveControl Is Maquina2(objCT) Then 'Inserido por Wagner
            objCT.gobjInfoUsu.iMaquina = MAQUINA_DOIS
            Call BotaoMaquinas_Click(objCT)
        ElseIf objCT.ActiveControl Is ProdutoBase(objCT) Then 'Inserido por Wagner
            Call BotaoProdutoBase_Click(objCT)
        ElseIf objCT.ActiveControl Is Fornecedor(objCT) Then 'Inserido por Wagner
            Call BotaoFornecedor_Click(objCT)
        ElseIf objCT.ActiveControl Is Lote(objCT) Then 'Inserido por Wagner
            Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoLote_Click(objCT)
        End If
    End If

End Function

'#########################################
'Inserido por Wagner - CROMATON 08/11/04

Public Function OP_TabStrip1_Click(ByVal objCT As Object) As Long

    '########################################
    'Inserido por Wagner - CROMATON 04/11/04
    If objCT.iFrameAtual = FRAME_CARGA Then
        'Preenche o grid de carga com a coleção global gcolItemOPCarga
        Call Preenche_GridCarga(objCT)
    End If
    '########################################
    
    OP_TabStrip1_Click = SUCESSO

End Function

Public Function OP_Saida_Celula_Versao(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'faz a critica da celula de Versao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_OP_Saida_Celula_Versao

    If objCT.Versao.ListIndex >= 0 Then
            
        '########################################
        'Inserido por Wagner - CROMATON 04/11/04
        lErro = OP_GridMov_AtualizaLinha(objCT)
        If lErro <> SUCESSO Then gError 131090
        '########################################
        
        objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Versao_Col) = objCT.Versao.Text
        
    End If

    OP_Saida_Celula_Versao = SUCESSO

    Exit Function

Erro_OP_Saida_Celula_Versao:

    OP_Saida_Celula_Versao = gErr

    Select Case gErr

        Case 131090
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Public Sub Maquina_GotFocus(ByVal objCT As Object)

    objCT.gobjInfoUsu.iMaquina = MAQUINA_UM
    Call Grid_Campo_Recebe_Foco(objCT.objGrid)

End Sub

Public Function OP_Executa_Relatorio(ByVal objCT As Object, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Executa o(s) relatorio(s) de acordo com a selecao no frame de relatorios

Dim lErro As Long, lNumIntRel As Long
Dim objRelatorio1 As New AdmRelatorio, objRelatorio2 As New AdmRelatorio
Dim objRelatorio3 As New AdmRelatorio
Dim objItemOP As ClassItemOP
Dim objItemOPCarga As ClassItemOPCarga

On Error GoTo Erro_OP_Executa_Relatorio
    
    '##########################################
    'Inserido por Wagner 26/09/05
    For Each objItemOP In objOrdemDeProducao.colItens
    
        For Each objItemOPCarga In objCT.gobjInfoUsu.gcolItemOPCarga
    
            If objItemOPCarga.lNumIntItemOP = objItemOP.lNumIntDoc Then
            
                Set objItemOP.objInfoUsu = objItemOPCarga
                Exit For
            End If
    
        Next
    
    Next
    '##########################################

    'Executa o Relatorio de OP caso o usuário tenha marcado somente o relatório de OP ou ambos
    If objCT.OpcaoRelatorio.Item(0).Value = True Or objCT.OpcaoRelatorio.Item(2).Value = True Then
    
        If OpcaoSimples(objCT).Value Or OpcaoAmbos(objCT).Value Then
        
            lErro = CF("RelOPCarga_Prepara", objOrdemDeProducao, lNumIntRel)
            If lErro <> SUCESSO Then gError 111828
        
            objRelatorio1.iLandscape = 1
        
            'Imprime o Relatorio de OP
            lErro = objRelatorio1.ExecutarDireto("Ordens de Produção", "NumIntRel = @NNUMINTREL", 0, "OPCargas", "NNUMINTREL", CStr(lNumIntRel), "DTDATAOP", CStr(objOrdemDeProducao.dtDataEmissao))
            If lErro <> SUCESSO Then gError 106406
            
            objRelatorio3.iLandscape = 1
            
            lErro = objRelatorio3.ExecutarDireto("Ficha de Processo", "OrdemProducao = @TORDPROD", 0, "FICHAPRO", "TORDPROD", objOrdemDeProducao.sCodigo)
            If lErro <> SUCESSO Then gError 106406
            
        End If
        
        If OpcaoDetalhado(objCT).Value Or OpcaoAmbos(objCT).Value Then

            lErro = CF("RelOPCarga_Prepara", objOrdemDeProducao, lNumIntRel, True)
            If lErro <> SUCESSO Then gError 111828

            objRelatorio1.iLandscape = 1
        
            'Imprime o Relatorio de OP
            lErro = objRelatorio1.ExecutarDireto("Espelho da OP", "NumIntRel = @NNUMINTREL", 0, "OPEspelh", "NNUMINTREL", CStr(lNumIntRel), "DTDATAOP", CStr(objOrdemDeProducao.dtDataEmissao))
            If lErro <> SUCESSO Then gError 106406
            
        End If
    
    End If
    
    'Executa o Relatório de Rótulos caso o usuário tenha marcado somente o relatório de Rótulos ou ambos
    If objCT.OpcaoRelatorio.Item(1).Value = True Or objCT.OpcaoRelatorio.Item(2).Value = True Then
    
        'Imprime o Relatório de Rótulos
        lErro = objRelatorio2.ExecutarDireto("Rótulos para Ordens de Produção", "OrdemProducao = @TORDPROD", 0, "SOLROTUL", "TORDPROD", objOrdemDeProducao.sCodigo)
        If lErro <> SUCESSO Then gError 106408
        
    End If
    
    OP_Executa_Relatorio = SUCESSO
    
    Exit Function
    
Erro_OP_Executa_Relatorio:

    OP_Executa_Relatorio = gErr
    
    Select Case gErr
    
        Case 106406, 106408, 111828
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select

End Function

'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Private Function CarregaCombo_Mistura(objCombo As ComboBox) As Long

On Error GoTo Erro_CarregaCombo_Mistura

    objCombo.AddItem (OP_FICHAPROC_MISTURA_TAMBOR & SEPARADOR & STRING_OP_FICHAPROC_MISTURA_TAMBOR)
    objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_TAMBOR
    
    objCombo.AddItem (OP_FICHAPROC_MISTURA_HENCHEL & SEPARADOR & STRING_OP_FICHAPROC_MISTURA_HENCHEL)
    objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_HENCHEL
    
    CarregaCombo_Mistura = SUCESSO

    Exit Function
    
Erro_CarregaCombo_Mistura:

    CarregaCombo_Mistura = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
    
    End Select
    
    Exit Function

End Function

Private Function CarregaCombo_TempoMistura(objCombo As ComboBox, iMistura As Integer) As Long

On Error GoTo Erro_CarregaCombo_TempoMistura

    If iMistura = OP_FICHAPROC_MISTURA_TAMBOR Then

        objCombo.AddItem (STRING_OP_FICHAPROC_MISTURA_TAMBOR_5_15)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_TAMBOR_5_15
        
        objCombo.AddItem (STRING_OP_FICHAPROC_MISTURA_TAMBOR_10_20)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_TAMBOR_10_20
    
    Else
    
        objCombo.AddItem (OP_FICHAPROC_MISTURA_HENCHEL_3)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_HENCHEL_3
        
        objCombo.AddItem (OP_FICHAPROC_MISTURA_HENCHEL_4)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_HENCHEL_4
    
        objCombo.AddItem (OP_FICHAPROC_MISTURA_HENCHEL_5)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_MISTURA_HENCHEL_5
    
    End If
    
    CarregaCombo_TempoMistura = SUCESSO

    Exit Function
    
Erro_CarregaCombo_TempoMistura:

    CarregaCombo_TempoMistura = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
    
    End Select
    
    Exit Function

End Function

Private Function CarregaCombo_CargaDrays(objCombo As ComboBox) As Long

On Error GoTo Erro_CarregaCombo_CargaDrays

    objCombo.AddItem (STRING_OP_FICHAPROC_CARGADRAYS_4_6)
    objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_CARGADRAYS_4_6
    
    objCombo.AddItem (STRING_OP_FICHAPROC_CARGADRAYS_BRANCO)
    objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_CARGADRAYS_BRANCO
    
    CarregaCombo_CargaDrays = SUCESSO

    Exit Function
    
Erro_CarregaCombo_CargaDrays:

    CarregaCombo_CargaDrays = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
    
    End Select
    
    Exit Function

End Function

Private Function CarregaCombo_Embalagem(ByVal objCT As Object, ByVal objCombo As ComboBox, ByVal sProduto As String) As Long

Dim lErro As Long, objProduto As New ClassProduto
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim objEmbalagem As New ClassEmbalagem

On Error GoTo Erro_CarregaCombo_Embalagem

    If objCT.gobjInfoUsu.sProdComboEmbalagens <> sProduto Then
    
        objCombo.Clear
        objCT.gobjInfoUsu.sProdComboEmbalagens = ""
        
        objCombo.AddItem (OP_FICHAPROC_EMBALAGEM_PEQUENA & SEPARADOR & STRING_OP_FICHAPROC_EMBALAGEM_PEQUENA)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_EMBALAGEM_PEQUENA
        
        objCombo.AddItem (OP_FICHAPROC_EMBALAGEM_GRANDE & SEPARADOR & STRING_OP_FICHAPROC_EMBALAGEM_GRANDE)
        objCombo.ItemData(objCombo.NewIndex) = OP_FICHAPROC_EMBALAGEM_GRANDE
        
        'Le as embalagens já relacionadas com o produto
        objProduto.sCodigo = sProduto
        lErro = CF("ProdutoEmbalagem_Le_Produto", objProduto)
        If lErro <> SUCESSO And lErro <> 95016 Then gError 95008
        
        If lErro = SUCESSO Then
        
            For Each objProdutoEmbalagem In objProduto.colProdutoEmbalagem
            
                Select Case objProdutoEmbalagem.iEmbalagem
                                
                    Case OP_FICHAPROC_EMBALAGEM_PEQUENA
                        objCombo.ListIndex = 0
                        
                    Case OP_FICHAPROC_EMBALAGEM_GRANDE
                        objCombo.ListIndex = 1
                        
                    Case Else
                
                        'Guarda em objEmbalagem o código da embalagem que está em objProdutoEmbalagem
                        objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
                        
                        'Le do BD a descricao da embalagem
                        lErro = CF("Embalagem_Le", objEmbalagem)
                        If lErro <> SUCESSO And lErro <> 82763 Then gError 95048
                    
                        If lErro = SUCESSO Then
                            objCombo.AddItem objProdutoEmbalagem.iEmbalagem & SEPARADOR & objEmbalagem.sDescricao
                            objCombo.ItemData(objCombo.NewIndex) = objProdutoEmbalagem.iEmbalagem
                            If objProdutoEmbalagem.iPadrao <> 0 Then
                                objCombo.ListIndex = objCombo.NewIndex
                            End If
                        End If
                        
                End Select
                
            Next
            
        End If
    
    End If
    
    objCT.gobjInfoUsu.sProdComboEmbalagens = sProduto
    
    CarregaCombo_Embalagem = SUCESSO

    Exit Function
    
Erro_CarregaCombo_Embalagem:

    CarregaCombo_Embalagem = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
    
    End Select
    
    Exit Function

End Function

Public Sub Mistura_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Mistura_Validate

    If Len(Trim(Mistura(objCT).Text)) = 0 Then Exit Sub

    If Mistura(objCT).Text = Mistura(objCT).List(Mistura(objCT).ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Mistura(objCT), iCodigo)
    If lErro <> SUCESSO Then gError 131093
    
    TempoMistura(objCT).Clear
    
    lErro = CarregaCombo_TempoMistura(TempoMistura(objCT), iCodigo)
    If lErro <> SUCESSO Then gError 131094

    Exit Sub

Erro_Mistura_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131094
    
        Case 131093
            Call Rotina_Erro(vbOKOnly, "ERRO_MISTURA_NAO_CADASTRADA", gErr, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Mistura_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TempoMistura_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TempoMistura_Validate

    If Len(Trim(TempoMistura(objCT).Text)) = 0 Then Exit Sub

    If TempoMistura(objCT).Text = TempoMistura(objCT).List(TempoMistura(objCT).ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TempoMistura(objCT), iCodigo)
    If lErro <> SUCESSO Then gError 131095

    Exit Sub

Erro_TempoMistura_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131095
            Call Rotina_Erro(vbOKOnly, "ERRO_TEMPOMISTURA_NAO_CADASTRADO", gErr, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub TempoMistura_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CargaDrays_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_CargaDrays_Validate

    If Len(Trim(CargaDrays(objCT).Text)) = 0 Then Exit Sub

    If CargaDrays(objCT).Text = CargaDrays(objCT).List(CargaDrays(objCT).ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CargaDrays(objCT), iCodigo)
    If lErro <> SUCESSO Then gError 131096

    Exit Sub

Erro_CargaDrays_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131096
            Call Rotina_Erro(vbOKOnly, "ERRO_CARGADRAYS_NAO_CADASTRADO", gErr, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub CargaDrays_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Embalagem_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Embalagem_Validate

    If Len(Trim(Embalagem(objCT).Text)) = 0 Then Exit Sub

    If Embalagem(objCT).Text = Embalagem(objCT).List(Embalagem(objCT).ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Embalagem(objCT), iCodigo)
    If lErro <> SUCESSO Then gError 131097

    Exit Sub

Erro_Embalagem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131097
            Call Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_CADASTRADO", gErr, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Embalagem_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Inicializa_GridCarga(ByVal objCT As Object) As Long

Dim iIndice As Integer, objGridInt As Object

    Set objGridInt = New AdmGrid
    Set objCT.gobjInfoUsu.objGridCarga = objGridInt

    'tela em questão
    Set objGridInt.objForm = objCT

    With objGridInt
        'titulos do grid
        .colColuna.Add ("")
        .colColuna.Add ("Item")
        .colColuna.Add ("Produto Base")
        .colColuna.Add ("Descrição")
        .colColuna.Add ("Kg Produto Base")
        .colColuna.Add ("Kg por Carga")
        .colColuna.Add ("Qtde Carga")
        .colColuna.Add ("Kg Total")
        
        'Controles que participam do Grid
        .colCampo.Add (Item(objCT).Name)
        .colCampo.Add (ProdutoBase(objCT).Name)
        .colCampo.Add (Descricao(objCT).Name)
        .colCampo.Add (KgProdBase(objCT).Name)
        .colCampo.Add (KgCarga(objCT).Name)
        .colCampo.Add (QtdCarga(objCT).Name)
        .colCampo.Add (QtdTotal(objCT).Name)
    End With
    
    With objCT.gobjInfoUsu
        'Colunas do Grid
        .iGrid_Item_Col = 1
        .iGrid_ProdutoBase_Col = 2
        .iGrid_Descricao_Col = 3
        .iGrid_KgProdBase_Col = 4
        .iGrid_KgCarga_Col = 5
        .iGrid_QtdCarga_Col = 6
        .iGrid_QtdTotal_Col = 7
    End With
    
    'objGridCarga.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    objGridInt.objGrid = GridCarga(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridCarga(objCT).ColWidth(0) = 200

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGridInt)

    Inicializa_GridCarga = SUCESSO

End Function

Private Function Inicializa_GridInsumo(ByVal objCT As Object) As Long

Dim iIndice As Integer, objGridInt As AdmGrid

    Set objGridInt = New AdmGrid
    Set objCT.gobjInfoUsu.objGridInsumo = objGridInt

    'tela em questão
    Set objGridInt.objForm = objCT

    With objGridInt
        'titulos do grid
        .colColuna.Add ("")
        .colColuna.Add ("Produto")
        .colColuna.Add ("Descrição")
        .colColuna.Add ("Kg Carga")
        '.colColuna.Add ("Fornecedor")
        .colColuna.Add ("Lote")
        
        'Controles que participam do Grid
        .colCampo.Add (CodInsumo(objCT).Name)
        .colCampo.Add (DescInsumo(objCT).Name)
        .colCampo.Add (KgCargaInsumo(objCT).Name)
        '.colCampo.Add (Fornecedor.Name)
        .colCampo.Add (Lote(objCT).Name)
    End With
    
    With objCT.gobjInfoUsu
        'Colunas do Grid
        .iGrid_Insumo_Col = 1
        .iGrid_DescInsumo_Col = 2
        .iGrid_KgCargaInsumo_Col = 3
        'iGrid_Fornecedor_Col = 4
        .iGrid_Lote_Col = 4
    End With
    
    objGridInt.objGrid = GridInsumo(objCT)
   
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridInsumo(objCT).ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objCT.gobjInfoUsu.objGridInsumo)

    Inicializa_GridInsumo = SUCESSO

End Function

'#######################################################################
'INÍCIO DO SCRIPT DO GRID
'#######################################################################
Public Sub GridCarga_GotFocus(ByVal objCT As Object)

    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub GridCarga_EnterCell(ByVal objCT As Object)

    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridCarga, objCT.iAlterado)

End Sub

Public Sub GridCarga_LeaveCell(ByVal objCT As Object)

    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub GridCarga_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridCarga, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridCarga, objCT.iAlterado)
    End If

    Call Preenche_GridInsumos(objCT, GridCarga(objCT).Row)

End Sub

Public Sub GridCarga_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)

Dim iLinha As Integer
Dim iLinhaAtual As Integer
Dim iLinhasExistentesAnterior As Integer

    iLinha = GridCarga(objCT).Row
    iLinhasExistentesAnterior = objCT.gobjInfoUsu.objGridCarga.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridCarga)

    iLinhaAtual = GridCarga(objCT).Row
       
    'iLinhasExistentesAnterior = objGridCarga.iLinhasExistentes
    
    If iLinhasExistentesAnterior <> objCT.gobjInfoUsu.objGridCarga.iLinhasExistentes Then
    'Foi excluído um item
        
        objCT.gobjInfoUsu.objGridCarga.iLinhasExistentes = iLinhasExistentesAnterior
        
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = 0
        
        Call Preenche_GridCarga(objCT)
    
    End If

End Sub

Public Sub GridCarga_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridCarga, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridCarga, objCT.iAlterado)
    End If


End Sub

Public Sub GridCarga_Validate(ByVal objCT As Object, Cancel As Boolean)

    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub GridCarga_RowColChange(ByVal objCT As Object)

    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub GridCarga_Scroll(ByVal objCT As Object)

    Call Grid_Scroll(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub GridInsumo_GotFocus(ByVal objCT As Object)

    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub GridInsumo_EnterCell(ByVal objCT As Object)

    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridInsumo, objCT.iAlterado)

End Sub

Public Sub GridInsumo_LeaveCell(ByVal objCT As Object)

    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub GridInsumo_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridInsumo, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridInsumo, objCT.iAlterado)
    End If

End Sub

Public Sub GridInsumo_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridInsumo)
    
End Sub

Public Sub GridInsumo_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridInsumo, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridInsumo, objCT.iAlterado)
    End If

End Sub

Public Sub GridInsumo_Validate(ByVal objCT As Object, Cancel As Boolean)

    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub GridInsumo_RowColChange(ByVal objCT As Object)

    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub GridInsumo_Scroll(ByVal objCT As Object)

    Call Grid_Scroll(objCT.gobjInfoUsu.objGridInsumo)

End Sub
'#######################################################################
'FIM DO SCRIPT DO GRID
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT DOS ITENS DO GRID
'#######################################################################
Public Sub ProdutoBase_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoBase_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub ProdutoBase_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub ProdutoBase_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridCarga.objControle = ProdutoBase(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCarga)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub KgProdBase_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub KgProdBase_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub KgProdBase_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub KgProdBase_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridCarga.objControle = KgProdBase(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCarga)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub KgCarga_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub KgCarga_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub KgCarga_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub KgCarga_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridCarga.objControle = KgCarga(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCarga)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QtdCarga_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QtdCarga_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub QtdCarga_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub QtdCarga_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridCarga.objControle = QtdCarga(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCarga)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Item_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Item_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub Item_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCarga)

End Sub

Public Sub Item_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridCarga.objControle = Item(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCarga)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CodInsumo_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodInsumo_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub CodInsumo_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub CodInsumo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridInsumo.objControle = CodInsumo(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridInsumo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescInsumo_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescInsumo_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub DescInsumo_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub DescInsumo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridInsumo.objControle = DescInsumo(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridInsumo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub KgCargaInsumo_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub KgCargaInsumo_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub KgCargaInsumo_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub KgCargaInsumo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridInsumo.objControle = KgCargaInsumo(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridInsumo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Fornecedor_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Fornecedor_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub Fornecedor_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub Fornecedor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridInsumo.objControle = Fornecedor(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridInsumo)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Lote_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub Lote_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub Lote_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjInfoUsu.objGridInsumo.objControle = Lote(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridInsumo)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'#######################################################################
'FIM DO SCRIPT DOS ITENS DO GRID
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT DE SAÍDA DE CÉLULAS
'#######################################################################
Private Function Saida_Celula_ProdutoBase(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objMaquina As New ClassMaquinas
Dim objKit As New ClassKit
Dim objProdutoFilial As New ClassProdutoFilial
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_ProdutoBase

    Set objGridInt.objControle = ProdutoBase(objCT)

    lErro = CF("Produto_Formata", ProdutoBase(objCT).Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 131098
    
    iLinha = GridCarga(objCT).Row
    
    'se o produto foi preenchido
    If Len(Trim(ProdutoBase(objCT).ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica_Estoque", ProdutoBase(objCT).Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 131099

        If lErro = 25077 Then gError 131100
        
        If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = 0 Then
            
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = iLinha
            If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgProdBase_Col) = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase
            If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdCarga_Col) = StrParaDbl(Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga))
            If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgCarga <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgCarga_Col) = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgCarga)
            If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit <> 0 And objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdTotal_Col) = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit * objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga)
                        
        End If
        
        lErro = Trata_Alteracao_ProdutoBase(objCT, iLinha, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 131135
        
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).sProdutoBase = objProduto.sCodigo

        'Descricao
        GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Descricao_Col) = objProduto.sDescricao
        
    Else
    
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).sProdutoBase = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131102

    Saida_Celula_ProdutoBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoBase:

    Saida_Celula_ProdutoBase = gErr

    Select Case gErr

        Case 131098, 131099, 131102, 131135
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 131100
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objCT.Produto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = objCT.Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    '#########################################
    'Inserido por Wagner - CROMATON 10/11/04
    'Limpa o lixo deixado quando ocorre um erro nessa função
    Call Limpa_Linha_GridCarga(objCT, GridCarga(objCT).Row, False)
    '#########################################

    Exit Function

End Function

Private Function Saida_Celula_KgProdBase(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dFator As Double
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_KgProdBase

    Set objGridInt.objControle = KgProdBase(objCT)
    
    iLinha = GridCarga(objCT).Row
    
    'se a quantidade foi preenchida
    If Len(KgProdBase(objCT).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(KgProdBase(objCT).Text)
        If lErro <> SUCESSO Then gError 131108
       
        KgProdBase(objCT).Text = Formata_Estoque(KgProdBase(objCT).Text)
    
        'Fator de alteração = Valor Atual dividido pelo valor antigo
        dFator = StrParaDbl(KgProdBase(objCT).Text) / objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase
    
        'Atualiza os Kg do Produto Base
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase = StrParaDbl(KgProdBase(objCT).Text)
    
        If dFator <> 1 Then
        
            'Atualiza a Qtd de Cargas
            If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga <> 0 Then objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga / dFator)
        
            'Atualiza as Cargas com base no fator de alteração
            lErro = Atualiza_Carga(objCT, iLinha, dFator)
            If lErro <> SUCESSO Then gError 131141
        End If
        
    Else
    
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131109

    Saida_Celula_KgProdBase = SUCESSO

    Exit Function

Erro_Saida_Celula_KgProdBase:

    Saida_Celula_KgProdBase = gErr

    Select Case gErr

        Case 131108
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 131109, 131141
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QtdCarga(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim dFator As Double
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim iLinha As Integer
Dim iNum As Integer

On Error GoTo Erro_Saida_Celula_QtdCarga

    Set objGridInt.objControle = QtdCarga(objCT)
    
    'se a quantidade foi preenchida
    If Len(QtdCarga(objCT).ClipText) > 0 Then

        lErro = Inteiro_Critica(QtdCarga(objCT).Text)
        If lErro <> SUCESSO Then gError 131114
       
        QtdCarga(objCT).Text = Formata_Estoque(QtdCarga(objCT).Text)
        
        iLinha = GridCarga(objCT).Row
               
        'Se houve alteração na quantidade
        If Abs(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga - StrParaDbl(QtdCarga(objCT).Text)) > QTDE_ESTOQUE_DELTA2 Then
               
            'Pergunta se quer atualizar quantidade produzida
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ATUALIZAR_KGCARGA")
        
            If vbMsg = vbYes Then
                'Atualiza as quantidades
                
                If objCT.gcolItemOP.Count >= GridCarga(objCT).Row And GridCarga(objCT).Row <> 0 Then
                    iNum = objCT.gcolItemOP.Item(GridCarga(objCT).Row)
                Else
                    iNum = 0
                End If
                
                If iNum <> 0 Then gError 131216
                    
                sProduto = objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Produto_Col)
                
                'Critica o formato do Produto
                lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 131201
                
                objProduto.sCodigo = sProdutoFormatado
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO Then gError 131202
                
                '==>> Obtem Fator de conversão para KG
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, "KG", objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 131203
                
                'Atualiza a quantidade do item da OP na unidade de medida informada no grid
                objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Quantidade_Col) = Formata_Estoque(dFator * objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit * StrParaDbl(QtdCarga(objCT).Text))
                
                'Atualiza a quantidate total em KG
                GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdTotal_Col) = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit * StrParaDbl(QtdCarga(objCT).Text))
                
                'Considera que não houve alteração na quantidade para não
                'recalcular os dados com base no Kit
'                objCT.iAlteradoQuantidade = 0
                
            Else
                'Senão atualiza as cargas
                
                If StrParaDbl(GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdCarga_Col)) = 0 Then gError 131269

                'Calcula fator de ajuste das cargas
                dFator = StrParaDbl(GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdCarga_Col)) / StrParaDbl(QtdCarga(objCT).Text)
                                    
            End If
            
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = StrParaDbl(QtdCarga(objCT).Text)
        
        End If
        
    Else
    
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = 0
       
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131115
    
    If vbMsg = vbNo Then
        '=====> Atualiza Carga do Item e atualiza Grids
        lErro = Atualiza_Carga(objCT, iLinha, dFator)
        If lErro <> SUCESSO And lErro <> 131268 Then gError 131204
                
        lErro = Trata_Alteracao_ProdutoBase(objCT, iLinha, objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).sProdutoBase)
        If lErro <> SUCESSO Then gError 131205
    End If

    Saida_Celula_QtdCarga = SUCESSO

    Exit Function

Erro_Saida_Celula_QtdCarga:

    Saida_Celula_QtdCarga = gErr

    Select Case gErr

        Case 131114
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 131115, 131143, 131201, 131202, 131203, 131204, 131205
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 131216
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_QTD_NAO_PERMITIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 131269
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_KGCARGA", gErr, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Fornecedor(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim objItemOPCargaInsumo As New ClassItemOPCargaInsumo
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Fornecedor

    Set objGridInt.objControle = Fornecedor(objCT)
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridCarga(objCT).Row = 0 Then
        iLinha2 = 1
    Else
        iLinha2 = GridCarga(objCT).Row
    End If
   
    iLinha = GridInsumo(objCT).Row
        
    If Len(Trim(Fornecedor(objCT).Text)) > 0 Then
            
        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le3(Fornecedor(objCT), objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 140335
        
        sProduto = GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Insumo_Col)
        
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141761
        
        If iCodFilial = 0 Then iCodFilial = 1
        
        objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo
        objFornecedorProdutoFF.sProduto = sProdutoFormatado
        objFornecedorProdutoFF.iFilialEmpresa = giFilialEmpresa
        objFornecedorProdutoFF.iFilialForn = iCodFilial
        
        'verifica se o produto em questão está associado ao fornecedor
        lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
        If lErro <> SUCESSO And lErro <> 54217 Then gError 141762
            
        'se não hover associação
        If lErro <> SUCESSO Then
        
            'pergunta se deseja criar a associação
            lErro = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_FORNECEDOR_DESEJA_CADASTRAR", objFornecedorProdutoFF.sProduto, objFornecedor.sNomeReduzido)
            
            'Se o usuário respondeu que deseja
            If lErro = vbYes Then
                
                'Chama a tela de produto vs fornecedor
                Call Chama_Tela("FornFilialProduto", objFornecedorProdutoFF)
                
            Else
            
                gError 141763
                
            End If
                
        End If
    
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha2).colInsumos.Item(iLinha).lFornecedor = objFornecedor.lCodigo
            
    Else
    
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha2).colInsumos.Item(iLinha).lFornecedor = 0
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha2).colInsumos.Item(iLinha).sLote = ""
    
        GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Lote_Col) = ""
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131116

    Saida_Celula_Fornecedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Fornecedor:

    Saida_Celula_Fornecedor = gErr

    Select Case gErr
        
        Case 131116, 140335, 141761, 141762, 141763
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Lote(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'faz a critica da celula de lote do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim sProduto As String
Dim objFornecedor As New ClassFornecedor
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastroLote As New ClassRastreamentoLote
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As ClassProduto

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote(objCT)
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridCarga(objCT).Row = 0 Then
        iLinha2 = 1
    Else
        iLinha2 = GridCarga(objCT).Row
    End If
   
    iLinha = GridInsumo(objCT).Row
    
    If Len(Trim(Lote(objCT).Text)) > 0 Then
    
'        objFornecedor.sNomeReduzido = GridInsumo.TextMatrix(GridInsumo.Row, iGrid_Fornecedor_Col)
'
'        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
'        If lErro <> SUCESSO And lErro <> 6681 Then gError 141752
        
        sProduto = GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Insumo_Col)
        
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141753
        
'        lErro = CF("RastroLoteLista_Le", objFornecedor.lCodigo, sProdutoFormatado, Lote.Text)
'        If lErro <> SUCESSO And lErro <> 141751 Then gError 141754
'
'        If lErro <> SUCESSO Then gError 141755

         Set objProduto = New ClassProduto
    
         'le o produto para obter sua descricao
         objProduto.sCodigo = sProdutoFormatado
         lErro = CF("Produto_Le", objProduto)
         If lErro <> SUCESSO Then gError 131180
         
         If objProduto.iRastro = PRODUTO_RASTRO_OP Then objRastroLote.iFilialOP = giFilialEmpresa
        
        objRastroLote.sCodigo = Lote(objCT).Text
        objRastroLote.sProduto = sProdutoFormatado

        'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
        lErro = CF("RastreamentoLote_Le", objRastroLote)
        If lErro <> SUCESSO And lErro <> 75710 Then gError 141754

        'Se não encontrou --> Erro
        If lErro = 75710 Then gError 141755
        
        
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha2).colInsumos.Item(iLinha).sLote = Lote(objCT).Text
        
    Else

        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha2).colInsumos.Item(iLinha).sLote = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141756

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr
        
        Case 141752, 141753, 141754, 141756
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 141755
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORN_LOTE_PROD_NAO_CADASTRADO", gErr, objFornecedor.lCodigo, sProduto, Lote.Text)
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function
'#######################################################################
'FIM DO SCRIPT DE SAÍDA DE CÉLULAS
'#######################################################################

Private Function Preenche_GridCarga(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado  As String
Dim objItemOPCarga As ClassItemOPCarga
Dim objProduto As ClassProduto
Dim dFator As Double
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Preenche_GridCarga

    'Se não houve alteração, não é necessário fazer nada
    If objCT.iAlterado = 0 Then Exit Function
    
    'Se não existe carga, não é necessário fazer nada
    If objCT.gobjInfoUsu.gcolItemOPCarga Is Nothing Then Exit Function
        
    'limpa o Grid de Carga
    Call Limpa_GridCarga(objCT)
    
    'Para cada item da Carga
    For Each objItemOPCarga In objCT.gobjInfoUsu.gcolItemOPCarga
    
        iIndice = iIndice + 1
        
        'Se existe associação da Carga com o Item da OP traz os dados, senão limpa
        If objItemOPCarga.iItem <> 0 Then
           
            '1 - Preenche o Produto, Descrição
            If Len(Trim(objItemOPCarga.sProdutoBase)) > 0 Then
           
                sProdutoMascarado = String(STRING_PRODUTO, 0)
       
                'Mascara produto
                lErro = Mascara_RetornaProdutoTela(objItemOPCarga.sProdutoBase, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 131118
       
                ProdutoBase(objCT).PromptInclude = False
                ProdutoBase(objCT).Text = sProdutoMascarado
                ProdutoBase(objCT).PromptInclude = True
               
                GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col) = ProdutoBase(objCT).Text
           
                Set objProduto = New ClassProduto
           
                'le o produto para obter sua descricao
                objProduto.sCodigo = objItemOPCarga.sProdutoBase
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO Then gError 131180
    
                GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Descricao_Col) = objProduto.sDescricao
           
            End If
           
            '2 - Preenche Item, KgCarga, KgProdBase e QtdCarga
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Item_Col) = iIndice
           
            If Len(Trim(objItemOPCarga.sProdutoBase)) > 0 Then
                If objItemOPCarga.dKgCarga <> 0 Then GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_KgCarga_Col) = Formata_Estoque(objItemOPCarga.dKgCarga)
                If objItemOPCarga.dKgProdBase <> 0 Then GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_KgProdBase_Col) = Formata_Estoque(objItemOPCarga.dKgProdBase)
                If objItemOPCarga.dQtdCarga <> 0 Then GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QtdCarga_Col) = Formata_Estoque(objItemOPCarga.dQtdCarga)
            End If
                   
           '3 - Preenche a quantidade total convertendo para KG
           Set objProduto = New ClassProduto
           
           sProduto = objCT.GridMovimentos.TextMatrix(objItemOPCarga.iItem, objCT.iGrid_Produto_Col)
           
           'Critica o formato do Produto
           lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
           If lErro <> SUCESSO Then gError 31563
           
           objProduto.sCodigo = sProdutoFormatado
           lErro = CF("Produto_Le", objProduto)
           If lErro <> SUCESSO Then gError 131195
           
           'Calcula o fator de conversão para KG
           lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objCT.GridMovimentos.TextMatrix(objItemOPCarga.iItem, objCT.iGrid_UnidadeMed_Col), "KG", dFator)
           If lErro <> SUCESSO Then gError 131196
           
           If Len(Trim(objCT.GridMovimentos.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))) > 0 Then GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QtdTotal_Col) = Formata_Estoque(StrParaDbl(objCT.GridMovimentos.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col)) * dFator)
           
           
        Else
            'Limpa os Dados
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QtdTotal_Col) = ""
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col) = ""
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Item_Col) = iIndice
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_KgCarga_Col) = ""
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_KgProdBase_Col) = ""
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QtdCarga_Col) = ""
            GridCarga(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Descricao_Col) = ""
        End If
           
    Next
    
    objCT.gobjInfoUsu.objGridCarga.iLinhasExistentes = iIndice

    Preenche_GridCarga = SUCESSO

    Exit Function

Erro_Preenche_GridCarga:

    Preenche_GridCarga = gErr

    Select Case gErr
    
        Case 131118, 131180, 131195, 131196

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Private Function Atualiza_gcolItemOPCarga(ByVal objCT As Object, ByVal iLinha As Integer, ByVal sProduto As String, ByVal sVersao As String, ByVal dQuantidade As Double, Optional ByVal bForcar As Boolean = False) As Long
'Mantém consistente a coleção global gcolItemOPCarga, tratando alterações
'do produto (Raiz do Kit), da versão do Kit e da quantidade a ser produzida

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoKit As ClassProdutoKit
Dim objKit As New ClassKit
Dim objItemOPCargaInsumo As New ClassItemOPCargaInsumo
Dim objProduto As New ClassProduto
Dim dCargaTotal As Double
Dim dFator As Double
Dim bAchou As Boolean
Dim vbMsgRes As VbMsgBoxResult
Dim dFatorConversao As Double
Dim sUM As String
Dim iAlteradoVersao As Integer
Dim iAlteradoProduto As Integer
Dim iAlteradoQuantidade As Integer

On Error GoTo Erro_Atualiza_gcolItemOPCarga

    If sVersao <> objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Versao_Col) Then iAlteradoVersao = REGISTRO_ALTERADO
    If bForcar Or (sProduto <> objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Produto_Col)) Then iAlteradoProduto = REGISTRO_ALTERADO
    If dQuantidade <> StrParaDbl(objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Quantidade_Col)) Then iAlteradoQuantidade = REGISTRO_ALTERADO

    'Se não houve alterações na versão, no produto ou na quantidade,
    'não é necessário fazer nada
    If iAlteradoVersao = 0 And iAlteradoProduto = 0 And iAlteradoQuantidade = 0 Then Exit Function

    'Se o produto estiver em branco não será possível calcular nada
    '===> Sai da função
    If Len(Trim(sProduto)) = 0 Then Exit Function
    
    'Se a versão estiver em branco não será possível calcular nada
    '===> Sai da função
    If Len(Trim(sVersao)) = 0 Then Exit Function
    
    'Critica o formato do Produto
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 131119

    Call CarregaCombo_Embalagem(objCT, Embalagem(objCT), sProdutoFormatado)
    
    objKit.sProdutoRaiz = sProdutoFormatado
    objKit.sVersao = sVersao
    
    'Le os componentes do Kit
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 131120
    
    'Produto sem Kit => Não pode ser produzido
    If lErro <> SUCESSO Then gError 131121
    
    'Exclui os Insumos
    If iAlteradoProduto <> 0 Or iAlteradoVersao <> 0 Then
        For iIndice = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Count To 1 Step -1
        
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Remove (iIndice)
        
        Next
    End If
    
    'Flag para achar produto base
    bAchou = False
    
    iIndice = 0
    
    'Para cada componente do Kit
    For Each objProdutoKit In objKit.colComponentes
    
        iIndice = iIndice + 1
       
        'Lê o produto para obter sua descrição
        objProduto.sCodigo = objProdutoKit.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 131124
      
        'Calcula o fator de conversão para Kg
        'Se não puder ser convertido para KG => ERRO
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, "KG", dFatorConversao)
        If lErro <> SUCESSO Then gError 131197
               
        'Se o produto não for o Raiz
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
        
            'Se faz parte da carga
            If objProdutoKit.objInfoUsu.iTipoCarga <> PRODUTOKIT_NAOCARGA Then
            
                'Verifica se é um produto Base Para Carga
                If objProdutoKit.objInfoUsu.iTipoCarga = PRODUTOKIT_BASECARGA Then bAchou = True
                
                'Se trocou o Kit(Produto ou Versao) altera o produto base do Kit
                If iAlteradoProduto <> 0 Or iAlteradoVersao <> 0 Then
                    If objProdutoKit.objInfoUsu.iTipoCarga = PRODUTOKIT_BASECARGA Then
                        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).sProdutoBase = objProdutoKit.sProduto
                        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase = objProdutoKit.dQuantidade * dFatorConversao
                    End If
                
                    Set objItemOPCargaInsumo = New ClassItemOPCargaInsumo
                                           
                    'Todos componentes do Kit exceto o Raiz são insumos e os que não fazem parte da carga
                    objItemOPCargaInsumo.iSeq = iIndice
                    objItemOPCargaInsumo.dKgCarga = objProdutoKit.dQuantidade * dFatorConversao
                    objItemOPCargaInsumo.sProduto = objProdutoKit.sProduto
                    
                    objItemOPCargaInsumo.sDescricao = objProduto.sDescricao
                                   
                    objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Add objItemOPCargaInsumo
            
                End If
            
            End If
        
        Else
        'Se é produto Raiz, verifica quantidade necessária para produzi-lo
        
            If iAlteradoProduto <> 0 Or iAlteradoVersao <> 0 Then
                objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit = objProdutoKit.dQuantidade * dFatorConversao
            End If
                
        End If
        
    Next
    
    'Calcula a carga total do item de carga
    For Each objItemOPCargaInsumo In objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos
    
        dCargaTotal = dCargaTotal + objItemOPCargaInsumo.dKgCarga
    
    Next
    
    'Preenche o Kg por Carga = Carga Total calculada no Loop
    objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgCarga = dCargaTotal
       
    'Se a quantiadde a ser produzida for diferente de zero
    If dQuantidade > 0 Then
    
        'Se só mexeu na quantidade
        If Len(Trim(objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Quantidade_Col))) > 0 And iAlteradoProduto = 0 And iAlteradoVersao = 0 Then
            
            'Quantidade de Cargas = Quantidade de Cargas antigas * (Fator de alteração)
            'Fator de alteração = Quantidade a ser produzida nova / Quantidade a ser produzida antiga
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga * (dQuantidade / StrParaDbl(objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_Quantidade_Col))))
            
        Else
            'Senão recalcula tudo com base no kit
        
            'Converte tudo para KG e
            'KgCarga = 1 Kit
            'Formúla: Quantos Kits são necessários para produzir N produtos Raizes
            'Quantidade de Cargas = Quantidade a ser produzida / Quantidade do produto Raiz no Kit
            sUM = objCT.GridMovimentos.TextMatrix(iLinha, objCT.iGrid_UnidadeMed_Col)
        
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO Then gError 131124
          
            'Calcula fator de conversao para KG
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, "KG", dFatorConversao)
            If lErro <> SUCESSO Then gError 131197
            
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = Formata_Estoque((dQuantidade * dFatorConversao) / (objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit)) 'A quantidade do Kit já foi convertida para KG
    
        End If
    
    Else
    'Se nada vai ser produzido a quantidade de cargas é igual a zero
    
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga = 0
    
    End If
               
    'Se houve alteração no produto
    If iAlteradoProduto = REGISTRO_ALTERADO Or iAlteradoVersao = REGISTRO_ALTERADO Then
        If Not bAchou Then
        
            'Verifica se á para validar o produto base
            If gobjEST.iValidaProdBase = VALIDA_PRODUTO_BASE_CARGA Then
    '            vbMsgRes = Rotina_Aviso(vbokonly, "AVISO_KIT_SEM_PRODUTOBASE", objKit.sProdutoRaiz, objKit.sVersao)
            End If
            
            'Se não tem produto base considera como não preenchido, ou seja,
            'Não vai trazer dados para tela nem gravar até o produto base
            'Ser preenchido manualmente
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = 0
        
        End If
    End If
    
    iAlteradoProduto = 0
    iAlteradoQuantidade = 0
    iAlteradoVersao = 0
    
    Atualiza_gcolItemOPCarga = SUCESSO

    Exit Function

Erro_Atualiza_gcolItemOPCarga:

    Atualiza_gcolItemOPCarga = gErr

    Select Case gErr
    
        Case 131119 To 131121
        
        Case 131197

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Private Function Preenche_GridInsumos(ByVal objCT As Object, iLinha As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado  As String
Dim objItemOPCargaInsumo As ClassItemOPCargaInsumo
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Preenche_GridInsumos
   
    Call Limpa_GridInsumo(objCT)
    
    'Se a linha seleciona da ainda não foi preenchida => Sai da função
    If iLinha > objCT.gobjInfoUsu.gcolItemOPCarga.Count Then Exit Function
    
    'Se não existe linha selecionada => Sai da função
    If iLinha = 0 Then Exit Function
    
    'Se o item de carga foi excluído => Sai da função
    'If gcolItemOPCarga.Item(iLinha).iItem = 0 Then Exit Function
     
    'Para cada insumo da carga
    For Each objItemOPCargaInsumo In objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos
    
        iIndice = iIndice + 1
    
        '1- Preenche a Carga e a descrição
        GridInsumo(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_KgCargaInsumo_Col) = Formata_Estoque(objItemOPCargaInsumo.dKgCarga)
        GridInsumo(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DescInsumo_Col) = objItemOPCargaInsumo.sDescricao
        
        '2- Preenche o produto
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_RetornaProdutoTela(objItemOPCargaInsumo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 131122

        CodInsumo(objCT).PromptInclude = False
        CodInsumo(objCT).Text = sProdutoMascarado
        CodInsumo(objCT).PromptInclude = True
        
        If objItemOPCargaInsumo.lFornecedor <> 0 Then
        
            Set objFornecedor = New ClassFornecedor
            
            objFornecedor.lCodigo = objItemOPCargaInsumo.lFornecedor
            
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 140334
            
            GridInsumo(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
              
        End If
        
        GridInsumo(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Insumo_Col) = sProdutoMascarado
        GridInsumo(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Lote_Col) = objItemOPCargaInsumo.sLote
    
    Next
    
    objCT.gobjInfoUsu.objGridInsumo.iLinhasExistentes = iIndice

    Preenche_GridInsumos = SUCESSO

    Exit Function

Erro_Preenche_GridInsumos:

    Preenche_GridInsumos = gErr

    Select Case gErr
    
        Case 131122, 140334

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Sub CargaDrays_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Maquina2_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Embalagem_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Mistura_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

    TempoMistura(objCT).Clear
    
    Call CarregaCombo_TempoMistura(TempoMistura(objCT), Mistura(objCT).ItemData(Mistura(objCT).ListIndex))

End Sub

Sub Limpa_GridCarga(ByVal objCT As Object)

    Call Grid_Limpa(objCT.gobjInfoUsu.objGridCarga)

End Sub

Sub Limpa_GridInsumo(ByVal objCT As Object)

    Call Grid_Limpa(objCT.gobjInfoUsu.objGridInsumo)

End Sub

Public Sub TempoMistura_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Saida_Celula_Maquina2(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'faz a critica da celula de Equipamento do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_Saida_Celula_Maquina2

    Set objGridInt.objControle = Maquina2(objCT)

    'Se a Máquina foi especificada => Faz a Validacao da Máquina
    If Len(Trim(Maquina2(objCT).Text)) > 0 Then
        
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", Maquina2(objCT), objMaquina)
        If lErro <> SUCESSO Then gError 131091
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131092

    Saida_Celula_Maquina2 = SUCESSO

    Exit Function

Erro_Saida_Celula_Maquina2:

    Saida_Celula_Maquina2 = gErr

    Select Case gErr

        Case 131091 To 131092
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

End Function

Public Sub Maquina2_GotFocus(ByVal objCT As Object)

    objCT.gobjInfoUsu.iMaquina = MAQUINA_DOIS
    Call Grid_Campo_Recebe_Foco(objCT.objGrid)

End Sub

Public Sub Maquina2_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGrid)

End Sub

Public Sub Maquina2_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGrid.objControle = Maquina2(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Atualiza_Carga(ByVal objCT As Object, ByVal iLinha As Integer, ByVal dFator As Double) As Long
'Recalcula a carga de cada insumo e a carga total dos insumos com base em um
'fator de alteração passado e depois preenche os Grids com os Novos Dados

Dim dCargaTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Atualiza_Carga

    If iLinha = 0 Then Exit Function
    
    If dFator = 0 Then gError 131268

    'Para cada Insumo
    For iIndice = 1 To objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Count

        'Carga Atual = Carga Antiga * Fator de Alteração
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Item(iIndice).dKgCarga = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Item(iIndice).dKgCarga * dFator

        'Soma as Cargas
        dCargaTotal = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos.Item(iIndice).dKgCarga + dCargaTotal

    Next

    objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgCarga = dCargaTotal
    objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit * dFator
    
    'Preenche o Grid de Insumos da linha selecionada
    lErro = Preenche_GridCarga(objCT)
    If lErro <> SUCESSO Then gError 131133
    
    'Preenche o Grid de Insumos da linha selecionada
    lErro = Preenche_GridInsumos(objCT, iLinha)
    If lErro <> SUCESSO Then gError 131134

    Atualiza_Carga = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Carga:

    Atualiza_Carga = gErr

    Select Case gErr
    
        Case 131133 To 131134
        
        Case 131268
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_KGCARGA", gErr, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub BotaoProdutoBase_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridCarga(objCT).Row = 0 Then gError 131125

    'Verifica se o Produto está preenchido
    If Len(Trim(GridCarga(objCT).TextMatrix(GridCarga(objCT).Row, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridCarga(objCT).TextMatrix(GridCarga(objCT).Row, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 131126
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto
    
    Set objEventoProdutoBase = New AdmEvento
    Set gobjCT = objCT

    'Lista de produtos
    Call Chama_Tela("ProdutosKitLista", colSelecao, objProduto, objEventoProdutoBase)
        
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 131125
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 131126
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
End Sub

Private Sub objEventoProdutoBase_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objCT As Object

On Error GoTo Erro_objEventoProdutoBase_evSelecao

    Set objProduto = obj1
    Set objCT = gobjCT

    If GridCarga(objCT).Row <> 0 Then

        lErro = CF("Produto_Formata", GridCarga(objCT).TextMatrix(GridCarga(objCT).Row, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 131127

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'mascara produto escolhido
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 131128
       
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 131129

        If lErro = 28030 Then gError 131131

        ProdutoBase(objCT).PromptInclude = False
        ProdutoBase(objCT).Text = sProdutoMascarado
        ProdutoBase(objCT).PromptInclude = True
        
        iLinha = GridCarga(objCT).Row
        
        If Not (objCT.ActiveControl Is ProdutoBase(objCT)) Then
            
            lErro = CF("Produto_Critica_Estoque", ProdutoBase(objCT).Text, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25077 Then gError 131099
    
            If lErro = 25077 Then gError 131100
            
            
            If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = 0 Then
                
                objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = iLinha
                If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgProdBase_Col) = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase
                If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdCarga_Col) = StrParaDbl(Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga))
                If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgCarga <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgCarga_Col) = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgCarga)
                If objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit <> 0 And objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga <> 0 Then GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdTotal_Col) = Formata_Estoque(objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgKit * objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dQtdCarga)
                            
            End If
            
            lErro = Trata_Alteracao_ProdutoBase(objCT, iLinha, objProduto.sCodigo)
            If lErro <> SUCESSO Then gError 131209
            
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).sProdutoBase = objProduto.sCodigo
                
            'Descricao
            GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Descricao_Col) = objProduto.sDescricao
            GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ProdutoBase_Col) = ProdutoBase(objCT).Text
                  
        End If
            
    End If

    objCT.Show

    Exit Sub

Erro_objEventoProdutoBase_evSelecao:

    Select Case gErr

        Case 131127, 131129, 131209

        Case 131128
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
            
        Case 131131
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Function Trata_Alteracao_ProdutoBase(ByVal objCT As Object, ByVal iLinha As Integer, ByVal sNovoProdBase As String) As Long
'Altera o produto base e a carga do produto base

Dim lErro As Long
Dim objItemOPCargaInsumo As ClassItemOPCargaInsumo
Dim iCount As Integer

On Error GoTo Erro_Trata_Alteracao_ProdutoBase

    iCount = 0
    
    'Para cada insumo
    For Each objItemOPCargaInsumo In objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).colInsumos

        'Se o insumo é o produto base
        If objItemOPCargaInsumo.sProduto = sNovoProdBase Then
        
            iCount = iCount + 1

            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).sProdutoBase = objItemOPCargaInsumo.sProduto
            objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).dKgProdBase = objItemOPCargaInsumo.dKgCarga
            
            GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgProdBase_Col) = Formata_Estoque(objItemOPCargaInsumo.dKgCarga)
            
        End If
        
    Next
    
    'Não encontrou o novo produto dentre os insumos
    If iCount = 0 Then gError 131138
    
    'Encontrou mais de um insumo do mesmo produto
    If iCount > 1 Then gError 131139

    Trata_Alteracao_ProdutoBase = SUCESSO

    Exit Function

Erro_Trata_Alteracao_ProdutoBase:

    Trata_Alteracao_ProdutoBase = gErr

    Select Case gErr
    
        Case 131138
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INSUMO", gErr, sNovoProdBase)
        
        Case 131139
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_KIT", gErr, sNovoProdBase)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Function Limpa_Linha_GridCarga(ByVal objCT As Object, ByVal iLinha As Integer, ByVal bInseriuLinha) As Long

On Error GoTo Erro_Limpa_Linha_GridCarga

    If iLinha <= 0 Then Exit Function
    
    objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha).iItem = 0

    GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Descricao_Col) = ""
    GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgCarga_Col) = ""
    GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_KgProdBase_Col) = ""
    GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdCarga_Col) = ""
    GridCarga(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QtdTotal_Col) = ""
    
    Limpa_Linha_GridCarga = SUCESSO
    
    Exit Function
    
Erro_Limpa_Linha_GridCarga:

    Limpa_Linha_GridCarga = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Private Function Valida_gcolItemOPCarga(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemOPCarga As ClassItemOPCarga
Dim objItemOPCargaInsumo As ClassItemOPCargaInsumo
Dim objItemOPCargaInsumoAux As ClassItemOPCargaInsumo
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim bAchouRepetido As Boolean

On Error GoTo Erro_Valida_gcolItemOPCarga

    For Each objItemOPCarga In objCT.gobjInfoUsu.gcolItemOPCarga
    
        iIndice = iIndice + 1
    
        If Abs(StrParaInt(CStr(objItemOPCarga.dQtdCarga)) - objItemOPCarga.dQtdCarga) > QTDE_ESTOQUE_DELTA2 Then gError 131207
    
        iIndice1 = 0
    
        'Para ca Insumo
        For Each objItemOPCargaInsumo In objItemOPCarga.colInsumos
        
            iIndice1 = iIndice1 + 1
        
            iIndice2 = 0
        
            'Para cada Insumo
            For Each objItemOPCargaInsumoAux In objItemOPCarga.colInsumos
            
                iIndice2 = iIndice2 + 1
                
                'Se não for a mesma linha
                If iIndice1 <> iIndice2 Then
                                
                    'Se forem insumos do mesmo produto
                    If objItemOPCargaInsumo.sProduto = objItemOPCargaInsumoAux.sProduto Then
                        'Se o lote não estiver preenchido => Erro
                        If Len(Trim(objItemOPCargaInsumoAux.sLote)) = 0 Then gError 140336
                    
                        If objItemOPCargaInsumo.sLote = objItemOPCargaInsumoAux.sLote Then gError 140336
                    
                    End If
                
                End If
            
            Next
        
        Next
    
    Next

    Valida_gcolItemOPCarga = SUCESSO

    Exit Function

Erro_Valida_gcolItemOPCarga:

    Valida_gcolItemOPCarga = gErr

    Select Case gErr
    
        Case 131207
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDCARGA_NAO_INT", gErr, iIndice)
    
        Case 140336
            Call Rotina_Erro(vbOKOnly, "ERRO_CARGA_PRODUTO_REPETIDO", gErr, objItemOPCargaInsumoAux.sProduto)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Sub Agua_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Oleo_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Amperagem_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VRDe_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VRAte_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Gate_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TCZonaDe_Change(ByVal objCT As Object, Index As Integer)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TCZonaAte_Change(ByVal objCT As Object, Index As Integer)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tela_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TEZonaDe_Change(ByVal objCT As Object, Index As Integer)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TEZonaAte_Change(ByVal objCT As Object, Index As Integer)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VVDe_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VVAte_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Amostras_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Agua_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Agua_Validate

    If Len(Agua(objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(Agua(objCT).Text)
        If lErro <> SUCESSO Then gError 136600
    
    End If

    Exit Sub

Erro_Agua_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136600
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Oleo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Oleo_Validate

    If Len(Oleo(objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(Oleo(objCT).Text)
        If lErro <> SUCESSO Then gError 136601
    
    End If

    Exit Sub

Erro_Oleo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136601
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Amperagem_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Amperagem_Validate

    If Len(Amperagem(objCT).ClipText) > 0 Then
    
        lErro = Valor_Inteiro_Critica(Amperagem(objCT).Text)
        If lErro <> SUCESSO Then gError 136602
    
    End If

    Exit Sub

Erro_Amperagem_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136602
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub VRDe_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VRDe_Validate

    If Len(VRDe(objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(VRDe(objCT).Text)
        If lErro <> SUCESSO Then gError 136603
    
    End If

    Exit Sub

Erro_VRDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136603
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub VRAte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VRAte_Validate

    If Len(VRAte(objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(VRAte(objCT).Text)
        If lErro <> SUCESSO Then gError 136604
    
    End If

    Exit Sub

Erro_VRAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136604
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Gate_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Gate_Validate

    If Len(Gate(objCT).ClipText) > 0 Then
    
        lErro = Porcentagem_Critica(Gate(objCT).Text)
        If lErro <> SUCESSO Then gError 136605
    
    End If

    Exit Sub

Erro_Gate_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136605
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub TCZonaDe_Validate(ByVal objCT As Object, iIndice As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TCZonaDe_Validate

    If Len(TCZonaDe(iIndice, objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(TCZonaDe(iIndice, objCT).Text)
        If lErro <> SUCESSO Then gError 136606
    
    End If

    Exit Sub

Erro_TCZonaDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 136606

        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub TCZonaAte_Validate(ByVal objCT As Object, iIndice As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TCZonaAte_Validate

    If Len(TCZonaAte(iIndice, objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(TCZonaAte(iIndice, objCT).Text)
        If lErro <> SUCESSO Then gError 136607
    
    End If

    Exit Sub

Erro_TCZonaAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136607
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub TEZonaDe_Validate(ByVal objCT As Object, iIndice As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TEZonaDe_Validate

    If Len(TEZonaDe(iIndice, objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(TEZonaDe(iIndice, objCT).Text)
        If lErro <> SUCESSO Then gError 136608
    
    End If

    Exit Sub

Erro_TEZonaDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136608
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub TEZonaAte_Validate(ByVal objCT As Object, iIndice As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TEZonaAte_Validate

    If Len(TEZonaAte(iIndice, objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(TEZonaAte(iIndice, objCT).Text)
        If lErro <> SUCESSO Then gError 136609
    
    End If

    Exit Sub

Erro_TEZonaAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136609
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub VVDe_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VVDe_Validate

    If Len(VVDe(objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(VVDe(objCT).Text)
        If lErro <> SUCESSO Then gError 136610
    
    End If

    Exit Sub

Erro_VVDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 136610
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub VVAte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VVAte_Validate

    If Len(VVAte(objCT).ClipText) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(VVAte(objCT).Text)
        If lErro <> SUCESSO Then gError 136611
    
    End If

    Exit Sub

Erro_VVAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 136611
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$)

    End Select

    Exit Sub

End Sub

'#########################################

'################################################################
'Inserido por Wagner 22/09/05
Public Sub BotaoFornecedor_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoFornecedor_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridInsumo(objCT).Row = 0 Then gError 140333
        
    lErro = CF("Produto_Formata", GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Insumo_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141760
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        colSelecao.Add sProdutoFormatado
    
        objFornecedor.sNomeReduzido = GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Fornecedor_Col)
    
        Set objEventoFornecedor = New AdmEvento
    
        Set gobjCT = objCT
    
        'Lista de produtos
        Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor, " Codigo IN (SELECT Fornecedor FROM ProdutoFilialForn WHERE Produto = ?) ")
        
    End If
        
    Exit Sub

Erro_BotaoFornecedor_Click:

    Select Case gErr
    
        Case 140333
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 141760
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim iLinha As Integer
Dim objFornecedor As ClassFornecedor
Dim objCT As Object

On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1
    
    Set objCT = gobjCT

    If GridCarga(objCT).Row <> 0 Then

        Fornecedor(objCT).Text = objFornecedor.sNomeReduzido
            
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(GridCarga(objCT).Row).colInsumos.Item(GridInsumo(objCT).Row).lFornecedor = objFornecedor.lCodigo
            
        If Not (objCT.ActiveControl Is Fornecedor(objCT)) Then
        
            GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Fornecedor_Col) = objFornecedor.sNomeReduzido
        
        End If
            
            
    End If

    objCT.Show

    Exit Sub

Erro_objEventoFornecedor_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub
'##########################################################################

'################################################################
'Inserido por Wagner 10/02/2006
Public Sub BotaoLote_Click(ByVal objCT As Object)

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sLote As String
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoLote_Click

    If (GridInsumo(objCT).Row = 0) Then gError 141743

    sCodProduto = GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Insumo_Col)
    sLote = GridInsumo(objCT).TextMatrix(GridInsumo(objCT).Row, objCT.gobjInfoUsu.iGrid_Lote_Col)
    
    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141744

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 141745
    
'    If Len(Trim(GridInsumo.TextMatrix(GridInsumo.Row, iGrid_Fornecedor_Col))) = 0 Then gError 141746
'
'    objFornecedor.sNomeReduzido = GridInsumo.TextMatrix(GridInsumo.Row, iGrid_Fornecedor_Col)
'
'    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
'    If lErro <> SUCESSO And lErro <> 6681 Then gError 141747
    
    If Len(Trim(sLote)) > 0 Then
        objRastroLoteSaldo.sLote = sLote
    End If

    colSelecao.Add sProdutoFormatado
'    colSelecao.Add objFornecedor.lCodigo

    Set objEventoLote = New AdmEvento
    
    Set gobjCT = objCT

'    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoLote, "Produto = ? AND Fornecedor = ?")
    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoLote, "Produto = ? ")
    
    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr
        
        Case 141743
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 141744, 141747
                    
        Case 141745
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 141746
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim iLinha2 As Integer
Dim iLinha As Integer
Dim objCT As Object

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLoteSaldo = obj1
    
    Set objCT = gobjCT

    'Verifica se tem alguma linha selecionada no Grid
    If GridCarga(objCT).Row = 0 Then
        iLinha2 = 1
    Else
        iLinha2 = GridCarga(objCT).Row
    End If
   
    iLinha = GridInsumo(objCT).Row

    If (GridInsumo(objCT).Row > 0) Then
        Lote(objCT).Text = objRastroLoteSaldo.sLote
        GridInsumo(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Lote_Col) = objRastroLoteSaldo.sLote
        objCT.gobjInfoUsu.gcolItemOPCarga.Item(iLinha2).colInsumos.Item(iLinha).sLote = Lote(objCT).Text
    End If
       
    objCT.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
'######################################################################

Public Function OP_Item_IncluirGrid(ByVal objCT As Object) As Long

Dim objItemOPCarga As New ClassItemOPCarga
    
    '########################################
    'Inserido por Wagner - CROMATON 04/11/04
    objItemOPCarga.iItem = objCT.GridMovimentos.Row
    objCT.gobjInfoUsu.gcolItemOPCarga.Add objItemOPCarga
    '########################################

    OP_Item_IncluirGrid = SUCESSO
        
End Function

Public Function OP_Item_PreencheGridMov(ByVal objCT As Object, ByVal objItemOP As ClassItemOP, ByVal iIndice As Integer) As Long

Dim objMaquina As New ClassMaquinas
Dim lErro As Long

On Error GoTo Erro_OP_Item_PreencheGridMov

    Call CarregaCombo_Embalagem(objCT, Embalagem(objCT), objItemOP.sProduto)

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Tratamento para o segundo equipamento
    If objItemOP.lNumIntEquipamento2 <> 0 Then
    
        objMaquina.lNumIntDoc = objItemOP.lNumIntEquipamento2
        
        'Le a Máquina atraves do NumIntDoc
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 131085
        
        'Se nao encontrou => Erro
        If lErro = 106353 Then gError 131086
        
        'Preenche a máquina
        objCT.GridMovimentos.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Maquina2_Col) = objMaquina.sNomeReduzido
    
    End If
    '#########################################
    
    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    Set objItemOP.objInfoUsu = New ClassItemOPCarga
    
    objItemOP.objInfoUsu.lNumIntItemOP = objItemOP.lNumIntDoc
    
    'Tenta ler o Item de Carga do BD
    lErro = CF("ItemOPCarga_Le", objItemOP.objInfoUsu)
    If lErro <> SUCESSO And lErro <> 131199 Then gError 131084
    
    objCT.gobjInfoUsu.gcolItemOPCarga.Add objItemOP.objInfoUsu
    
    'Se não encontrou, atualiza com base no Kit
    If lErro = 131199 Then
        objItemOP.objInfoUsu.iItem = 0
        lErro = Atualiza_gcolItemOPCarga(objCT, iIndice, objCT.Produto.Text, objItemOP.sVersao, objItemOP.dQuantidade)
        If lErro <> SUCESSO Then gError 131200
        
    Else
        objItemOP.objInfoUsu.iItem = iIndice
    End If
    '#########################################

    OP_Item_PreencheGridMov = SUCESSO
    
Erro_OP_Item_PreencheGridMov:

    Exit Function

Erro_objEventoLote_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function OP_PreencheGridMov(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_OP_PreencheGridMov

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Preenche o Grid de Carga com base na coleção global gcolItemOPCarga
    lErro = Preenche_GridCarga(objCT)
    If lErro <> SUCESSO Then gError 131181
    '#########################################

    OP_PreencheGridMov = SUCESSO
    
    Exit Function
    
Erro_OP_PreencheGridMov:

    OP_PreencheGridMov = gErr
    
    Select Case gErr
    
        Case 131181
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function OP_GridMov_AtualizaLinha(ByVal objCT As Object, Optional ByVal bForcar As Boolean = False) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim dQtd As Double

On Error GoTo Erro_OP_GridMov_AtualizaLinha

    If objCT.ActiveControl Is objCT.Produto Then
        sProduto = objCT.Produto.Text
    Else
        sProduto = objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Produto_Col)
    End If
    
    If objCT.ActiveControl Is objCT.Quantidade Then
        dQtd = StrParaDbl(objCT.Quantidade.Text)
    Else
        dQtd = StrParaDbl(objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Quantidade_Col))
    End If
    
    If objCT.ActiveControl Is objCT.Versao Then
        sVersao = objCT.Versao.Text
    Else
        sVersao = objCT.GridMovimentos.TextMatrix(objCT.GridMovimentos.Row, objCT.iGrid_Versao_Col)
    End If

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    lErro = Atualiza_gcolItemOPCarga(objCT, objCT.GridMovimentos.Row, sProduto, sVersao, dQtd, bForcar)
    If lErro <> SUCESSO Then gError 131084
    '#########################################

    OP_GridMov_AtualizaLinha = SUCESSO
    
    Exit Function
    
Erro_OP_GridMov_AtualizaLinha:

    OP_GridMov_AtualizaLinha = gErr
    
    Select Case gErr
    
        Case 131084
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Function OP_Limpa_Linha_GridMov(ByVal objCT As Object, ByVal iLinha As Integer, ByVal bInseriuLinha As Boolean) As Long

On Error GoTo Erro_OP_Limpa_Linha_GridMov

    If bInseriuLinha Then
        objCT.gobjInfoUsu.gcolItemOPCarga.Remove (iLinha)
    End If

    objCT.GridMovimentos.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Maquina2_Col) = ""

    OP_Limpa_Linha_GridMov = SUCESSO
    
    Exit Function
    
Erro_OP_Limpa_Linha_GridMov:

    OP_Limpa_Linha_GridMov = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function OP_GravarRegistro(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_OP_GravarRegistro

    '#########################################
    'Inserido por Wagner - CROMATON 10/11/04
    lErro = Valida_gcolItemOPCarga(objCT)
    If lErro <> SUCESSO Then gError 131206
    '#########################################

    OP_GravarRegistro = SUCESSO
    
    Exit Function
    
Erro_OP_GravarRegistro:

    OP_GravarRegistro = gErr

    Select Case gErr
    
        Case 131206
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function OP_Move_Grid_Memoria(ByVal objCT As Object, ByVal iIndice As Integer, ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_OP_Move_Grid_Memoria

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Tratamento para o segundo Equipamento
    If Len(Trim(objCT.GridMovimentos.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Maquina2_Col))) > 0 Then
        
        'Le a Maquina a partir do nome reduzido
        objMaquina.sNomeReduzido = objCT.GridMovimentos.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Maquina2_Col)
        
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 131087
        
        'Se nao encontrou => Erro
        If lErro = 103100 Then gError 131088
        
        objItemOP.lNumIntEquipamento2 = objMaquina.lNumIntDoc
        
    End If
    
    'Insere o Item de Carga no objInfoUsu do Item da OP
    Set objItemOP.objInfoUsu = objCT.gobjInfoUsu.gcolItemOPCarga.Item(iIndice)
    '#########################################
    
    OP_Move_Grid_Memoria = SUCESSO
    
    Exit Function
    
Erro_OP_Move_Grid_Memoria:

    OP_Move_Grid_Memoria = gErr

    Select Case gErr
    
        Case 131087
        
        Case 131088
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function OP_Limpa_Tela_OP(ByVal objCT As Object) As Long
    
    '#########################################
    'Inserido por Wagner - CROMATON 04/11/04
    Mistura(objCT).Text = ""
    TempoMistura(objCT).Clear
    CargaDrays(objCT).Text = ""
    
    Embalagem(objCT).Clear
    objCT.gobjInfoUsu.sProdComboEmbalagens = ""
    
    Call Limpa_GridCarga(objCT)
    Call Limpa_GridInsumo(objCT)
    
    Set objCT.gobjInfoUsu.gcolItemOPCarga = New Collection

    '#########################################

End Function

Public Function OP_Move_Tela_Memoria(ByVal objCT As Object, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim iIndice As Integer

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    Set objOrdemDeProducao.objInfoUsu = New ClassOPFichaProcesso
    
    If CargaDrays(objCT).ListIndex <> -1 Then objOrdemDeProducao.objInfoUsu.iCargaDrays = CargaDrays(objCT).ItemData(CargaDrays(objCT).ListIndex)
    If Embalagem(objCT).ListIndex <> -1 Then objOrdemDeProducao.objInfoUsu.iEmbalagem = Embalagem(objCT).ItemData(Embalagem(objCT).ListIndex)
    objOrdemDeProducao.objInfoUsu.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
    If Mistura(objCT).ListIndex <> -1 Then objOrdemDeProducao.objInfoUsu.iMistura = Mistura(objCT).ItemData(Mistura(objCT).ListIndex)
    If TempoMistura(objCT).ListIndex <> -1 Then objOrdemDeProducao.objInfoUsu.iTempoMistura = TempoMistura(objCT).ItemData(TempoMistura(objCT).ListIndex)
    objOrdemDeProducao.objInfoUsu.sCodigoOP = objOrdemDeProducao.sCodigo
    
    objOrdemDeProducao.objInfoUsu.dProcessoAgua = StrParaDbl(Agua(objCT).Text)
    objOrdemDeProducao.objInfoUsu.dProcessoOleo = StrParaDbl(Oleo(objCT).Text)
    objOrdemDeProducao.objInfoUsu.iAmperagem = StrParaInt(Amperagem(objCT).Text)
    objOrdemDeProducao.objInfoUsu.dVelRotoresDe = StrParaDbl(VRDe(objCT).Text)
    objOrdemDeProducao.objInfoUsu.dVelRotoresAte = StrParaDbl(VRAte(objCT).Text)
    objOrdemDeProducao.objInfoUsu.dAberturaGate = StrParaDbl(Gate(objCT).Text)
    
    For iIndice = 1 To FICHAPROCESSO_NUM_ZONAS_TC
        objOrdemDeProducao.objInfoUsu.dTempCamaraZDe(iIndice) = StrParaDbl(TCZonaDe(iIndice - 1, objCT).Text)
        objOrdemDeProducao.objInfoUsu.dTempCamaraZAte(iIndice) = StrParaDbl(TCZonaAte(iIndice - 1, objCT).Text)
    Next
    
    objOrdemDeProducao.objInfoUsu.sTela = Tela(objCT).Text
    
    For iIndice = 1 To FICHAPROCESSO_NUM_ZONAS_TE
        objOrdemDeProducao.objInfoUsu.dTempExtrusoraZDe(iIndice) = StrParaDbl(TEZonaDe(iIndice - 1, objCT).Text)
        objOrdemDeProducao.objInfoUsu.dTempExtrusoraZAte(iIndice) = StrParaDbl(TEZonaAte(iIndice - 1, objCT).Text)
    Next
    
    objOrdemDeProducao.objInfoUsu.dVelVariadorDe = StrParaDbl(VVDe(objCT).Text)
    objOrdemDeProducao.objInfoUsu.dVelVariadorAte = StrParaDbl(VVAte(objCT).Text)
    objOrdemDeProducao.objInfoUsu.sObservacao = Observacao(objCT).Text
    objOrdemDeProducao.objInfoUsu.sAmostra = Amostras(objCT).Text
    '#########################################

End Function

Public Function OP_Traz_Tela_OP(ByVal objCT As Object, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OP_Traz_Tela_OP

    '#########################################
    'Inserido por Wagner - CROMATON 03/11/04
    'Tenta ler a Ficha de Processo da OP
    lErro = CF("OPFichaProcesso_Le", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 131194 Then gError 131089
    
    'Se encontrou => atualiza tela
    If lErro = SUCESSO Then
    
        If objOrdemDeProducao.objInfoUsu.iCargaDrays <> 0 Then
            CargaDrays(objCT).Text = objOrdemDeProducao.objInfoUsu.iCargaDrays
            Call CargaDrays_Validate(objCT, bSGECancelDummy)
        End If
        
        If objOrdemDeProducao.objInfoUsu.iEmbalagem <> 0 Then
            Embalagem(objCT).Text = objOrdemDeProducao.objInfoUsu.iEmbalagem
            Call Embalagem_Validate(objCT, bSGECancelDummy)
        End If
        
        If objOrdemDeProducao.objInfoUsu.iMistura <> 0 Then
            Mistura(objCT).Text = objOrdemDeProducao.objInfoUsu.iMistura
            Call Mistura_Validate(objCT, bSGECancelDummy)
        End If
           
        If objOrdemDeProducao.objInfoUsu.iTempoMistura <> 0 Then
            TempoMistura(objCT).Text = objOrdemDeProducao.objInfoUsu.iTempoMistura
            Call TempoMistura_Validate(objCT, bSGECancelDummy)
        End If
   
        If objOrdemDeProducao.objInfoUsu.dProcessoAgua <> 0 Then
            Agua(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dProcessoAgua, Agua(objCT).Format)
        End If
        
        If objOrdemDeProducao.objInfoUsu.dProcessoOleo <> 0 Then
            Oleo(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dProcessoOleo, Oleo(objCT).Format)
        End If
        
        If objOrdemDeProducao.objInfoUsu.iAmperagem <> 0 Then
            Amperagem(objCT).PromptInclude = False
            Amperagem(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.iAmperagem, Amperagem(objCT).Format)
            Amperagem(objCT).PromptInclude = True
        End If
        
        If objOrdemDeProducao.objInfoUsu.dVelRotoresDe <> 0 Then
            VRDe(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dVelRotoresDe, VRDe(objCT).Format)
        End If
        
        If objOrdemDeProducao.objInfoUsu.dVelRotoresAte <> 0 Then
            VRAte(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dVelRotoresAte, VRAte(objCT).Format)
        End If
        
        If objOrdemDeProducao.objInfoUsu.dAberturaGate <> SUCESSO Then
            Gate(objCT).PromptInclude = False
            Gate(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dAberturaGate, Gate(objCT).Format)
            Gate(objCT).PromptInclude = True
        End If
        
        For iIndice = 1 To FICHAPROCESSO_NUM_ZONAS_TC
        
            If objOrdemDeProducao.objInfoUsu.dTempCamaraZDe(iIndice) <> 0 Then
                TCZonaDe(iIndice - 1, objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dTempCamaraZDe(iIndice), TCZonaDe(iIndice - 1, objCT).Format)
            End If
            
            If objOrdemDeProducao.objInfoUsu.dTempCamaraZAte(iIndice) <> 0 Then
                TCZonaAte(iIndice - 1, objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dTempCamaraZAte(iIndice), TCZonaAte(iIndice - 1, objCT).Format)
            End If
        
        Next
        
        Tela(objCT).Text = objOrdemDeProducao.objInfoUsu.sTela
        
        For iIndice = 1 To FICHAPROCESSO_NUM_ZONAS_TE
        
            If objOrdemDeProducao.objInfoUsu.dTempExtrusoraZDe(iIndice) <> 0 Then
                TEZonaDe(iIndice - 1, objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dTempExtrusoraZDe(iIndice), TEZonaDe(iIndice - 1, objCT).Format)
            End If
            
            If objOrdemDeProducao.objInfoUsu.dTempExtrusoraZAte(iIndice) <> 0 Then
                TEZonaAte(iIndice - 1, objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dTempExtrusoraZAte(iIndice), TEZonaAte(iIndice - 1, objCT).Format)
            End If
        
        Next
        
        If objOrdemDeProducao.objInfoUsu.dVelVariadorDe <> 0 Then
            VVDe(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dVelVariadorDe, VVDe(objCT).Format)
        End If
        
        If objOrdemDeProducao.objInfoUsu.dVelVariadorAte <> 0 Then
            VVAte(objCT).Text = Format(objOrdemDeProducao.objInfoUsu.dVelVariadorAte, VVAte(objCT).Format)
        End If
        
        Observacao(objCT).Text = objOrdemDeProducao.objInfoUsu.sObservacao
        Amostras(objCT).Text = objOrdemDeProducao.objInfoUsu.sAmostra
    
    End If
    '#########################################
    
    OP_Traz_Tela_OP = SUCESSO
    
    Exit Function
    
Erro_OP_Traz_Tela_OP:

    OP_Traz_Tela_OP = gErr

    Select Case gErr
    
        Case 131089
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

