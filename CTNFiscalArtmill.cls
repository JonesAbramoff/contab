VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get DescricaoItem2(ByVal objct As Object) As Object
     Set DescricaoItem2 = objct.objUserControl.Controls("DescricaoItem2")
End Property

Public Property Get DescricaoItem3(ByVal objct As Object) As Object
     Set DescricaoItem3 = objct.objUserControl.Controls("DescricaoItem3")
End Property

'precodesc
Public Property Get PrecoDesc(ByVal objct As Object) As Object
     Set PrecoDesc = objct.objUserControl.Controls("PrecoDesc")
End Property

Public Function NF_Form_Load(ByVal objct As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NF_Form_Load

'    Call CF2(objct, "NFCF_Form_Load_PrecoDesc", "Frame2", 3)
    
    NF_Form_Load = SUCESSO
    
    Exit Function
    
Erro_NF_Form_Load:

    NF_Form_Load = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184276)

    End Select
    
    Exit Function

End Function

'ver inicializacao de grid de itens
Private Sub GridItens_InicializaColunas(ByVal objct As Object)

    With objct
    
        'Colunas da Grid
        .iGrid_Produto_Col = 1
        .iGrid_DescProduto_Col = 2
        .gobjInfoUsu.iGrid_DescProduto2_Col = 3
        .gobjInfoUsu.iGrid_DescProduto3_Col = 4
        .iGrid_UnidadeMed_Col = 5
        .iGrid_Quantidade_Col = 6
        .iGrid_ValorUnitario_Col = 7
        'precodesc
        .gobjInfoUsu.iGrid_PrecoDesc_Col = 8
        .iGrid_PercDesc_Col = 9
        .iGrid_Desconto_Col = 10
        .iGrid_PrecoTotal_Col = 11
        .iGrid_PrecoTotalB_Col = 12
        
    End With

End Sub

Public Function NF_Inicializa_Grid_Itens(ByVal objct As Object, ByVal objGridInt As AdmGrid) As Long

    Set objGridInt.objForm = objct

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Cor")
    objGridInt.colColuna.Add ("Detalhe")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    objGridInt.colColuna.Add ("Preço Desc.")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objct.Produto.Name)
    objGridInt.colCampo.Add (objct.DescricaoItem.Name)
    objGridInt.colCampo.Add (DescricaoItem2(objct).Name)
    objGridInt.colCampo.Add (DescricaoItem3(objct).Name)
    objGridInt.colCampo.Add (objct.UnidadeMed.Name)
    objGridInt.colCampo.Add (objct.Quantidade.Name)
    objGridInt.colCampo.Add (objct.PrecoUnitario.Name)
    'precodesc
    objGridInt.colCampo.Add (PrecoDesc(objct).Name)
    objGridInt.colCampo.Add (objct.PercentDesc.Name)
    objGridInt.colCampo.Add (objct.Desconto.Name)
    objGridInt.colCampo.Add (objct.PrecoTotal.Name)
    objGridInt.colCampo.Add (objct.PrecoTotalB.Name)

    Call GridItens_InicializaColunas(objct)
    
    objct.TabPrecoItemNF.left = POSICAO_FORA_TELA
    objct.ComissaoItemNF.left = POSICAO_FORA_TELA
    
    'Grid do GridInterno
    objGridInt.objGrid = objct.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    objct.GridItens.ColWidth(0) = 500

    'Largura automática para o grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NF_Inicializa_Grid_Itens = SUCESSO
    
End Function

Public Function NF_PrecoTotal_Calcula(ByVal objct As Object, ByVal iLinha As Integer, ByVal dQuantidade As Double, ByVal dPrecoUnitario As Double, ByVal dPercentDesc As Double) As Long
    
    'precodesc
    objct.GridItens.TextMatrix(iLinha, objct.gobjInfoUsu.iGrid_PrecoDesc_Col) = Format(dPrecoUnitario * (1 - dPercentDesc), "Standard")
    
    NF_PrecoTotal_Calcula = SUCESSO
    
End Function

Public Function NF_Produto_Saida_Celula(ByVal objct As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_NF_Produto_Saida_Celula

    Call CF2(objct, "Artmill_PreencheCorDetalhe", objProduto.sCodigo, objct.GridItens.Row)

    If Len(Trim(objct.Cliente.Text)) > 0 Then

        'Alterado por Wagner 23/11/04
        lErro = CF("Artmill_PreencheDescontoICMS", objct.Cliente.Text, Codigo_Extrai(objct.Filial.Text), dPercentDesc, objct.gobjInfoUsu.gdDesconto1)
        If lErro <> SUCESSO Then gError 126426

        If dPercentDesc <> 0 Then objct.GridItens.TextMatrix(objct.GridItens.Row, objct.iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

    End If

    NF_Produto_Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_NF_Produto_Saida_Celula:

    NF_Produto_Saida_Celula = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184278)

    End Select
    
    Exit Function

End Function

Public Function NF_Preenche_GridItens(ByVal objct As Object, ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objProduto As ClassProduto, ByVal iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_NF_Preenche_GridItens

    Call CF2(objct, "Artmill_PreencheCorDetalhe", objItemNF.sProduto, iIndice)
    
    If objct.gobjInfoUsu.iGrid_PrecoDesc_Col <> 0 Then
    
        'precodesc
        objct.GridItens.TextMatrix(iIndice, objct.gobjInfoUsu.iGrid_PrecoDesc_Col) = Format(objItemNF.dPrecoUnitario * (1 - objItemNF.dPercDesc), gobjFAT.sFormatoPrecoUnitario)
        
    End If
    
    NF_Preenche_GridItens = SUCESSO
    
    Exit Function
    
Erro_NF_Preenche_GridItens:

    NF_Preenche_GridItens = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184280)

    End Select
    
    Exit Function

End Function

