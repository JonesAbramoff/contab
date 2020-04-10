VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFaturaArt"
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

Public Function NFF_Form_Load2(ByVal objct As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFF_Form_Load2

'    Call CF2(objct, "NFCF_Form_Load_PrecoDesc", "Frame1", 62)
    
    NFF_Form_Load2 = SUCESSO
    
    Exit Function
    
Erro_NFF_Form_Load2:

    NFF_Form_Load2 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184276)

    End Select
    
    Exit Function

End Function

Private Function NFF_GridItens_InicializaColunas(ByVal objct As Object) As Long

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
    
        '#########################
        'Inserido por Wagner
        .iGrid_Contrato_Col = 13
        .iGrid_ItemContrato_Col = 14
        '#########################
        .iGrid_DataCobranca_Col = 15
        
        .iGrid_SerieNFOrig_Col = 16
        .iGrid_NumNFOrig_Col = 17
        .iGrid_ItemNFOrig_Col = 18

    End With
    
End Function

Public Function NFF_Inicializa_Grid_Itens(ByVal objct As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens
    
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

    '##############################################
    'Inserido por Wagner
    objGridInt.colColuna.Add ("Contrato")
    objGridInt.colColuna.Add ("Item")
    '##############################################

    objGridInt.colColuna.Add ("Data Cobrança")

    objGridInt.colColuna.Add ("Série NF Orig")
    objGridInt.colColuna.Add ("Núm NF Orig")
    objGridInt.colColuna.Add ("Item NF Orig")

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

    '##############################################
    'Inserido por Wagner
    objGridInt.colCampo.Add (objct.Contrato.Name)
    objGridInt.colCampo.Add (objct.Item.Name)
    '##############################################

    objGridInt.colCampo.Add (objct.DataCobranca.Name)
    
    objGridInt.colCampo.Add (objct.SerieNFOriginal.Name)
    objGridInt.colCampo.Add (objct.NumNFOriginal.Name)
    objGridInt.colCampo.Add (objct.ItemNFOriginal.Name)

    Call NFF_GridItens_InicializaColunas(objct)

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

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFF_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

