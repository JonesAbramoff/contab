VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalRemArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get DescricaoItem2(ByVal objCT As Object) As Object
     Set DescricaoItem2 = objCT.objUserControl.Controls("DescricaoItem2")
End Property

Public Property Get DescricaoItem3(ByVal objCT As Object) As Object
     Set DescricaoItem3 = objCT.objUserControl.Controls("DescricaoItem3")
End Property

Private Sub GridItens_InicializaColunas(ByVal objCT As Object)
    
    With objCT
    
        'Colunas da Grid
        .iGrid_Produto_Col = 1
        .iGrid_DescProduto_Col = 2
        .gobjInfoUsu.iGrid_DescProduto2_Col = 3
        .gobjInfoUsu.iGrid_DescProduto3_Col = 4
        .iGrid_UnidadeMed_Col = 5
        .iGrid_Quantidade_Col = 6
    ''''    iGrid_Almoxarifado_Col = 5
        .iGrid_ValorUnitario_Col = 7
        .iGrid_PercDesc_Col = 8
        .iGrid_Desconto_Col = 9
        .iGrid_PrecoTotal_Col = 10
        .iGrid_PrecoTotalB_Col = 11
        
    End With

End Sub

Public Function NFR_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Cor")
    objGridInt.colColuna.Add ("Detalhe")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
''''    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
    objGridInt.colCampo.Add (DescricaoItem2(objCT).Name)
    objGridInt.colCampo.Add (DescricaoItem3(objCT).Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
''''    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)

    Call GridItens_InicializaColunas(objCT)
    
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

    NFR_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function


