VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBaixaPagAlac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function BAIXAPAG_Inicializa_Grid_Parcelas(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = objCT

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Valor Baixar")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Portador")
    objGridInt.colColuna.Add ("Valor a Pagar")
    objGridInt.colColuna.Add ("Valor Parcela")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Filial")
            
    'campos de edição do grid
    objGridInt.colCampo.Add (objCT.Selecionada.Name)
    objGridInt.colCampo.Add (objCT.DataVencimento.Name)
    objGridInt.colCampo.Add (objCT.Saldo.Name)
    objGridInt.colCampo.Add (objCT.ValorBaixado.Name)
    objGridInt.colCampo.Add (objCT.FornItem.Name)
    objGridInt.colCampo.Add (objCT.ValorDesconto.Name)
    objGridInt.colCampo.Add (objCT.ValorMulta.Name)
    objGridInt.colCampo.Add (objCT.ValorJuros.Name)
    objGridInt.colCampo.Add (objCT.DataEmissaoTitulo.Name)
    objGridInt.colCampo.Add (objCT.Numero.Name)
    objGridInt.colCampo.Add (objCT.Parcela.Name)
    objGridInt.colCampo.Add (objCT.Tipo.Name)
    objGridInt.colCampo.Add (objCT.Cobranca.Name)
    objGridInt.colCampo.Add (objCT.NomePortador.Name)
    objGridInt.colCampo.Add (objCT.ValorAPagar.Name)
    objGridInt.colCampo.Add (objCT.ValorParcela.Name)
    objGridInt.colCampo.Add (objCT.FilialEmpresa.Name)
    objGridInt.colCampo.Add (objCT.FilialFornItem.Name)

    objCT.iGrid_Selecionar_Col = 1
    objCT.iGrid_DataVencimento_Col = 2
    objCT.iGrid_Saldo_Col = 3
    objCT.iGrid_ValorBaixar_Col = 4
    objCT.iGrid_Fornecedor_Col = 5
    objCT.iGrid_Desconto_Col = 6
    objCT.iGrid_Multa_Col = 7
    objCT.iGrid_Juros_Col = 8
    objCT.iGrid_DataEmissaoTitulo_Col = 9
    objCT.iGrid_Numero_Col = 10
    objCT.iGrid_Parcela_Col = 11
    objCT.iGrid_Tipo_Col = 12
    objCT.iGrid_Cobranca_Col = 13
    objCT.iGrid_Portador_Col = 14
    objCT.iGrid_ValorPagar_Col = 15
    objCT.iGrid_ValorParcela_Col = 16
    objCT.iGrid_FilialEmpresa_Col = 17
    objCT.iGrid_Filial_Col = 18

    objGridInt.objGrid = objCT.GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6

    'largura da primeira coluna
    objCT.GridParcelas.ColWidth(0) = 600

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    BAIXAPAG_Inicializa_Grid_Parcelas = SUCESSO
        
    Exit Function

End Function
