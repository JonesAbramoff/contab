VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrdemProdVGCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjTelaPadrao As Object 'ct com codigo padrao
Public gobjTelaUsu As Object 'ct com customizacao

Public iMaquina As Integer

Public iGrid_Maquina2_Col  As Integer

Public sProdComboEmbalagens As String

'Grid de Carga
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoBase_Col As Integer
Public iGrid_Descricao_Col As Integer
Public iGrid_KgProdBase_Col As Integer
Public iGrid_KgCarga_Col As Integer
Public iGrid_QtdCarga_Col As Integer
Public iGrid_QtdTotal_Col As Integer

'Grid de Insumos
Public iGrid_Insumo_Col As Integer
Public iGrid_DescInsumo_Col As Integer
Public iGrid_KgCargaInsumo_Col As Integer
Public iGrid_Fornecedor_Col As Integer
Public iGrid_Lote_Col As Integer

Public objGridCarga As AdmGrid
Public objGridInsumo As AdmGrid

Public gcolItemOPCarga As New Collection




