VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTClientesVGTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_objUserControl As Object

Public iGrid_Emissor_Col As Integer
Public iGrid_PercComiss_Col As Integer

Public iGrid_Produto_Col As Integer
Public iGrid_DescricaoProduto_Col As Integer
Public iGrid_PercComissProd_Col As Integer

Public gcolCliEmissoresTRP As Collection

Public objGridComissao As AdmGrid
Public objGridExcecoes As AdmGrid

Public gobjTelaUsu As Object 'ct com customizacao

Public giLinhaAntiga As Integer

Public objGridExcRep As AdmGrid
Public iGrid_ExcRep_Col As Integer
Public iGrid_ExcRepProduto_Col As Integer
Public iGrid_ExcRepPercComiss_Col As Integer

Public objGridExcCor As AdmGrid
Public iGrid_ExcCor_Col As Integer
Public iGrid_ExcCorProduto_Col As Integer
Public iGrid_ExcCorPercComiss_Col As Integer

Public objGridExcAg As AdmGrid
Public iGrid_ExcAgProduto_Col As Integer
Public iGrid_ExcAgPercComiss_Col As Integer

Public gobjUsuarioWeb As ClassTRPUsuarios



