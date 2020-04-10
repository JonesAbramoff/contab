VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProdutoVGTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_objUserControl As Object

Public gobjTelaUsu As Object 'ct com customizacao

Public objGridTarifa As AdmGrid
Public iGrid_DiasDe_Col As Integer
Public iGrid_DiasAte_Col As Integer
Public iGrid_Diario_Col As Integer
Public iGrid_ValorTarifa_Col As Integer
Public iGrid_ValorTarifaAdicional_Col As Integer

Public gobjProduto As ClassProduto

Public sDataAnt As String

Public objGridDestinos As AdmGrid
Public iGrid_Destino_Col As Integer



