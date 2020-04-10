VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTVendedoresVGTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_objUserControl As Object

Public objGridComissao As AdmGrid
Public objGridRedComissao As AdmGrid
Public objGridRegiaoVenda As AdmGrid

Public iGrid_ComissValorDe_Col As Integer
Public iGrid_ComissValorAte_Col As Integer
Public iGrid_ComissMoeda_Col As Integer
Public iGrid_PercComiss_Col As Integer

Public iGrid_RedValorDe_Col As Integer
Public iGrid_RedValorAte_Col As Integer
Public iGrid_RedMoeda_Col As Integer
Public iGrid_PercRed_Col As Integer

Public iGrid_CodRegiao_Col As Integer
Public iGrid_DescRegiao_Col As Integer
Public iGrid_PercComissReg_Col As Integer

Public gobjTelaUsu As Object

Public gobjUsuarioWeb As ClassTRPUsuarios



