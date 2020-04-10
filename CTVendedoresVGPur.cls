VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTVendedoresVGPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_objUserControl As Object

Public objGridExc As AdmGrid

Public iGrid_ExcProd_Col As Integer
Public iGrid_ExcProdDesc_Col As Integer
Public iGrid_ExcMeta_Col As Integer
Public iGrid_ExcPercComiss_Col As Integer
Public iGrid_ExcMetaP_Col As Integer
Public iGrid_ExcPercComissP_Col As Integer

Public gobjTelaUsu As Object


