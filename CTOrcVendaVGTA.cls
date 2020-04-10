VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcVendaVGTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjTelaPadrao As Object 'ct com codigo padrao
Public gobjTelaUsu As Object 'ct com customizacao

Public iGrid_PrecoBase_Col As Integer
Public iGrid_PercAcresc_Col As Integer
Public iGrid_Acrescimo_Col As Integer

Public dValorAcrescimoItensAnt As Double
Public dPercAcrescimoItensAnt As Double
Public iAcrescimoAlterado As Integer


