VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFATPAGVGTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjTelaPadrao As Object 'ct com codigo padrao
Public gobjTelaUsu As Object 'ct com customizacao

Public lAporteAnt As Long
Public iTipoPagtoAnt As Integer
Public iIndexPagtoAnt As Integer
Public gobjAporte As New ClassTRVAportes
Public gobjPagtoCond As New ClassTRVAportePagtoCond
Public gobjPagtoDireto As New ClassTRVAportePagtoDireto
Public giTrazendoAporte As Integer

