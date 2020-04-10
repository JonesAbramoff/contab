VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassExtrBcoCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iCodConta As Integer
Public iNumExtrato As Integer
Public dtDataGravacao As Date
Public iConciliado As Integer
Public sNomeArquivo As String
Public dtDataSaldoInicial As Date
Public dSaldoInicial As Double
Public dtDataSaldoFinal As Date
Public dSaldoFinal As Double
Public dHoraGeracao As Double
Public dHoraImport As Double
Public dtDataImport As Date
Public lSeqArquivo As Long
Public iBanco As Integer
