VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFImportCompl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntNFImportacao As Long

Public iSeq As Integer 'para posicionar no grid

Public iTipo As Integer
Public sDescricao As String
Public dValor As Double
Public sObservacao As String

Public iManual As Integer 'se o calculo dele foi manual (informado pelo usuario) ou nao (calculado pelo sistema.
