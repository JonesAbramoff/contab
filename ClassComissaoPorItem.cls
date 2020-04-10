VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissaoPorItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iTipoDoc As Integer
Public lNumIntDocItem As Long
Public iSeqComissao As Integer
Public dValorBase As Double
Public dPercentual As Double
Public dValor As Double
Public dPercentualEmissao As Double
Public dValorEmissao As Double
Public iLinha As Integer
