VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdicaoDIComplNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'??? eliminar esta classe

Public lNumIntAdicaoDINF As Long
Public iSeqNaAdicaoDINF As Integer 'sequencial do complemento dentro da adicaonf

Public dBaseCalculo As Double
Public dAliquota As Double
Public dValor As Double

Public iTipo As Integer

Public iManual As Integer


