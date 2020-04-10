VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemAdicaoDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntDoc As Long

Public lNumIntAdicaoDI As Long

Public iAdicao As Integer

Public iSeq As Integer 'sequencial do item dentro da adicao

Public sProduto As String
Public sDescricao As String
Public sUM As String
Public dQuantidade As Double

Public dValorUnitFOBNaMoeda As Double
Public dValorUnitFOBEmReal As Double
Public dValorUnitCIFNaMoeda As Double
Public dValorUnitCIFEmReal As Double
Public dValorTotalFOBNaMoeda As Double
Public dValorTotalFOBEmReal As Double
Public dValorTotalCIFNaMoeda As Double
Public dValorTotalCIFEmReal As Double

Public iTotalCIFEmRealManual As Integer

Public dPesoBruto As Double
Public dPesoLiq As Double

Public dValorUnitTrib As Double

Public dIPIUnidadePadraoValor As Double
