VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImportCompl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntDoc As Long

Public iTipoDocOrigem As Integer  'IMPORTCOMPL_ORIGEM_DI, IMPORTCOMPL_ORIGEM_NF,...
Public lNumIntDocOrigem As Long

Public iSeq As Integer 'para posicionar no grid

Public iTipo As Integer
Public sDescricao As String
Public dValorBase As Double
Public dAliquota As Double
Public dValor As Double
Public dPerc As Double
Public iDias As Integer

Public iManual As Integer

Public objTipoImportCompl As ClassTipoImportCompl 'apenas para ganhar performance evitando leituras e simplificar o codigo
