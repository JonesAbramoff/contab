VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdicaoDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'adiçao de declaracao de importação

Public lNumIntDoc As Long

Public lNumIntDI As Long

Public iSeq As Integer 'sequencial da adicao dentro da DI

Public sIPICodigo As String

Public dValorAduaneiro As Double

Public dIIAliquota As Double
Public dIPIAliquota As Double
Public dPISAliquota As Double
Public dCOFINSAliquota As Double
Public dICMSAliquota As Double

Public dIIValor As Double
Public dIPIValor As Double
Public dPISValor As Double
Public dCOFINSValor As Double
Public dICMSValor As Double

Public dIPIBase As Double
Public dPISBase As Double
Public dCOFINSBase As Double
Public dICMSBase As Double

Public colItensAdicaoDI As New Collection

Public dDespesaAduaneira As Double
Public dTaxaSiscomex As Double
Public sCodFabricante As String

Public sDescricao As String

Public dICMSPercRedBase As Double

Public sNumDrawback As String
