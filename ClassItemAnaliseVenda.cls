VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemAnaliseVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sProduto As String
Public sDescricao As String
Public sUnidadeMed As String
Public dQuantidade As Double
Public dPrecoBase As Double
Public dPrecoUnitario As Double
Public dPercDesc As Double
Public dValorDesconto As Double
Public dPrecoTotal As Double
Public dPrecoTotalB As Double
Public iTabelaPreco As Integer
Public iIncluiValorTotal As Integer
