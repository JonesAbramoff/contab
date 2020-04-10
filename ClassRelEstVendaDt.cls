VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelEstVendaDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'parametros

'de entrada
Public iFilialEmpresa As Integer
Public dtData1 As Date
Public iTabelaPreco As Integer
Public iAlmoxarifado As Integer
Public iTipoProdutoInicial As Integer
Public iTipoProdutoFinal As Integer
Public sProdInicial As String
Public sProdFinal As String
Public lFornecedor As Long

'de saida
Public lNumIntRel As Long

