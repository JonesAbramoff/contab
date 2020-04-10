VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelCompConsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntRel As Long
Public iFilialEmpresa As Integer
Public sProdutoFilho As String
Public sProdutoFilhoDescricao As String
Public sUMProdutoFilho As String
Public dQtdeConsumida As Double
Public sKitVersao As String
Public dKitPerda As Double
Public dKitQtdePrevista As Double
Public dCustoContabil As Double
Public iSequencial As Integer

