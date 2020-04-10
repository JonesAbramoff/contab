VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelCompCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntDocItemOP As Long
Public sCodigoOP As String
Public sProdutoPai As String
Public sProdutoPaiDescricao As String
Public dtDataFinalProducao As Date
Public sUMProdutoPai As String
Public dQtdeProduzida As Double
Public sKitVersao As String
Public dKitPerda As Double
Public dCustoContabil As Double
Public dQtdeReqPropProd As Double
Public iClasseUMPai As Integer

Public colItens As New Collection

Public objParam As ClassRelCompConsParam



