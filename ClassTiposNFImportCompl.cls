VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTiposNFImportCompl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iCodigo As Integer

Public sDescReduzida As String
Public sDescricao As String
Public iIncluiValorProdutos As Integer
Public iIncluiBaseICMS As Integer
Public iImpressaoNaNF As Integer '0:nao imprime, 1:no corpo, 2:como obs
Public iSeqImpressaoNF As Integer 'para manter sempre a mesma ordem
Public iLinhaPadraoNaTela As Integer 'zero = nao entra, senao a linha do grid
Public iComAliquota As Integer
