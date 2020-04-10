VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntDoc As Long 'excuir este campo

Public lNumIntNF As Long
Public lNumIntDI As Long

Public colAdicoesNF As New Collection
    'com objetos ClassAdicaoDINF
    
Public colComplNF As New Collection
    'para valores como despesas aduaneiras, pis e cofins da importacao
    'com objetos ClassImportCompl

Public colItensNF As New Collection
    'com objetos ClassItemAdicaoDIItemNF

Public objDIInfo As New ClassDIInfo

