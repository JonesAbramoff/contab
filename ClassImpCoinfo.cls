VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sSiglaEmpresa As String 'TVA ou TVI

Public sArqTipoEmpr As String
Public sArqTipoEmprHora As String
Public lNumIntArqTipoEmpr As Long
Public lNumRegsArqTipoEmpr As Long

Public sArqGrupo As String
Public sArqGrupoHora As String
Public lNumIntArqGrupo As Long
Public lNumRegsArqGrupo As Long

Public sArqEmpr As String
Public sArqEmprHora As String
Public lNumIntArqEmpr As Long
Public lNumRegsArqEmpr As Long

Public sArqProd As String
Public sArqProdHora As String
Public lNumIntArqProd As Long
Public lNumRegsArqProd As Long

Public sArqFat As String
Public sArqFatHora As String
Public lNumIntArqFat As Long
Public lNumRegsArqFat As Long

Public sArqItemFat As String
Public sArqItemFatHora As String
'Public lNumIntArqItemFat As Long
Public lNumRegsArqItemFat As Long

Public sArqOcor As String
Public sArqOcorHora As String
Public lNumIntArqOcor As Long
Public lNumRegsArqOcor As Long

Public sArqNF As String
Public sArqNFHora As String
Public lNumIntArqNF As Long
Public lNumRegsArqNF As Long

Public sArqItemNF As String
Public sArqItemNFHora As String
'Public lNumIntArqItemNF As Long
Public lNumRegsArqItemNF As Long

Public sArqComisOver As String
Public sArqComisOverHora As String
Public lNumIntArqComisOver As Long
Public lNumRegsArqComisOver As Long

Public sArqEmiss As String
Public sArqEmissHora As String
Public lNumIntArqEmiss As Long
Public lNumRegsArqEmiss As Long

Public sArqComissao As String
Public sArqComissaoHora As String
Public lNumIntArqComissao As Long
Public lNumRegsArqComissao As Long

Public sArqAcordo As String
Public sArqAcordoHora As String
Public lNumIntArqAcordo As Long
Public lNumRegsArqAcordo As Long

Public sArqAcordoComis As String
Public sArqAcordoComisHora As String
Public lNumIntArqAcordoComis As Long
Public lNumRegsArqAcordoComis As Long

Public sArqDestino As String
Public sArqDestinoHora As String
Public lNumIntArqDestino As Long
Public lNumRegsArqDestino As Long

Public sDiretorio As String
Public sDiretorioNovo As String

Public objFiltro As Object
