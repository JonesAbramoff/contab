VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntArq As Long
Public lNumRegArqTexto As Long
Public sTipVou As String
Public sSerie As String
Public lNumVou As Long
Public sTipDoc As String
Public dValBase As Double
Public dValComi As Double
Public lrecnsica As Long
Public dPercComi As Double
Public iExportado As Integer
Public dtDataExp As Date
Public lcodemp As Long

Public iTipoDocCorporator As Integer
Public lNumIntDocCorporator As Long
Public dtNoCorporatorEm As Date
Public iComErro As Integer

