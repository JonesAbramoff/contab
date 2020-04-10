VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoEmissor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntArq As Long
Public lNumRegArqTexto As Long

Public lCodEmi As Long
Public sapelido As String
Public snome As String
Public lcodemp As Long
Public semail As String
Public sobs As String
Public sTipoFat As String
Public iAtivo As Integer
Public sUsuario As String
Public dpercover As Double
Public slogin As String
Public ssenha As String
Public ialterar As Integer
Public dtdataalt As Date
Public icondpag As Integer
Public lemiger As Long
Public lemisup As Long
Public lcodimas As Long
Public iExportado As Integer
Public dtDataExp As Date

Public iTipoDocCorporator As Integer
Public lNumIntDocCorporator As Long
Public dtNoCorporatorEm As Date
Public iComErro As Integer
