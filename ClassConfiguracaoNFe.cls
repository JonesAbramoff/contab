VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConfiguracaoNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sToken As String
Public sNFCeSerie As String
Public lNFCeProximoNum As Long
Public lNFCeProximoLote As Long
Public sDirArqXml As String
Public sCertificadoA1A3 As String
Public iNFeAmbiente As Integer
Public sNFeSerie As String
Public lNFeProximoNum As Long
Public lNFeProximoLote As Long
Public sDirXsd As String
Public iNFCeAmbiente As Integer
Public sNFCECSC As String
Public sidNFCECSC As String

Public iEmContingencia As Integer
Public dtContingenciaDataEntrada As Date
Public dContingenciaHoraEntrada As Double
Public sContigenciaxJust As String

Public iModeloImpressora As Integer
Public sPortaImpressora As String

Public sSMTP As String
Public sSMTPUsu As String
Public sSMTPSenha As String
Public lSMTPPorta As Long
Public iNFCeImprimir As Integer
Public iNFCeEnviarEmail As Integer

Public iNFDescricaoProd As Integer '0:usar o nome reduzido, 1: usar descricao
Public iVersaoNFe As Integer
Public iFocaTipoVenda As Integer
