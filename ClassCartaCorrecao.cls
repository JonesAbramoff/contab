VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCartaCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntDoc As Long
Public sCorrecao As String
Public schNFe As String
Public inSeqEvento As Integer

Public iFilialEmpresa As Integer
Public lidLote As Long

Public scStat As String
Public dtDataReg As Date
Public dHoraReg As Double
