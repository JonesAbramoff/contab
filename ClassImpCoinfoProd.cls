VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntArq As Long
Public sCodPro As String
Public sDescricao As String
Public iCodSeg As Integer
Public sAlcance As String
Public smoeda As String
Public sTipProd As String
Public dISS As Double
Public dExento As Double
Public dVlPadrao As Double
Public dPercDesc As Double
Public iLFamiliar As Integer
Public dPercTitular As Double
Public dPercAcomp As Double
Public dcomismax As Double
Public iLimIdade As Integer
Public sUsuario As String
Public dtdatareg As Date
Public iValidade As Integer
Public dcomispro As Double
Public iLcatalogo As Integer
Public lEmpPai As Long
Public iDiasVal As Integer
Public imaxfamil As Integer
Public isupinf
Public iExportado As Integer
Public dtDataExp As Date
