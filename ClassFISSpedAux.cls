VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISSpedAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public objSpedFiscal As ClassSpedFiscal

Public dtDataIni As Date
Public dtDataFim As Date

Public iFilialEmpresa As Integer

Public iIncluiRCPE As Integer

Public objSpedRegistroG001 As ClassSpedRegistro
Public objSpedRegistroG110 As ClassSpedRegistro
Public objSpedRegistroG125 As ClassSpedRegistro
Public objSpedRegistroG126 As ClassSpedRegistro
Public objSpedRegistroG130 As ClassSpedRegistro
Public objSpedRegistroG140 As ClassSpedRegistro
Public objSpedRegistroG990 As ClassSpedRegistro

Public objSpedRegistro0150 As ClassSpedRegistro
Public objSpedRegistro0175 As ClassSpedRegistro
Public objSpedRegistro0190 As ClassSpedRegistro
Public objSpedRegistro0200 As ClassSpedRegistro
Public objSpedRegistro0205 As ClassSpedRegistro
Public objSpedRegistro0210 As ClassSpedRegistro
Public objSpedRegistro0220 As ClassSpedRegistro
Public objSpedRegistro0300 As ClassSpedRegistro
Public objSpedRegistro0305 As ClassSpedRegistro
Public objSpedRegistro0500 As ClassSpedRegistro
Public objSpedRegistro0600 As ClassSpedRegistro

Public objSpedRegistroK001 As ClassSpedRegistro
Public objSpedRegistroK100 As ClassSpedRegistro
Public objSpedRegistroK200 As ClassSpedRegistro
Public objSpedRegistroK220 As ClassSpedRegistro
Public objSpedRegistroK230 As ClassSpedRegistro
Public objSpedRegistroK235 As ClassSpedRegistro
Public objSpedRegistroK250 As ClassSpedRegistro
Public objSpedRegistroK255 As ClassSpedRegistro
Public objSpedRegistroK990 As ClassSpedRegistro

Public objSpedRegistro1001 As ClassSpedRegistro
Public objSpedRegistro1010 As ClassSpedRegistro
Public objSpedRegistro1100 As ClassSpedRegistro
Public objSpedRegistro1105 As ClassSpedRegistro
Public objSpedRegistro1400 As ClassSpedRegistro

Public objSpedRegistroC101 As ClassSpedRegistro
Public objSpedRegistroC191 As ClassSpedRegistro

Public objSpedRegistroE300 As ClassSpedRegistro
Public objSpedRegistroE310 As ClassSpedRegistro
Public objSpedRegistroE311 As ClassSpedRegistro
Public objSpedRegistroE312 As ClassSpedRegistro
Public objSpedRegistroE313 As ClassSpedRegistro
Public objSpedRegistroE316 As ClassSpedRegistro

Public colUM As Collection
Public colC101 As New Collection
Public colC101UF As New Collection


Public Property Get iAno() As Integer
    iAno = Year(dtDataIni)
End Property

Public Property Get iMes() As Integer
    iMes = Month(dtDataIni)
End Property

