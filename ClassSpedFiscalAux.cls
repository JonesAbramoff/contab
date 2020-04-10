VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedFiscalAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colRegistros As New Collection
Public objSpedFiscal As New ClassSpedFiscal

Public Sub Inicia_Reg()
    Set colRegistros = New Collection
End Sub

Public Sub Reg_Campo_add(ByVal iIndice As Integer, ByVal sCampo As String, ByVal sValor As String)
    Dim objCampo As New ClassSpedFiscalReg
    objCampo.iIndice = iIndice
    objCampo.sCampo = sCampo
    objCampo.sValor = Replace(sValor, "|", " ")
    colRegistros.Add objCampo
End Sub

Public Sub Finaliza_Reg()
    Call Monta_Reg
    Set colRegistros = Nothing
End Sub

Private Sub Monta_Reg()
Dim sReg As String
Dim objCampo As New ClassSpedFiscalReg
    sReg = "|"
    For Each objCampo In colRegistros
        sReg = sReg & objCampo.sValor & "|"
    Next
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedFiscal.sRegistro = sReg
End Sub


