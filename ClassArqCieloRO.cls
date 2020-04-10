VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqCieloRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sReg As String
Public lNumIntDoc As Long
Public colCV As New Collection
Public lNumRegArq As Long

Public iCodConta As Integer

Public Property Get sEstabelecimento() As String
    sEstabelecimento = Mid(sReg, 2, 10)
End Property

Public Property Get iBanco() As Integer
    iBanco = StrParaInt(Mid(sReg, 100, 4))
End Property

Public Property Get iAgencia() As Integer
    iAgencia = StrParaInt(Mid(sReg, 104, 5))
End Property

Public Property Get iBandeira() As Integer
Dim sAg As Integer
    sAg = Mid(sReg, 185, 3)
    Select Case sAg
        Case "001" 'VISA
            iBandeira = ADMEXTFIN_VISANET
        Case "002" 'MASTER
            iBandeira = ADMEXTFIN_REDECARD
        Case "007" 'ELO
            iBandeira = ADMEXTFIN_ELO
    End Select
End Property

Public Property Get sProduto() As String
    sProduto = Mid(sReg, 131, 2)
End Property

Public Property Get sConta() As String
    sConta = Mid(sReg, 109, 14)
End Property

Public Property Get sRO() As String
    sRO = Mid(sReg, 12, 7)
End Property

Public Property Get sStatus() As String
    sStatus = Mid(sReg, 123, 2)
End Property

Public Property Get dtDeposito() As Date
    dtDeposito = AAMMDD_ParaDate(Mid(sReg, 32, 6))
End Property

Public Property Get dtPrevPag() As Date
    dtPrevPag = AAMMDD_ParaDate(Mid(sReg, 32, 6))
End Property

Public Property Get dtEnvBco() As Date
    dtEnvBco = AAMMDD_ParaDate(Mid(sReg, 38, 6))
End Property

Public Property Get dValorBruto() As Double
    
    dValorBruto = StrParaDbl(Mid(sReg, 45, 13))
    If Mid(sReg, 44, 1) = "-" Then dValorBruto = -dValorBruto
    
End Property

Public Property Get dValorComissao() As Double
    
    dValorComissao = StrParaDbl(Mid(sReg, 59, 13))
    If Mid(sReg, 58, 1) = "-" Then dValorComissao = -dValorComissao
    
End Property

Public Property Get dValorRejeitado() As Double
    
    dValorRejeitado = StrParaDbl(Mid(sReg, 73, 13))
    If Mid(sReg, 72, 1) = "-" Then dValorRejeitado = -dValorRejeitado
    
End Property

Public Property Get dValorLiq() As Double
    
    dValorLiq = StrParaDbl(Mid(sReg, 87, 13))
    If Mid(sReg, 86, 1) = "-" Then dValorLiq = -dValorLiq
    
End Property

