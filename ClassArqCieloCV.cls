VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqCieloCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sReg As String
Public lNumRegArq As Long

Public Function PreencherMovDet(ByVal objMovDet As ClassAdmExtFinMovDet) As Long

    With objMovDet
        .iTipo = ADMEXTFIN_TIPOMOVDET_BXCARTAO
        .dValor = dValorParcela
        .sNumCartao = left(sNumCartao, 16)
        .dtDataCompra = dtCompraAjuste
        .iNumParcela = iParcela
        .sAutorizacao = TiraZerosEsq(sAutorizacao)
        .sRO = sRO
        .sNSU = sNSU
    End With
       
End Function

Public Property Get sTID() As String
    sTID = Mid(sReg, 73, 20)
End Property

Public Property Get sNSU() As String
    sNSU = Mid(sReg, 93, 6)
End Property

Public Property Get sRO() As String
    sRO = Mid(sReg, 12, 7)
End Property

Public Property Get sNumCartao() As String
    sNumCartao = Mid(sReg, 19, 19)
End Property

Public Property Get dtCompraAjuste() As Date
    dtCompraAjuste = AAAAMMDD_ParaDate(Mid(sReg, 38, 8))
End Property

Public Property Get dValorParcela() As Double
    
    dValorParcela = StrParaDbl(Mid(sReg, 47, 13))
    If Mid(sReg, 46, 1) = "-" Then dValorParcela = -dValorParcela
    
End Property

Public Property Get iParcela() As Integer
    iParcela = StrParaInt(Mid(sReg, 60, 2))
End Property

Public Property Get iTotalParcelas() As Integer
    iTotalParcelas = StrParaInt(Mid(sReg, 62, 2))
End Property

Public Property Get sMotivoRejeicao() As String
    sMotivoRejeicao = Mid(sReg, 64, 3)
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = Mid(sReg, 67, 6)
End Property
