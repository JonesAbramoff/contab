Attribute VB_Name = "GlobalRotBaixaCartao"
Option Explicit

Public Const ADMEXTFIN_VARIAS = 0
Public Const ADMEXTFIN_VISANET = 1
Public Const ADMEXTFIN_REDECARD = 2
Public Const ADMEXTFIN_AMEX = 3
Public Const ADMEXTFIN_ELO = 7

Public Const ADMEXTFIN_TIPOMOV_ROBXCARTAO = 1

Public Const ADMEXTFIN_TIPOMOVDET_BXCARTAO = 1

Public Const ADMEXTFIN_ERRO_CV_SEMPARCELA = 1
Public Const ADMEXTFIN_ERRO_CV_MUITASPARCELAS = 2
Public Const ADMEXTFIN_ERRO_CV_PARCNAOABERTA = 3
Public Const ADMEXTFIN_ERRO_CV_DIFNUMCARTAO = 4

Public Function AAAAMMDD_ParaDate(ByVal sData As String) As Date
    If sData = "00000000" Then
        AAAAMMDD_ParaDate = DATA_NULA
    Else
        AAAAMMDD_ParaDate = StrParaDate(right(sData, 2) & "/" & Mid(sData, 5, 2) & "/" & left(sData, 4))
    End If
End Function

Public Function AAMMDD_ParaDate(ByVal sData As String) As Date
    If sData = "000000" Then
        AAMMDD_ParaDate = DATA_NULA
    Else
        AAMMDD_ParaDate = StrParaDate(right(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & left(sData, 2))
    End If
End Function

Public Function TiraZerosEsq(ByVal sData As String) As String
Dim iPOS As Integer

    If sData = "" Then
        TiraZerosEsq = ""
    Else
        iPOS = 1
        Do While Mid(sData, iPOS, 1) = "0"
            iPOS = iPOS + 1
        Loop
        TiraZerosEsq = Mid(sData, iPOS)
    End If
    

End Function




