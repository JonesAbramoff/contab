Attribute VB_Name = "GlobalFatDan"
Option Explicit

Public Const STRING_DAN_OS = 9
Public Const STRING_DAN_OS_MODELO = 50
Public Const STRING_DAN_OS_NUMSERIE = 20

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeDan_OS
    sOS As String
    lCliente As Long
    sModelo As String
    sNumSerie As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeDan_ItensOS
    sOS As String
    iItem As Integer
    sProduto As String
    dQuantidade As Double
End Type

