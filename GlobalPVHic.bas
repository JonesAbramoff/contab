Attribute VB_Name = "GlobalPVHic"
Option Explicit

Type typePedidoVendaHic
    dValorFrete1 As Double
    dValorSeguro1 As Double
    dValorOutrasDespesas1 As Double
    dValorFrete2 As Double
    dValorSeguro2 As Double
    dValorOutrasDespesas2 As Double
    iFlagCompl1 As Integer
    iFlagCompl2 As Integer
End Type

Public Const NUM_MAXIMO_ITENS_HICARE = 19
