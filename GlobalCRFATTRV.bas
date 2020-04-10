Attribute VB_Name = "GlobalCRFATTRV"
Option Explicit

'Tipo do Título para Comissões
Public Const TIPO_COMISSAO_TRV = 6

Public Const VENDEDOR_CARGO_TRV_PROMOTOR = 1
Public Const VENDEDOR_CARGO_TRV_SUPERVISOR = 2
Public Const VENDEDOR_CARGO_TRV_GERENTE = 3
Public Const VENDEDOR_CARGO_TRV_DIRETOR = 4

Type typeVendedorTRV
    iCargo As Integer
    iSuperior As Integer
    dPercCallCenter As Double
End Type

Type typeVNDComissaoTRV
    lNumIntDoc As Long
    iVendedor As Integer
    iSeq As Integer
    dValorDe As Double
    dValorAte As Double
    iMoeda As Integer
    dPercComissao As Double
End Type

Type typeVNDReducaoTRV
    lNumIntDoc As Long
    iVendedor As Integer
    iSeq As Integer
    dValorDe As Double
    dValorAte As Double
    iMoeda As Integer
    dPercComissaoMax As Double
End Type

Type typeVNDRegiaoTRV
    lNumIntDoc As Long
    iVendedor As Integer
    iSeq As Integer
    iRegiaoVenda As Integer
    dPercComissao As Double
End Type

