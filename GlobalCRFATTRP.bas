Attribute VB_Name = "GlobalCRFATTRP"
Option Explicit

'Tipo do Título para Comissões
Public Const TIPO_COMISSAO_TRP = 6

Public Const VENDEDOR_CARGO_TRP_PROMOTOR = 1
Public Const VENDEDOR_CARGO_TRP_SUPERVISOR = 2
Public Const VENDEDOR_CARGO_TRP_GERENTE = 3
Public Const VENDEDOR_CARGO_TRP_DIRETOR = 4

Type typeVendedorTRP
    iCargo As Integer
    iSuperior As Integer
End Type

Type typeVNDComissaoTRP
    lNumIntDoc As Long
    iVendedor As Integer
    iSeq As Integer
    dValorDe As Double
    dValorAte As Double
    iMoeda As Integer
    dPercComissao As Double
End Type

Type typeVNDReducaoTRP
    lNumIntDoc As Long
    iVendedor As Integer
    iSeq As Integer
    dValorDe As Double
    dValorAte As Double
    iMoeda As Integer
    dPercComissaoMax As Double
End Type

Type typeVNDRegiaoTRP
    lNumIntDoc As Long
    iVendedor As Integer
    iSeq As Integer
    iRegiaoVenda As Integer
    dPercComissao As Double
End Type

