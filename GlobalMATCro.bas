Attribute VB_Name = "GlobalMATCro"
Option Explicit

'#########################################
'Inserido por Wagner - CROMATON 17/07/07
Public Const STRING_GRUPO_PESAGEM = 1
'#########################################

'#########################################
'Inserido por Wagner 18/10/05
Public Const NUM_LINHAS_RELOPCARGA = 14
'#########################################

'#########################################
'Inserido por Wagner 21/09/05
Public Const STRING_MOVESTOQUE_RESPONSAVEL = 50
Public Const STRING_PEDIDOVENDA_OBSERVACAO = 250
'#########################################

'#########################################
'Inserido por Wagner - CROMATON 03/06/05
Public Const STRING_FICHAPROCESSO_TELA = 50
Public Const STRING_FICHAPROCESSO_OBSERVACAO = 255
Public Const STRING_FICHAPROCESSO_AMOSTRA = 255

Public Const FICHAPROCESSO_NUM_ZONAS_TC = 3
Public Const FICHAPROCESSO_NUM_ZONAS_TE = 7
'#########################################

'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Type typeOPFichaProcesso
     iFilialEmpresa As Integer
     sCodigoOP As String
     iMistura As Integer
     iTempoMistura As Integer
     iCargaDrays As Integer
     iEmbalagem As Integer
     '############################
     '03/06/05
     dProcessoAgua As Double
     dProcessoOleo As Double
     iAmperagem As Integer
     dVelRotoresDe As Double
     dVelRotoresAte As Double
     dAberturaGate As Double
     dTempCamaraZDe(1 To 3) As Double
     dTempCamaraZAte(1 To 3) As Double
     sTela As String
     dTempExtrusoraZDe(1 To 7) As Double
     dTempExtrusoraZAte(1 To 7) As Double
     dVelVariadorDe As Double
     dVelVariadorAte As Double
     sObservacao As String
     sAmostra As String
     '############################
End Type

Type typeItemOPCarga
    lNumIntItemOP As Long
    sProdutoBase As String
    dKgProdBase As Double
    dKgCarga As Double
    dQtdCarga As Double
End Type

Type typeItemOPCargaInsumo
    lNumIntDoc As Long
    lNumIntItemOP As Long
    sProduto As String
    dKgCarga As Double
    iSeq As Integer
    lFornecedor As Long
    sLote As String
End Type
'#########################################

'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Public Const PRODUTOKIT_PARTECARGA = 1
Public Const PRODUTOKIT_BASECARGA = 2
Public Const PRODUTOKIT_NAOCARGA = 3

Public Const PRODUTOKIT_STRING_PARTECARGA = "Faz parte da carga"
Public Const PRODUTOKIT_STRING_BASECARGA = "É base para carga"
Public Const PRODUTOKIT_STRING_NAOCARGA = "Não faz parte da carga"

Public Const ESTCFG_VALIDA_PRODUTO_BASE_CARGA = "VALIDA_PRODUTO_BASE_CARGA"

Public Const NAO_VALIDA_PRODUTO_BASE_CARGA = 0
Public Const VALIDA_PRODUTO_BASE_CARGA = 1

Public Const OP_FICHAPROC_MISTURA_TAMBOR = 1
Public Const OP_FICHAPROC_MISTURA_HENCHEL = 2

Public Const STRING_OP_FICHAPROC_MISTURA_TAMBOR = "Tambor"
Public Const STRING_OP_FICHAPROC_MISTURA_HENCHEL = "Henchel"

Public Const OP_FICHAPROC_EMBALAGEM_PEQUENA = 1
Public Const OP_FICHAPROC_EMBALAGEM_GRANDE = 2

Public Const STRING_OP_FICHAPROC_EMBALAGEM_PEQUENA = "Pequena"
Public Const STRING_OP_FICHAPROC_EMBALAGEM_GRANDE = "Grande"

Public Const OP_FICHAPROC_MISTURA_TAMBOR_5_15 = 1
Public Const OP_FICHAPROC_MISTURA_TAMBOR_10_20 = 2

Public Const STRING_OP_FICHAPROC_MISTURA_TAMBOR_5_15 = "5 a 15"
Public Const STRING_OP_FICHAPROC_MISTURA_TAMBOR_10_20 = "10 a 20"

Public Const OP_FICHAPROC_MISTURA_HENCHEL_3 = 3
Public Const OP_FICHAPROC_MISTURA_HENCHEL_4 = 4
Public Const OP_FICHAPROC_MISTURA_HENCHEL_5 = 5

Public Const OP_FICHAPROC_CARGADRAYS_4_6 = 1
Public Const OP_FICHAPROC_CARGADRAYS_BRANCO = 0

Public Const STRING_OP_FICHAPROC_CARGADRAYS_4_6 = "4 a 6"
Public Const STRING_OP_FICHAPROC_CARGADRAYS_BRANCO = ""
'#########################################

