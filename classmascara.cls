VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMascara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Mascara_MascararElemInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_MascaraExt" (ByVal lID_Mascara As Long, ByVal sContaNaoMascarada As String, ByVal sContaMascarada As String) As Long
Private Declare Function Mascara_RetornaElemNoNivelInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_RetornaNoNivelExt" (ByVal lID_Mascara As Long, ByVal i_nivel As Integer, ByVal lp_elem As String, ByVal lp_elem_no_nivel As String) As Long
Private Declare Function Mascara_RetornaElemPaiInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_RetornaPaiExt" (ByVal lID_Mascara As Long, ByVal lp_elem As String, ByVal lp_elem_pai As String) As Long
Private Declare Function Mascara_RetornaElemEnxutoInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_RetornaEnxuto" (ByVal lID_Mascara As Long, ByVal lp_elem As String, ByVal lpElemEnxuto As String) As Long
Private Declare Function Mascara_RetornaElemTelaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_RetornaTela" (ByVal lID_Mascara As Long, ByVal lp_elem As String, ByVal lpElemEnxuto As String) As Long
Private Declare Function Mascara_RetornaMaiorElemInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_RetornaMaior" (ByVal lID_Mascara As Long, ByVal lp_elem As String, ByVal lpMaiorElem As String) As Long
Private Declare Function Mascara_ExisteProxNivelElemInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_ExisteProxNivel" (ByVal lID_Mascara As Long, ByVal i_nivel As Integer, ByVal lp_elem As String) As Long
Private Declare Function Mascara_Elem_ObterInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_ObterExt" (ByVal lID_Mascara As Long, ByVal sMascaraConta As String) As Long
Private Declare Function Mascara_Elem_ObterNivelInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_ObterNivel" (ByVal lID_Mascara As Long, ByVal lp_elem As String, i_nivel As Integer) As Long
Private Declare Function NumSegElemInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_NumSeg" (ByVal lID_Mascara As Long) As Long
Private Declare Function SegmentoElemInt Lib "ADCUSR.DLL" Alias "AD_Mascara_Elem_Segmento" (ByVal lID_Mascara As Long, ByVal iNivel As Integer, iPreenchimento As Integer, iTamanho As Integer, iTipo As Integer, ByVal sDelimitador As String) As Long

Private mvariTamMaxElem As Integer
Private mvarlMascara As Long

Function NumSeg() As Long
    NumSeg = NumSegElemInt(mvarlMascara)
End Function

Function Segmento(iNivel As Integer, objSegmento As ClassSegmento) As Long
'obtem a descricao de um segmento da mascara
Dim lErro As Long
Dim tSegmento As typeSegmento
    
On Error GoTo Erro_Segmento

    tSegmento.sDelimitador = " "
    
    lErro = SegmentoElemInt(mvarlMascara, iNivel, tSegmento.iPreenchimento, tSegmento.iTamanho, tSegmento.iTipo, tSegmento.sDelimitador)
    If lErro <> AD_BOOL_TRUE Then Error 27000
    
    objSegmento.iPreenchimento = tSegmento.iPreenchimento
    objSegmento.iTamanho = tSegmento.iTamanho
    objSegmento.iTipo = tSegmento.iTipo
    objSegmento.sDelimitador = tSegmento.sDelimitador
    
    Segmento = SUCESSO
    
    Exit Function
    
Erro_Segmento:

    Segmento = Err
    
    Select Case Err
    
        Case 27000
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151547)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_MascararElem(sElem As String, sElemMascarado As String) As Long
'recebe o elemento no formato do bd e o coloca mascarado (com delimitadores)
Dim lErro As Long

On Error GoTo Erro_Mascara_MascararElem

    'p/garantir que haja espaco p/o retorno
    sElemMascarado = String(255, 0)
    
    lErro = Mascara_MascararElemInt(mvarlMascara, sElem, sElemMascarado)
    If lErro <> AD_BOOL_TRUE Then Error 27001
    
    sElemMascarado = StringZ(sElemMascarado)
    
    Mascara_MascararElem = SUCESSO
    
    Exit Function
    
Erro_Mascara_MascararElem:

    Mascara_MascararElem = Err
    
    Select Case Err
    
        Case 27001
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151548)
    
    End Select
    
    Exit Function

End Function

Function Mascara_RetornaElemNoNivel(ByVal i_nivel As Integer, ByVal lp_elem As String, lp_elem_no_nivel As String) As Long
'recebe um elemento no formato do bd e retorna o elemento no i_nivel
Dim lErro As Long

On Error GoTo Erro_Mascara_RetornaElemNoNivel

    lErro = Mascara_RetornaElemNoNivelInt(mvarlMascara, i_nivel, lp_elem, lp_elem_no_nivel)
    If lErro <> AD_BOOL_TRUE Then Error 27002
    
    lp_elem_no_nivel = StringZ(lp_elem_no_nivel)
    
    Mascara_RetornaElemNoNivel = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaElemNoNivel:

    Mascara_RetornaElemNoNivel = Err
    
    Select Case Err
    
        Case 27002
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151549)
    
    End Select
    
    Exit Function

End Function


Function Mascara_RetornaElemPai(ByVal lp_elem As String, lpElemPai As String) As Long
'coloca em lpElemPai o pai de lp_elem. Se nao houver coloca "".

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaElemPai

    lErro = Mascara_RetornaElemPaiInt(mvarlMascara, lp_elem, lpElemPai)
    If lErro <> AD_BOOL_TRUE Then Error 27003
    
    lpElemPai = StringZ(lpElemPai)
    Mascara_RetornaElemPai = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaElemPai:

    Mascara_RetornaElemPai = Err
    
    Select Case Err
    
        Case 27003
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151550)
    
    End Select
    
    Exit Function
    
End Function


Function Mascara_RetornaElemEnxuto(ByVal lp_elem As String, lpElemEnxuto As String) As Long
'pega o elemento no formato do banco de dados e retorna somente com os segmentos preenchidos

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaElemEnxuto

    lpElemEnxuto = String(255, 0)
    
    lErro = Mascara_RetornaElemEnxutoInt(mvarlMascara, lp_elem, lpElemEnxuto)
    If lErro <> AD_BOOL_TRUE Then Error 27004
    
    lpElemEnxuto = StringZ(lpElemEnxuto)
    
    Mascara_RetornaElemEnxuto = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaElemEnxuto:

    Mascara_RetornaElemEnxuto = Err
    
    Select Case Err
    
        Case 27004
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151551)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaElemTela(ByVal lp_elem As String, lpElemTela As String) As Long
'pega o elemento no formato do banco de dados e retorna-o como se tivesse sido colocado numa masked edit

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaElemTela

    lpElemTela = String(255, 0)
    
    lErro = Mascara_RetornaElemTelaInt(mvarlMascara, lp_elem, lpElemTela)
    If lErro <> AD_BOOL_TRUE Then Error 27004
    
    lpElemTela = StringZ(lpElemTela)
    
    Mascara_RetornaElemTela = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaElemTela:

    Mascara_RetornaElemTela = Err
    
    Select Case Err
    
        Case 59396
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151552)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaUltimoElem(ByVal lp_elem As String, lpMaiorElem As String) As Long
'pega o elemento no formato do banco de dados e retorna o maior elemento possivel no nivel dele,
'para o elemento pai (no nivel imediatamente acima)

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaUltimoElem

    lpMaiorElem = String(255, 0)
    
    lErro = Mascara_RetornaMaiorElemInt(mvarlMascara, lp_elem, lpMaiorElem)
    If lErro <> AD_BOOL_TRUE Then Error 27005
    
    lpMaiorElem = StringZ(lpMaiorElem)
    
    Mascara_RetornaUltimoElem = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaUltimoElem:

    Mascara_RetornaUltimoElem = Err
    
    Select Case Err
    
        Case 27005
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151553)
    
    End Select
    
    Exit Function
    
End Function


Function Mascara_ExisteProxNivelElem(ByVal i_nivel As Integer, ByVal lp_elem As String) As Long
'se existe ou nao um nivel abaixo (maior)
Dim lErro As Long
    
On Error GoTo Erro_Mascara_ExisteProxNivelElem

    lErro = Mascara_ExisteProxNivelElemInt(mvarlMascara, i_nivel, lp_elem)
    If lErro <> AD_BOOL_TRUE Then Error 27006
    
    Mascara_ExisteProxNivelElem = SUCESSO
    
    Exit Function
    
Erro_Mascara_ExisteProxNivelElem:

    Mascara_ExisteProxNivelElem = Err
    
    Select Case Err
    
        Case 27006
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151554)
    
    End Select
    
    Exit Function
    
End Function

Function MascaraElem(sMascara As String) As Long
'coloca em sMascara a mascara do elemento

Dim lErro As Long

On Error GoTo Erro_MascaraElem

    sMascara = String(255, 0)

    lErro = Mascara_Elem_ObterInt(mvarlMascara, sMascara)
    If lErro <> AD_BOOL_TRUE Then Error 27007
    
    sMascara = StringZ(sMascara)
    
    MascaraElem = SUCESSO
    
    Exit Function
    
Erro_MascaraElem:

    MascaraElem = Err
    
    Select Case Err
    
        Case 27007
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151555)
    
    End Select
    
    Exit Function

End Function


Function Mascara_Elem_ObterNivel(sElem As String, i_nivel As Integer) As Long
'retorna nivel do Elem
Dim lErro As Long

On Error GoTo Erro_Mascara_Elem_ObterNivel

    lErro = Mascara_Elem_ObterNivelInt(mvarlMascara, sElem, i_nivel)
    If lErro <> AD_BOOL_TRUE Then Error 27008
    
    Mascara_Elem_ObterNivel = SUCESSO
    
    Exit Function
    
Erro_Mascara_Elem_ObterNivel:

    Mascara_Elem_ObterNivel = Err
    
    Select Case Err
    
        Case 27008
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151556)
    
    End Select
    
    Exit Function
    
End Function

Sub Init(iTamMaxElem As Integer, lMascara As Long)
    mvariTamMaxElem = iTamMaxElem
    mvarlMascara = lMascara
End Sub
