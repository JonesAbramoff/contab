VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMascaraGenerica"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function InitMasc Lib "ADCUSR.DLL" Alias "AD_Sistema_InitMasc" (ByVal lID_Sistema As Long, ByVal sNomeMascara As String) As Long

Private mvarcolMascara As New Collection

Private Function Retorna_Mascara(sNomeMascara As String) As ClassMascara

Dim objMascaraGenerica1 As ClassMascaraGenerica1
Dim objMascara As ClassMascara

    For Each objMascaraGenerica1 In mvarcolMascara
        If LCase(objMascaraGenerica1.sNomeMascara) = LCase(sNomeMascara) Then Exit For
    Next
    
    If objMascaraGenerica1 Is Nothing Then
        Set objMascaraGenerica1 = New ClassMascaraGenerica1
        objMascaraGenerica1.sNomeMascara = sNomeMascara
        Set objMascaraGenerica1.objMascara = New ClassMascara
        Call objMascaraGenerica1.objMascara.Init(0, InitMasc(GL_lSistema, sNomeMascara))
        mvarcolMascara.Add objMascaraGenerica1
    End If
    
    Set Retorna_Mascara = objMascaraGenerica1.objMascara

End Function

Function NumSegItem(sNomeMascara As String) As Long

Dim objMascara As ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    NumSegItem = objMascara.NumSeg()
    
End Function

Function SegmentoItem(sNomeMascara As String, iNivel As Integer, objSegmento As ClassSegmento) As Long
'obtem a descricao de um segmento da mascara de produtos

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)
    
    SegmentoItem = objMascara.Segmento(iNivel, objSegmento)

End Function

Function Mascara_MascararItem(sNomeMascara As String, sProduto As String, sProdutoMascarado As String) As Long
'recebe o produto no formato do bd e o coloca mascarado (com delimitadores)

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_MascararItem = objMascara.Mascara_MascararElem(sProduto, sProdutoMascarado)

End Function

Function Mascara_RetornaItemNoNivel(sNomeMascara As String, ByVal i_nivel As Integer, ByVal lp_prod As String, lp_prod_no_nivel As String) As Long
'recebe um produto no formato do bd e retorna o produto no i_nivel

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_RetornaItemNoNivel = objMascara.Mascara_RetornaElemNoNivel(i_nivel, lp_prod, lp_prod_no_nivel)

End Function

Function Mascara_RetornaItemPai(sNomeMascara As String, ByVal lp_prod As String, lpProdPai As String) As Long
'coloca em lpProdPai o pai de lp_prod. Se nao houver coloca "".

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_RetornaItemPai = objMascara.Mascara_RetornaElemPai(lp_prod, lpProdPai)

End Function

Function Mascara_RetornaItemEnxuto(sNomeMascara As String, ByVal lp_prod As String, lpProdutoEnxuto As String) As Long
'pega o produto no formato do banco de dados e retorna somente com os segmentos preenchidos

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_RetornaItemEnxuto = objMascara.Mascara_RetornaElemEnxuto(lp_prod, lpProdutoEnxuto)

End Function

Function Mascara_RetornaItemTela(sNomeMascara As String, ByVal lp_prod As String, lpProdutoTela As String) As Long
'pega o produto no formato do banco de dados e retorna-o como se tivesse sido colocado numa masked edit

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_RetornaItemTela = objMascara.Mascara_RetornaElemTela(lp_prod, lpProdutoTela)

End Function

Function Mascara_RetornaUltimoItem(sNomeMascara As String, ByVal lp_prod As String, lpMaiorProduto As String) As Long
'pega o produto no formato do banco de dados e retorna o maior produto possivel no nivel dele,
'para o produto pai (no nivel imediatamente acima)

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)
 
    Mascara_RetornaUltimoItem = objMascara.Mascara_RetornaUltimoElem(lp_prod, lpMaiorProduto)

End Function

Function Mascara_ExisteProxNivelItem(sNomeMascara As String, ByVal i_nivel As Integer, ByVal lp_prod As String) As Long
'se existe ou nao um nivel abaixo (maior)

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_ExisteProxNivelItem = objMascara.Mascara_ExisteProxNivelElem(i_nivel, lp_prod)
    
End Function

Function MascaraItem(sNomeMascara As String, sMascara As String) As Long
'coloca em sMascara a mascara do produto

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    MascaraItem = objMascara.MascaraElem(sMascara)

End Function

Function Mascara_Item_ObterNivel(sNomeMascara As String, sProduto As String, i_nivel As Integer) As Long
'retorna nivel do Produto

Dim objMascara As New ClassMascara
    
    Set objMascara = Retorna_Mascara(sNomeMascara)

    Mascara_Item_ObterNivel = objMascara.Mascara_Elem_ObterNivel(sProduto, i_nivel)

End Function

Private Sub Class_Initialize()
    If GL_objMascaraGenerica Is Nothing Then
        Set GL_objMascaraGenerica = Me
    End If
End Sub
