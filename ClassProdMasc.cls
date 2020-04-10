VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdMasc"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NumSegProduto() As Long

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    NumSegProduto = mvarobjMascara.NumSeg()
    
End Function

Function SegmentoProduto(iNivel As Integer, objSegmento As ClassSegmento) As Long
'obtem a descricao de um segmento da mascara de produtos

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    SegmentoProduto = mvarobjMascara.Segmento(iNivel, objSegmento)

End Function

Function Mascara_MascararProduto(sProduto As String, sProdutoMascarado As String) As Long
'recebe o produto no formato do bd e o coloca mascarado (com delimitadores)

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_MascararProduto = mvarobjMascara.Mascara_MascararElem(sProduto, sProdutoMascarado)

End Function

Function Mascara_RetornaProdutoNoNivel(ByVal i_nivel As Integer, ByVal lp_prod As String, lp_prod_no_nivel As String) As Long
'recebe um produto no formato do bd e retorna o produto no i_nivel

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_RetornaProdutoNoNivel = mvarobjMascara.Mascara_RetornaElemNoNivel(i_nivel, lp_prod, lp_prod_no_nivel)

End Function

Function Mascara_RetornaProdutoPai(ByVal lp_prod As String, lpProdPai As String) As Long
'coloca em lpProdPai o pai de lp_prod. Se nao houver coloca "".

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_RetornaProdutoPai = mvarobjMascara.Mascara_RetornaElemPai(lp_prod, lpProdPai)

End Function

Function Mascara_RetornaProdutoEnxuto(ByVal lp_prod As String, lpProdutoEnxuto As String) As Long
'pega o produto no formato do banco de dados e retorna somente com os segmentos preenchidos

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_RetornaProdutoEnxuto = mvarobjMascara.Mascara_RetornaElemEnxuto(lp_prod, lpProdutoEnxuto)

End Function

Function Mascara_RetornaProdutoTela(ByVal lp_prod As String, lpProdutoTela As String) As Long
'pega o produto no formato do banco de dados e retorna-o como se tivesse sido colocado numa masked edit

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_RetornaProdutoTela = mvarobjMascara.Mascara_RetornaElemTela(lp_prod, lpProdutoTela)

End Function

Function Mascara_RetornaUltimoProduto(ByVal lp_prod As String, lpMaiorProduto As String) As Long
'pega o produto no formato do banco de dados e retorna o maior produto possivel no nivel dele,
'para o produto pai (no nivel imediatamente acima)

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)
 
    Mascara_RetornaUltimoProduto = mvarobjMascara.Mascara_RetornaUltimoElem(lp_prod, lpMaiorProduto)

End Function

Function Mascara_ExisteProxNivelProd(ByVal i_nivel As Integer, ByVal lp_prod As String) As Long
'se existe ou nao um nivel abaixo (maior)

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_ExisteProxNivelProd = mvarobjMascara.Mascara_ExisteProxNivelElem(i_nivel, lp_prod)
    
End Function

Function MascaraProduto(sMascara As String) As Long
'coloca em sMascara a mascara do produto

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    MascaraProduto = mvarobjMascara.MascaraElem(sMascara)

End Function

Function Mascara_Produto_ObterNivel(sProduto As String, i_nivel As Integer) As Long
'retorna nivel do Produto

Dim mvarobjMascara As New ClassMascara
    
    Call mvarobjMascara.Init(STRING_PRODUTO, GL_lMascProd)

    Mascara_Produto_ObterNivel = mvarobjMascara.Mascara_Elem_ObterNivel(sProduto, i_nivel)

End Function

Private Sub Class_Initialize()
    If GL_lMascProd = 0 Then Call Init_Fest '??? depois mover p/inicializacao do modulo
End Sub
