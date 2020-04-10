VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGravaTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Grava_Usu(ByVal objProduto As ClassProduto, ByVal colTabelaPrecoItem As Collection) As Long
'para gravar parte customizada do produto

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objProdutoTRV As ClassProdutoTRV

On Error GoTo Erro_Produto_Grava_Usu

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197239
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRVProdutos WHERE Codigo = ?", objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197241
    
    If Not (objProduto.objInfoUsu Is Nothing) Then
    
        Set objProdutoTRV = objProduto.objInfoUsu

        lErro = Comando_Executar(alComando(2), "INSERT INTO TRVProdutos (Codigo, Vendedor, PercComissao) VALUES (?,?,?)", objProduto.sCodigo, objProdutoTRV.iCodVendedor, objProdutoTRV.dPercComissao)
        If lErro <> AD_SQL_SUCESSO Then gError 197242
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_Grava_Usu = SUCESSO

    Exit Function

Erro_Produto_Grava_Usu:

    Produto_Grava_Usu = gErr

    Select Case gErr

        Case 197239
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197240
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVPRODUTOS", gErr)

        Case 197241
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRVPRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197243)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Produto_Exclui_Usu(ByVal sProduto As String) As Long
'para gravar parte customizada do produto

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_Produto_Exclui_Usu

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197244
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRVProdutos WHERE Codigo = ?", sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 197245

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_Exclui_Usu = SUCESSO

    Exit Function

Erro_Produto_Exclui_Usu:

    Produto_Exclui_Usu = gErr

    Select Case gErr

        Case 197244
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197245
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVPRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197246)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

