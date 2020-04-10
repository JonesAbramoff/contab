VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelectTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Le_Customizado(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objProdutoTRV As New ClassProdutoTRV
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim iVendedor As Integer
Dim dPercComissao As Double

On Error GoTo Erro_Produto_Le_Customizado

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197235
    Next

    lErro = Comando_Executar(alComando(1), "SELECT Vendedor, PercComissao FROM TRVProdutos WHERE Codigo = ? ", _
                iVendedor, dPercComissao, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197236

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197237

    If lErro = AD_SQL_SUCESSO Then
        objProdutoTRV.iCodVendedor = iVendedor
        objProdutoTRV.dPercComissao = dPercComissao
    End If
    
    Set objProduto.objInfoUsu = objProdutoTRV

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_Le_Customizado = SUCESSO

    Exit Function

Erro_Produto_Le_Customizado:

    Produto_Le_Customizado = gErr

    Select Case gErr

        Case 197235
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197236, 197237
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVPRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197238)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

