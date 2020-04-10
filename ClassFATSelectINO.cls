VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Produto_Saida_Celula_PrecoDesc(ByVal objGridItens As Object, ByVal iIndice As Integer, ByVal iGrid_PrecoDesc_Col As Integer, ByVal dPrecoUnitario As Double, ByVal sProduto As String, ByVal iGrid_DescProduto_Col As Integer, ByVal sCliente As String, ByVal iFilialCliente As Integer, ByVal dDescontoCli As Double, ByVal iGrid_PercDesc_Col As Integer)

Dim lErro As Long, objProd As New ClassProduto

On Error GoTo Erro_Produto_Saida_Celula_PrecoDesc

    objProd.sCodigo = sProduto
    
    lErro = Produto_Le_Desc_INO(objProd)
    If lErro <> SUCESSO Then gError 99999

    'Descricao Produto
    objGridItens.TextMatrix(objGridItens.Row, iGrid_DescProduto_Col) = objProd.sDescricao

    Exit Sub
    
Erro_Produto_Saida_Celula_PrecoDesc:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192977)
     
    End Select
     
    Exit Sub

End Sub

Public Function PV_Produto_Saida_Celula5(ByVal objCT As Object, ByVal objProd As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Produto_Saida_Celula5

    lErro = Produto_Le_Desc_INO(objProd)
    If lErro <> SUCESSO Then gError 99999

    'Descricao Produto
    If objCT.ActiveControl Is objCT.DescricaoProduto Then
        objCT.DescricaoProduto.Text = objProd.sDescricao
    Else
        objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_DescProduto_Col) = objProd.sDescricao
    End If

    PV_Produto_Saida_Celula5 = SUCESSO
     
    Exit Function
    
Erro_PV_Produto_Saida_Celula5:

    PV_Produto_Saida_Celula5 = gErr
     
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192977)
     
    End Select
     
    Exit Function

End Function

Public Function NF_Produto_Saida_Celula(ByVal objCT As Object, ByVal objProd As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_NF_Produto_Saida_Celula

    lErro = Produto_Le_Desc_INO(objProd)
    If lErro <> SUCESSO Then gError 99999

    'Descricao Produto
    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_DescProduto_Col) = objProd.sDescricao

    NF_Produto_Saida_Celula = SUCESSO
     
    Exit Function
    
Erro_NF_Produto_Saida_Celula:

    NF_Produto_Saida_Celula = gErr
     
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192977)
     
    End Select
     
    Exit Function

End Function

Public Function Produto_Le_Desc_INO(ByVal objProd As ClassProduto) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim sDescricao As String
Dim sModelo As String
Dim sCodProdForn As String
Dim sDescFinal As String

On Error GoTo Erro_Produto_Le_Desc_INO

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198520
    Next
    
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sModelo = String(STRING_PRODUTO_MODELO, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Descricao, Modelo FROM Produtos WHERE Codigo = ? ", sDescricao, sModelo, objProd.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198511

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198512
    
    sCodProdForn = String(STRING_PRODUTO, 0)

    lErro = Comando_Executar(alComando(2), "SELECT ProdutoFornecedor FROM FornecedorProdutoFF WHERE Produto = ? ", sCodProdForn, objProd.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 198511

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198512
 
    objProd.sDescricao = sDescricao & SEPARADOR & sCodProdForn & SEPARADOR & sModelo
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Produto_Le_Desc_INO = SUCESSO
    
    Exit Function
    
Erro_Produto_Le_Desc_INO:

    Produto_Le_Desc_INO = gErr
    
    Select Case gErr
    
        Case 198520
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 198511, 198512
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198529)

    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

