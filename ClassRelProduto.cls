VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sUltProduto As String

Private sReferencia As String
Private sNomeReduzido As String

Private dUltVlrUnitCompra As Double
Private dtDataUltCom As Date
Private dAliquotaICMSUltCom As Double
Private dFrete As Double

Private sUltProdutoQtd As String
Private dQtdPedCompra As Double
Private dQtdReqCompra As Double
Private dQtdPedCompraT As Double
Private dQtdReqCompraT As Double

Private lComando As Long
Private lComandoQtdeOv As Long

Public Function Obter_NomeReduzido(sNome As String, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_Obter_NomeReduzido
    
    If sProduto <> sUltProduto Then
        
        lErro = Obter_Produto(sProduto)
        If lErro <> SUCESSO Then gError 196905
    
    End If
    
    sNome = sNomeReduzido
    
    Obter_NomeReduzido = SUCESSO
    
    Exit Function
    
Erro_Obter_NomeReduzido:

    Obter_NomeReduzido = gErr
    
    Select Case gErr

        Case 196905
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196906)

    End Select

    Exit Function

End Function

Public Function Obter_Referencia(sRef As String, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_Obter_Referencia
    
    If sProduto <> sUltProduto Then
        
        lErro = Obter_Produto(sProduto)
        If lErro <> SUCESSO Then gError 196907
    
    End If
    
    sRef = sReferencia
    
    Obter_Referencia = SUCESSO
    
    Exit Function
    
Erro_Obter_Referencia:

    Obter_Referencia = gErr
    
    Select Case gErr

        Case 196907
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196908)

    End Select

    Exit Function

End Function

Private Function Obter_Produto(ByVal sProduto As String) As Long

Dim lErro As Long
Dim sNome As String
Dim sRef As String

On Error GoTo Erro_Obter_Produto

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 196909
    End If

    sNome = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    sRef = String(STRING_PRODUTO_REFERENCIA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NomeReduzido, Referencia FROM Produtos WHERE Codigo = ? ", sNome, sRef, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 196910
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196911
    
    If lErro <> AD_SQL_SUCESSO Then gError 196912
    
    sUltProduto = sProduto
    sReferencia = sRef
    sNomeReduzido = sNome

    Obter_Produto = SUCESSO

    Exit Function

Erro_Obter_Produto:

    Obter_Produto = gErr

    Select Case gErr

        Case 196909
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196910, 196911
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 196912
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196913)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Obter_UltVlrUnitCompra(dValor As Double, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Obter_UltVlrUnitCompra
    
    If sProduto <> sUltProduto Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 196907
        
        lErro = CF("Produto_ObterDadosUltEnt", objProduto, giFilialEmpresa, dUltVlrUnitCompra, dtDataUltCom, dAliquotaICMSUltCom, dFrete)
        If lErro <> SUCESSO Then gError 196907
    
    End If
    
    dValor = dUltVlrUnitCompra
    
    Obter_UltVlrUnitCompra = SUCESSO
    
    Exit Function
    
Erro_Obter_UltVlrUnitCompra:

    Obter_UltVlrUnitCompra = gErr
    
    Select Case gErr

        Case 196907
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196908)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_PedCompra(dQtd As Double, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Obter_Qtd_PedCompra
    
    If sProduto <> sUltProdutoQtd Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompra, dQtdPedCompra, giFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompraT, dQtdPedCompraT, giFilialEmpresa, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        sUltProdutoQtd = sProduto
    
    End If
    
    dQtd = dQtdPedCompra
    
    Obter_Qtd_PedCompra = SUCESSO
    
    Exit Function
    
Erro_Obter_Qtd_PedCompra:

    Obter_Qtd_PedCompra = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196908)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_ReqCompra(dQtd As Double, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Obter_Qtd_ReqCompra
    
    If sProduto <> sUltProdutoQtd Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompra, dQtdPedCompra, giFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompraT, dQtdPedCompraT, giFilialEmpresa, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        sUltProdutoQtd = sProduto
    
    End If
    
    dQtd = dQtdReqCompra
    
    Obter_Qtd_ReqCompra = SUCESSO
    
    Exit Function
    
Erro_Obter_Qtd_ReqCompra:

    Obter_Qtd_ReqCompra = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196908)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_PedCompraT(dQtd As Double, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Obter_Qtd_PedCompraT
    
    If sProduto <> sUltProdutoQtd Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompra, dQtdPedCompra, giFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompraT, dQtdPedCompraT, giFilialEmpresa, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        sUltProdutoQtd = sProduto
    
    End If
    
    dQtd = dQtdPedCompraT
    
    Obter_Qtd_PedCompraT = SUCESSO
    
    Exit Function
    
Erro_Obter_Qtd_PedCompraT:

    Obter_Qtd_PedCompraT = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196908)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_ReqCompraT(dQtd As Double, ByVal sProduto As String) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Obter_Qtd_ReqCompraT
    
    If sProduto <> sUltProdutoQtd Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompra, dQtdPedCompra, giFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("Produto_Le_Quantidades_Compras", objProduto, dQtdReqCompraT, dQtdPedCompraT, giFilialEmpresa, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        sUltProdutoQtd = sProduto
    
    End If
    
    dQtd = dQtdReqCompraT
    
    Obter_Qtd_ReqCompraT = SUCESSO
    
    Exit Function
    
Erro_Obter_Qtd_ReqCompraT:

    Obter_Qtd_ReqCompraT = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196908)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComandoQtdeOv <> 0 Then Call Comando_Fechar(lComandoQtdeOv)
End Sub

Public Function Obter_FatorConvUM(dFator As Double, ByVal iClasseUM As Integer, ByVal sSiglaUMDe As String, ByVal sSiglaUMPara As String) As Long
'obter o fator pelo qual deve ser multiplicada a qtde na sSilgaUMDe para que fique na sSiglaUMPara

Dim lErro As Long
Dim lComando As Long, dQtde As Double
Dim lComando1 As Long, dqtde1 As Double

On Error GoTo Erro_Obter_FatorConvUM
    
    lComando = 0
    lComando1 = 0
    
    If sSiglaUMDe <> sSiglaUMPara Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 201214
        lComando1 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando1 = 0 Then gError 201215
    
        lErro = Comando_Executar(lComando, "SELECT Quantidade FROM UnidadesDeMedida WHERE Classe = ? AND Sigla = ?", dQtde, iClasseUM, sSiglaUMDe)
        If lErro <> AD_SQL_SUCESSO Then gError 201216
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201217
        If lErro <> AD_SQL_SUCESSO Then gError 201218
        
        lErro = Comando_Executar(lComando1, "SELECT Quantidade FROM UnidadesDeMedida WHERE Classe = ? AND Sigla = ?", dqtde1, iClasseUM, sSiglaUMPara)
        If lErro <> AD_SQL_SUCESSO Then gError 201219
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201220
        If lErro <> AD_SQL_SUCESSO Then gError 201221
        
        dFator = dQtde / dqtde1
        
        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando1)
        
    Else
        dFator = 1
    End If
    
    Obter_FatorConvUM = SUCESSO
    
    Exit Function
    
Erro_Obter_FatorConvUM:

    Obter_FatorConvUM = gErr
    
    Select Case gErr

        Case 201214, 201215
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201216 To 201221
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201222)

    End Select

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComando1 <> 0 Then Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function Obter_QtdeOVProdPer(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal lStatusMin As Long, dQtdeOV As Double) As Long
'Para obter a qtde em ov por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_QtdeOVProdPer

    If lComandoQtdeOv = 0 Then
        lComandoQtdeOv = Comando_AbrirExt(GL_lConexaoRel)
        If lComandoQtdeOv = 0 Then gError 196909
    End If
    
    lErro = CF("Prod_Obtem_QtdeEmOVPerRec", dtDataInicial, dtDataFinal, iFilialEmpresa, sProduto, lStatusMin, dQtdeOV, lComandoQtdeOv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Obter_QtdeOVProdPer = SUCESSO

    Exit Function

Erro_Obter_QtdeOVProdPer:

    Obter_QtdeOVProdPer = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 196909
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152793)

    End Select

    Call Comando_Fechar(lComandoQtdeOv)

    Exit Function

End Function


