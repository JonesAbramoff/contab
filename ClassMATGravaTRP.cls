VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGravaTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Grava_Usu(ByVal objProduto As ClassProduto, ByVal colTabelaPrecoItem As Collection) As Long
'para gravar parte customizada do produto

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim iIndice As Integer
Dim objProdutoTRP As ClassProdutoTRP
Dim objTarifa As ClassTRPProdTarifa
Dim iDiasDe As Integer
Dim iDestino As Variant
Dim sTextoAux As String
Dim iCount As Integer

On Error GoTo Erro_Produto_Grava_Usu

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197239
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRPProdutos WHERE Codigo = ?", objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197241
    
    lErro = Comando_Executar(alComando(5), "DELETE FROM TRPProdTextoVou WHERE Produto = ?", objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197241
    
    lErro = Comando_Executar(alComando(6), "DELETE FROM TRPProdDestino WHERE Produto = ?", objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197241
    
    If Not (objProduto.objInfoUsu Is Nothing) Then
    
        Set objProdutoTRP = objProduto.objInfoUsu

        lErro = Comando_Executar(alComando(2), "INSERT INTO TRPProdutos (Codigo, Vendedor, PercComissao, LimiteIdade, DescontoTitular, DescontoAcompanhante, MaxAcompanhantes, Moeda, ExibeWeb,AceitaIdadeMaiorLimite, PercAcrescIdadeMaiorLimite, TodosDestValidos, VigenciaFixa ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objProduto.sCodigo, objProdutoTRP.iVendedor, objProdutoTRP.dPercComissao, objProdutoTRP.iLimiteIdade, objProdutoTRP.dDescontoTitular, objProdutoTRP.dDescontoAcompanhante, objProdutoTRP.iMaxAcompanhante, objProdutoTRP.iMoeda, objProdutoTRP.iExibeWeb, objProdutoTRP.iAceitaIdadeMaiorLimite, objProdutoTRP.dPercAcrescIdadeMaiorLimite, objProdutoTRP.iTodosDestValidos, objProdutoTRP.iVigenciaFixa)
        If lErro <> AD_SQL_SUCESSO Then gError 197242
    
        If objProdutoTRP.dtDataVigencia <> DATA_NULA Then
    
            lErro = Comando_Executar(alComando(3), "DELETE FROM TRPProdTarifa WHERE Produto = ? AND DataVigencia = ?", objProduto.sCodigo, objProdutoTRP.dtDataVigencia)
            If lErro <> AD_SQL_SUCESSO Then gError 197241
            
            For Each objTarifa In objProdutoTRP.colTarifas
            
                If objTarifa.dtVigencia = objProdutoTRP.dtDataVigencia Then
            
                    lErro = Comando_Executar(alComando(4), "INSERT INTO TRPProdTarifa (Produto, DiasDe, DiasAte, Diario, Valor, ValorAdicional, DataVigencia ) VALUES (?,?,?,?,?,?,?)", objTarifa.sProduto, objTarifa.iDiasDe, objTarifa.iDiasAte, objTarifa.iDiario, objTarifa.dValor, objTarifa.dValorAdicional, objTarifa.dtVigencia)
                    If lErro <> AD_SQL_SUCESSO Then gError 197242
            
                End If
            
            Next
    
        End If
        
        For Each iDestino In objProdutoTRP.colDestinos
            lErro = Comando_Executar(alComando(7), "INSERT INTO TRPProdDestino (Produto, Destino) VALUES (?,?)", objProdutoTRP.sProduto, iDestino)
            If lErro <> AD_SQL_SUCESSO Then gError 197242
        Next
        
        sTextoAux = objProdutoTRP.sTextoVouPT
        iCount = 0
        Do While sTextoAux <> ""
            iCount = iCount + 1
            lErro = Comando_Executar(alComando(8), "INSERT INTO TRPProdTextoVou (Produto, Seq, Idioma, Texto) VALUES (?,?, ?,?)", objProdutoTRP.sProduto, iCount, 0, Left(sTextoAux, 250))
            If lErro <> AD_SQL_SUCESSO Then gError 197242
            If Len(sTextoAux) > 250 Then
                sTextoAux = Right(sTextoAux, Len(sTextoAux) - 250)
            Else
                sTextoAux = ""
            End If
        Loop
        
        sTextoAux = objProdutoTRP.sTextoVouEN
        iCount = 0
        Do While sTextoAux <> ""
            iCount = iCount + 1
            lErro = Comando_Executar(alComando(8), "INSERT INTO TRPProdTextoVou (Produto, Seq, Idioma, Texto) VALUES (?,?, ?,?)", objProdutoTRP.sProduto, iCount, 1, Left(sTextoAux, 250))
            If lErro <> AD_SQL_SUCESSO Then gError 197242
            If Len(sTextoAux) > 250 Then
                sTextoAux = Right(sTextoAux, Len(sTextoAux) - 250)
            Else
                sTextoAux = ""
            End If
        Loop
    
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

        Case 197241
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPPRODUTOS", gErr)

        Case 197242
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRPPRODUTOS", gErr)

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

    lErro = Comando_Executar(alComando(1), "DELETE FROM TRPProdutos WHERE Codigo = ?", sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 197245

    lErro = Comando_Executar(alComando(2), "DELETE FROM TRPProdTarifa WHERE Produto = ?", sProduto)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPPRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197246)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

