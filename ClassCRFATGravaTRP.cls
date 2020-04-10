VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGravaTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Vendedor_Grava_Customizacao(ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim objVendedorTRP As ClassVendedorTRP
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objVNDComissaoTRP As ClassVNDComissaoTRP
Dim objVNDReducaoTRP  As ClassVNDReducaoTRP
Dim objVNDRegiaoTRP As ClassVNDRegiaoTRP

On Error GoTo Erro_Vendedor_Grava_Customizacao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195872
    Next

    If Not (objVendedor.objInfoUsu Is Nothing) Then
    
        Set objVendedorTRP = objVendedor.objInfoUsu
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM VendedorTRP WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195873
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO VendedorTRP (Vendedor, Cargo, Superior) VALUES (?,?,?)", objVendedor.iCodigo, objVendedorTRP.iCargo, objVendedorTRP.iSuperior)
        If lErro <> AD_SQL_SUCESSO Then gError 195874
    
        lErro = Comando_Executar(alComando(3), "DELETE FROM VendedorTRPComissao WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195958
    
        lErro = Comando_Executar(alComando(4), "DELETE FROM VendedorTRPComiRed WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195959
    
        lErro = Comando_Executar(alComando(5), "DELETE FROM VendedorTRPRegVenda WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195960
    
        If objVendedorTRP.colVNDComissaoTRP.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_VENDEDORTRPCOMISSAO", lNumIntDoc, objVendedorTRP.colVNDComissaoTRP.Count)
            If lErro <> SUCESSO Then gError 195961

            iIndice = 0

            For Each objVNDComissaoTRP In objVendedorTRP.colVNDComissaoTRP
            
                iIndice = iIndice + 1
                
                objVNDComissaoTRP.lNumIntDoc = lNumIntDoc
                objVNDComissaoTRP.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(6), "INSERT INTO VendedorTRPComissao (NumIntDoc, Vendedor, Seq, ValorDe, ValorAte, Moeda, PercComissao) VALUES (?,?,?,?,?,?,?)", _
                         objVNDComissaoTRP.lNumIntDoc, objVendedor.iCodigo, objVNDComissaoTRP.iSeq, objVNDComissaoTRP.dValorDe, objVNDComissaoTRP.dValorAte, objVNDComissaoTRP.iMoeda, objVNDComissaoTRP.dPercComissao)
                If lErro <> AD_SQL_SUCESSO Then gError 195962
    
            Next
    
        End If
    
        If objVendedorTRP.colVNDReducaoTRP.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_VENDEDORTRPCOMIRED", lNumIntDoc, objVendedorTRP.colVNDReducaoTRP.Count)
            If lErro <> SUCESSO Then gError 195963

            iIndice = 0

            For Each objVNDReducaoTRP In objVendedorTRP.colVNDReducaoTRP
            
                iIndice = iIndice + 1
                
                objVNDReducaoTRP.lNumIntDoc = lNumIntDoc
                objVNDReducaoTRP.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(7), "INSERT INTO VendedorTRPComiRed (NumIntDoc, Vendedor, Seq, ValorDe, ValorAte, Moeda, PercComissaoMax) VALUES (?,?,?,?,?,?,?)", _
                         objVNDReducaoTRP.lNumIntDoc, objVendedor.iCodigo, objVNDReducaoTRP.iSeq, objVNDReducaoTRP.dValorDe, objVNDReducaoTRP.dValorAte, objVNDReducaoTRP.iMoeda, objVNDReducaoTRP.dPercComissaoMax)
                If lErro <> AD_SQL_SUCESSO Then gError 195964
    
            Next
    
        End If
    
        If objVendedorTRP.colVNDRegiaoTRP.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRPConfig", "NUM_INT_PROX_VENDEDORTRPREGVENDA", lNumIntDoc, objVendedorTRP.colVNDRegiaoTRP.Count)
            If lErro <> SUCESSO Then gError 195965

            iIndice = 0

            For Each objVNDRegiaoTRP In objVendedorTRP.colVNDRegiaoTRP
            
                iIndice = iIndice + 1
                
                objVNDRegiaoTRP.lNumIntDoc = lNumIntDoc
                objVNDRegiaoTRP.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(8), "INSERT INTO VendedorTRPRegVenda (NumIntDoc, Vendedor, Seq, RegiaoVenda, PercComissao) VALUES (?,?,?,?,?)", _
                         objVNDRegiaoTRP.lNumIntDoc, objVendedor.iCodigo, objVNDRegiaoTRP.iSeq, objVNDRegiaoTRP.iRegiaoVenda, objVNDRegiaoTRP.dPercComissao)
                If lErro <> AD_SQL_SUCESSO Then gError 195966
    
            Next
    
        End If
    
    End If
    
    lErro = CF("TRPUsuarios_Grava", objVendedorTRP.objUsuarioWeb)
    If lErro <> SUCESSO Then gError 195965

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Vendedor_Grava_Customizacao = SUCESSO

    Exit Function

Erro_Vendedor_Grava_Customizacao:

    Vendedor_Grava_Customizacao = gErr

    Select Case gErr

        Case 195872
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195873
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRP", gErr)

        Case 195874
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_VENDEDORTRP", gErr)
        
        Case 195958
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPCOMISSAO", gErr)

        Case 195959
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPCOMIRED", gErr)

        Case 195960
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPREGVENDA", gErr)

        Case 195961, 195963, 195965
            
        Case 195962
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPCOMISSAO", gErr)

        Case 195964
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPCOMIRED", gErr)

        Case 195966
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPREGVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195875)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Vendedor_Exclui_Customizado(ByVal objVendedor As ClassVendedor) As Long
'customizacoes da rotina de exclusao de vendedor

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim objVendedorTRP As ClassVendedorTRP
Dim iIndice As Integer

On Error GoTo Erro_Vendedor_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195892
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM VendedorTRP WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195893

    lErro = Comando_Executar(alComando(2), "DELETE FROM VendedorTRPComissao WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195955

    lErro = Comando_Executar(alComando(3), "DELETE FROM VendedorTRPComiRed WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195956

    lErro = Comando_Executar(alComando(4), "DELETE FROM VendedorTRPRegVenda WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195957
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Vendedor_Exclui_Customizado = SUCESSO

    Exit Function

Erro_Vendedor_Exclui_Customizado:

    Vendedor_Exclui_Customizado = gErr

    Select Case gErr

        Case 195892
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 195893
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRP", gErr)

        Case 195955
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPCOMISSAO", gErr)

        Case 195956
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPCOMIRED", gErr)

        Case 195957
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRPREGVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195894)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

