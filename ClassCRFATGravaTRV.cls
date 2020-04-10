VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGravaTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Vendedor_Grava_Customizacao(ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim objVendedorTRV As ClassVendedorTRV
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objVNDComissaoTRV As ClassVNDComissaoTRV
Dim objVNDReducaoTRV  As ClassVNDReducaoTRV
Dim objVNDRegiaoTRV As ClassVNDRegiaoTRV

On Error GoTo Erro_Vendedor_Grava_Customizacao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195872
    Next

    If Not (objVendedor.objInfoUsu Is Nothing) Then
    
        Set objVendedorTRV = objVendedor.objInfoUsu
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM VendedorTRV WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195873
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO VendedorTRV (Vendedor, Cargo, Superior, PercCallCenter) VALUES (?,?,?,?)", objVendedor.iCodigo, objVendedorTRV.iCargo, objVendedorTRV.iSuperior, objVendedorTRV.dPercCallCenter)
        If lErro <> AD_SQL_SUCESSO Then gError 195874
    
        lErro = Comando_Executar(alComando(3), "DELETE FROM VendedorTRVComissao WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195958
    
        lErro = Comando_Executar(alComando(4), "DELETE FROM VendedorTRVComiRed WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195959
    
        lErro = Comando_Executar(alComando(5), "DELETE FROM VendedorTRVRegVenda WHERE Vendedor = ?", objVendedor.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 195960
    
        If objVendedorTRV.colVNDComissaoTRV.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_VENDEDORTRVCOMISSAO", lNumIntDoc, objVendedorTRV.colVNDComissaoTRV.Count)
            If lErro <> SUCESSO Then gError 195961

            iIndice = 0

            For Each objVNDComissaoTRV In objVendedorTRV.colVNDComissaoTRV
            
                iIndice = iIndice + 1
                
                objVNDComissaoTRV.lNumIntDoc = lNumIntDoc
                objVNDComissaoTRV.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(6), "INSERT INTO VendedorTRVComissao (NumIntDoc, Vendedor, Seq, ValorDe, ValorAte, Moeda, PercComissao) VALUES (?,?,?,?,?,?,?)", _
                         objVNDComissaoTRV.lNumIntDoc, objVendedor.iCodigo, objVNDComissaoTRV.iSeq, objVNDComissaoTRV.dValorDe, objVNDComissaoTRV.dValorAte, objVNDComissaoTRV.iMoeda, objVNDComissaoTRV.dPercComissao)
                If lErro <> AD_SQL_SUCESSO Then gError 195962
    
            Next
    
        End If
    
        If objVendedorTRV.colVNDReducaoTRV.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_VENDEDORTRVCOMIRED", lNumIntDoc, objVendedorTRV.colVNDReducaoTRV.Count)
            If lErro <> SUCESSO Then gError 195963

            iIndice = 0

            For Each objVNDReducaoTRV In objVendedorTRV.colVNDReducaoTRV
            
                iIndice = iIndice + 1
                
                objVNDReducaoTRV.lNumIntDoc = lNumIntDoc
                objVNDReducaoTRV.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(7), "INSERT INTO VendedorTRVComiRed (NumIntDoc, Vendedor, Seq, ValorDe, ValorAte, Moeda, PercComissaoMax) VALUES (?,?,?,?,?,?,?)", _
                         objVNDReducaoTRV.lNumIntDoc, objVendedor.iCodigo, objVNDReducaoTRV.iSeq, objVNDReducaoTRV.dValorDe, objVNDReducaoTRV.dValorAte, objVNDReducaoTRV.iMoeda, objVNDReducaoTRV.dPercComissaoMax)
                If lErro <> AD_SQL_SUCESSO Then gError 195964
    
            Next
    
        End If
    
        If objVendedorTRV.colVNDRegiaoTRV.Count > 0 Then
        
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_VENDEDORTRVREGVENDA", lNumIntDoc, objVendedorTRV.colVNDRegiaoTRV.Count)
            If lErro <> SUCESSO Then gError 195965

            iIndice = 0

            For Each objVNDRegiaoTRV In objVendedorTRV.colVNDRegiaoTRV
            
                iIndice = iIndice + 1
                
                objVNDRegiaoTRV.lNumIntDoc = lNumIntDoc
                objVNDRegiaoTRV.iSeq = iIndice
                
                lNumIntDoc = lNumIntDoc + 1
    
                lErro = Comando_Executar(alComando(8), "INSERT INTO VendedorTRVRegVenda (NumIntDoc, Vendedor, Seq, RegiaoVenda, PercComissao) VALUES (?,?,?,?,?)", _
                         objVNDRegiaoTRV.lNumIntDoc, objVendedor.iCodigo, objVNDRegiaoTRV.iSeq, objVNDRegiaoTRV.iRegiaoVenda, objVNDRegiaoTRV.dPercComissao)
                If lErro <> AD_SQL_SUCESSO Then gError 195966
    
            Next
    
        End If
    
    End If

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
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRV", gErr)

        Case 195874
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_VENDEDORTRV", gErr)
        
        Case 195958
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVCOMISSAO", gErr)

        Case 195959
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVCOMIRED", gErr)

        Case 195960
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVREGVENDA", gErr)

        Case 195961, 195963, 195965
            
        Case 195962
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVCOMISSAO", gErr)

        Case 195964
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVCOMIRED", gErr)

        Case 195966
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVREGVENDA", gErr)

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
Dim objVendedorTRV As ClassVendedorTRV
Dim iIndice As Integer

On Error GoTo Erro_Vendedor_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195892
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM VendedorTRV WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195893

    lErro = Comando_Executar(alComando(2), "DELETE FROM VendedorTRVComissao WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195955

    lErro = Comando_Executar(alComando(3), "DELETE FROM VendedorTRVComiRed WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195956

    lErro = Comando_Executar(alComando(4), "DELETE FROM VendedorTRVRegVenda WHERE Vendedor = ?", objVendedor.iCodigo)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRV", gErr)

        Case 195955
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVCOMISSAO", gErr)

        Case 195956
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVCOMIRED", gErr)

        Case 195957
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORTRVREGVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195894)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

