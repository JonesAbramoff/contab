VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGravaPUR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function TituloReceber_Gera_Customizado(ByVal objNFiscal As ClassNFiscal, ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelas As ColParcelaReceber) As Long
    
Dim lErro As Long
Dim objParcelasRecDif As ClassParcelasRecDif
Dim objParcelasRec As ClassParcelaReceber
Dim objParcelasRecAux As ClassParcelaReceber
Dim dValorDifParc As Double
Dim dValorFalta As Double
Dim iIndice As Integer
Dim alComando(0 To 7) As Long

On Error GoTo Erro_TituloReceber_Gera_Customizado

    If gobjFAT.iAbateICMSSubstVlrCobr = MARCADO Then

        dValorFalta = objNFiscal.objTributacaoNF.dICMSSubstValor
        
        If dValorFalta > 0 Then
        
            'Abertura de Comando
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 198616
            Next
                        
            iIndice = 0
            For Each objParcelasRec In colParcelas
            
                Set objParcelasRecDif = New ClassParcelasRecDif
            
                iIndice = iIndice + 1
                
                Set objParcelasRecAux = objNFiscal.ColParcelaReceber.Item(iIndice)
            
                objParcelasRecDif.lNumIntParc = objParcelasRec.lNumIntDoc
                objParcelasRecDif.iSeq = 0
                objParcelasRecDif.dtDataRegistro = gdtDataAtual
                objParcelasRecDif.iCodTipoDif = 1

                objParcelasRecDif.dValorDiferenca = -objParcelasRecAux.dValorAux
                
                objParcelasRecDif.sObservacao = "Desconto do ICMS de substituição"
                
                lErro = CF("ParcelasRecDif_Grava_EmTrans", objParcelasRecDif, alComando)
                If lErro <> SUCESSO Then gError 198617
            
            Next
            
            'Fecha Comando
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
            
        End If
        
    End If
    
    TituloReceber_Gera_Customizado = SUCESSO

    Exit Function

Erro_TituloReceber_Gera_Customizado:

    TituloReceber_Gera_Customizado = gErr

    Select Case gErr

        Case 198616
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 198617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198618)

    End Select

    Exit Function
    
End Function

Public Function TituloReceber_Gera_Customizado2(ByVal objNFiscal As ClassNFiscal) As Long
    
Dim lErro As Long
Dim iIndice As Integer
Dim objParcelasRec As ClassParcelaReceber
Dim dValorFalta As Double
Dim dValorParcelas As Double
Dim dFatorParcela As Double
Dim dValorDif As Double

On Error GoTo Erro_TituloReceber_Gera_Customizado2

    If gobjFAT.iAbateICMSSubstVlrCobr = MARCADO Then

        dValorFalta = objNFiscal.objTributacaoNF.dICMSSubstValor
        
        If dValorFalta > 0 Then
        
            dValorParcelas = 0
            For Each objParcelasRec In objNFiscal.ColParcelaReceber
                dValorParcelas = dValorParcelas + objParcelasRec.dValor
            Next

            iIndice = 0
            For Each objParcelasRec In objNFiscal.ColParcelaReceber
                iIndice = iIndice + 1
                If iIndice = objNFiscal.ColParcelaReceber.Count Then
                    dValorDif = dValorFalta
                Else
                    dFatorParcela = objParcelasRec.dValor / dValorParcelas
                    dValorDif = Arredonda_Moeda(Fix((dFatorParcela * objNFiscal.objTributacaoNF.dICMSSubstValor) * 100) / 100)
                    dValorFalta = dValorFalta - dValorDif
                End If
                
                objParcelasRec.dValorAux = dValorDif
                objParcelasRec.dValorOriginal = objParcelasRec.dValorOriginal + dValorDif
                objParcelasRec.dValor = objParcelasRec.dValor + dValorDif
                objParcelasRec.dSaldo = objParcelasRec.dSaldo + dValorDif
                
            Next

        End If

    End If
    
    TituloReceber_Gera_Customizado2 = SUCESSO

    Exit Function

Erro_TituloReceber_Gera_Customizado2:

    TituloReceber_Gera_Customizado2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198618)

    End Select

    Exit Function
    
End Function

Public Function RelDespacho_Insere(colRegiao As Collection, lNumIntRel As Long) As Long
'Insere dados em RelDespacho
 
Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim vRegiao As Variant
Dim iSeq As Integer
Dim iRegiao As Integer
 
On Error GoTo Erro_RelDespacho_Insere
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 207096
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207097
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_PROX_REL_DESPACHO", lNumIntRel)
    If lErro <> SUCESSO Then gError 207098
    
    iSeq = 0
    
    For Each vRegiao In colRegiao
    
        iRegiao = vRegiao
    
        iSeq = iSeq + 1
    
        'Insere no BD
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelDespacho(NumIntRel, Seq, Regiao) VALUES (?,?,?)", lNumIntRel, iSeq, iRegiao)
        If lErro <> AD_SQL_SUCESSO Then gError 207099
    
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 207100
 
    RelDespacho_Insere = SUCESSO
 
    Exit Function
 
Erro_RelDespacho_Insere:
 
    RelDespacho_Insere = gErr
 
    Select Case gErr
 
        Case 207096
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 207097
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 207098
        
        Case 207099
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELBARCODE", gErr)
 
        Case 207100
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207101)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Function Vendedor_Grava_Customizacao(ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objVendedorPur As ClassVendedorPur
Dim objVendPurExc As ClassVendedorPurExc

On Error GoTo Erro_Vendedor_Grava_Customizacao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213791
    Next

    lErro = CF("Vendedor_Exclui_Customizado", objVendedor)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Not (objVendedor.objInfoUsu Is Nothing) Then
    
        Set objVendedorPur = objVendedor.objInfoUsu
   
        lErro = Comando_Executar(alComando(0), "INSERT INTO VendedoresPur (Vendedor, DataInicio, Observacao) VALUES (?,?,?)", objVendedor.iCodigo, objVendedorPur.dtDataInicio, objVendedorPur.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 195874
    
        If objVendedorPur.colExcecoes.Count > 0 Then

            iIndice = 0

            For Each objVendPurExc In objVendedorPur.colExcecoes
            
                iIndice = iIndice + 1
    
                lErro = Comando_Executar(alComando(1), "INSERT INTO VendedoresPurExc (Vendedor, Seq, Produto, PercComissao, MetaComissao) VALUES (?,?,?,?,?)", _
                         objVendedor.iCodigo, iIndice, objVendPurExc.sProduto, objVendPurExc.dPercComissao, objVendPurExc.dMetaComissao)
                If lErro <> AD_SQL_SUCESSO Then gError 195962
    
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
        
        Case 195874
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_VENDEDOR", gErr)
        
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
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_Vendedor_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 195892
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM VendedoresPur WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195893

    lErro = Comando_Executar(alComando(2), "DELETE FROM VendedoresPurExc WHERE Vendedor = ?", objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 195955
    
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

        Case 195893, 195955
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195894)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function



