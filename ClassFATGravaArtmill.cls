VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFiscal_ObtemFatorValor(ByVal iFilialEmpresa As Integer, ByVal iTipoDocInfo As Integer, ByVal sSerie As String, dFator As Double) As Long

Dim lErro As Long, objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscal_ObtemFatorValor

    If iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Or sSerie = "" Then
    
        dFator = FATOR_PROPORCAO_100
        
    Else
        
        objTipoDocInfo.iCodigo = iTipoDocInfo
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 500007
        
        Select Case objTipoDocInfo.iTipoTrib
        
            Case TRIB_ENTRADA_FORN, TRIB_SAIDA_FORN
                dFator = FATOR_PROPORCAO_100
            
            Case TRIB_SAIDA_CLI, TRIB_ENTRADA_CLI
                Select Case sSerie
                
                    Case "1", "1-e"
                        dFator = FATOR_PROPORCAO_100
                    
                    Case "2", "2-e"
                        dFator = FATOR_PROPORCAO_50
                                        
                    Case "3", "3-e"
                        dFator = FATOR_PROPORCAO_25
                                        
                    Case "4", "4-e"
                        dFator = FATOR_PROPORCAO_75
                                        
                    Case Else
                        dFator = FATOR_PROPORCAO_0
                    
                End Select
                
        End Select
    
    End If
     
    NFiscal_ObtemFatorValor = SUCESSO
     
    Exit Function
    
Erro_NFiscal_ObtemFatorValor:

    NFiscal_ObtemFatorValor = gErr
     
    Select Case gErr
          
        Case 500007
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150032)
     
    End Select
     
    Exit Function

End Function

Function NFiscal_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'grava uma nota fiscal

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscal_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35961

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscal_Grava_EmTrans", lNumNFOficial)
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscal_Grava_EmTrans", objNFiscal, objContabil)
    End If
    
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207691

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207692
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalFatura_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207693
        
            lErro = CF("NFiscalFatura_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207694
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207695
    
        Else
            gError 207696
    
        End If
    
    Else
    
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500020
        'FIM FOX
        
        'Confirma Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 39004

    End If

    NFiscal_Grava = SUCESSO

    Exit Function

Erro_NFiscal_Grava:

    NFiscal_Grava = gErr

    Select Case gErr

        Case 500020
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 35961, 207692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 39004, 207695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case 106557, 500007, 207691, 207693, 207694, 207696
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150033)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatura_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Grava NFiscal Fatura

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalFatura_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Inicia Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 42413
    
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatura_Grava_EmTrans", lNumNFOficial)
        
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalFatura_Grava_EmTrans", objNFiscal, objContabil)
    End If
    
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207681

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207683
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalFatura_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 500008
        
            lErro = CF("NFiscalFatura_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 106558
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207684
    
        Else
            gError 207685
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Or lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then gError 500014
        'FIM FOX
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 42438

    End If


    NFiscalFatura_Grava = SUCESSO

    Exit Function
    
Erro_NFiscalFatura_Grava:

    NFiscalFatura_Grava = gErr
    
    Select Case gErr
    
        Case 500014
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 106558, 500008, 207681, 207685
        
        Case 42413, 207683
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 42438, 207684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150034)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalDevolucao_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Grava uma NFiscal Devolução

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalDevolucao_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Inicia transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 39246
    
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalDevolucao_Grava_EmTrans", lNumNFOficial)
        
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalDevolucao_Grava_EmTrans", objNFiscal, objContabil)
    End If
       
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207770

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207771
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalRemessa_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207772
        
            lErro = CF("NFiscalRemessa_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207773
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207774
    
        Else
            gError 207775
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500015
        'FIM FOX
        
        'Confirma transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 39275
    
    End If
    
    
    NFiscalDevolucao_Grava = SUCESSO

    Exit Function
    
Erro_NFiscalDevolucao_Grava:

    NFiscalDevolucao_Grava = gErr
    
    Select Case gErr
    
        Case 500015
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 106559, 500009, 207770, 207772, 207773, 207775
    
        Case 39246, 207771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 39275, 207774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150035)
            
    End Select
    
    'Finaliza transação
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalRemessa_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long


On Error GoTo Erro_NFiscalRemessa_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 42059
    
    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalRemessa_Grava_EmTrans", lNumNFOficial)
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalRemessa_Grava_EmTrans", objNFiscal, objContabil)
    End If
    
    
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207720

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207721
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalRemessa_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207722
        
            lErro = CF("NFiscalRemessa_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207723
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207724
    
        Else
            gError 207725
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500016
        'FIM FOX
        
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 42088

    End If
    
    
    NFiscalRemessa_Grava = SUCESSO

    Exit Function
    
Erro_NFiscalRemessa_Grava:

    NFiscalRemessa_Grava = gErr
    
    Select Case gErr
    
        Case 500016
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 106560, 500010, 207720, 207722, 207723, 107725
        
        Case 42059, 207721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 42088, 207724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150036)
            
    End Select
    
    Call Transacao_Rollback
   
    Exit Function
    
End Function

Function NFiscalPedido_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal gerada a partir de um pedido de venda

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalPedido_Grava
   
    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

   'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 46598

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscal_Grava_EmTrans", lNumNFOficial)
        
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalPedido_Grava_EmTrans", objNFiscal, objContabil)
    End If
'    If lErro <> SUCESSO Then gError 106561
    
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207753

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207754
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalPedido_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207755
        
            lErro = CF("NFiscalPedido_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207756
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207757
    
        Else
            gError 207758
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500017
        'FIM FOX
        
        'Faz o commit
        lErro = Transacao_Commit()
        If lErro <> SUCESSO Then gError 35647

    End If
    
    
    NFiscalPedido_Grava = SUCESSO
    
    Exit Function
    
Erro_NFiscalPedido_Grava:

    NFiscalPedido_Grava = gErr
    
    Select Case gErr
    
        Case 500017
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 106561, 500013, 207753, 207755, 207756, 207758
        
        Case 35647, 207757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                    
        Case 46598, 207754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150037)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function NFiscalRemPedido_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objBloqueioPV As ClassBloqueioPV, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal remessa gerada a partir de um pedido de venda

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalRemPedido_Grava
   
    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

   'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 91499

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalRemessa_Grava_EmTrans", lNumNFOficial)
        
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalRemPedido_Grava_EmTrans", objNFiscal, objBloqueioPV, objContabil)
    End If
           
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207726

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207727
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalRemessa_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207728
        
            lErro = CF("NFiscalRemessa_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207729
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207730
    
        Else
            gError 207731
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500018
        'FIM FOX
        
        'Faz o commit
        lErro = Transacao_Commit()
        If lErro <> SUCESSO Then gError 96027
    
    End If
    
    
    NFiscalRemPedido_Grava = SUCESSO
    
    Exit Function
    
Erro_NFiscalRemPedido_Grava:

    NFiscalRemPedido_Grava = gErr
    
    Select Case gErr
      
        Case 500018
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 106562, 500011, 207726, 207728, 207729, 207731
        
        Case 91499, 207727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 96027, 207730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150038)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function NFiscalFaturaPedido_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal fatura gerada a partir de um pedido de venda

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNFOficial As Long
Dim lErro1 As Long
Dim lErro2 As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumNotaFiscal As Long

On Error GoTo Erro_NFiscalFaturaPedido_Grava
   
    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

   'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 46598

    lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatura_Grava_EmTrans", lNumNFOficial)
        
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalFaturaPedido_Grava_EmTrans", objNFiscal, objContabil)
    End If
               
    If lErro1 <> SUCESSO Or lErro2 <> SUCESSO Then
    
        'se era uma alteracao e deu erro na exclusao;insercao ==> tenta somente gravar os dados de contabilidade
        If lNumNotaFiscal <> 0 Then
    
            'Dá a Mensagem de aviso que só será gravado a contabilidade
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            If vbMsgRes = vbNo Then gError 207760

            Call Transacao_Rollback
    
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 207761
    
            lErro = NFiscal_Altera_Clone(objNFiscal, objContabil, "NFiscalPedido_AlteraDados_EmTrans", lNumNFOficial)
            If lErro <> SUCESSO Then gError 207762
        
            lErro = CF("NFiscalPedido_AlteraDados_EmTrans", objNFiscal, objContabil)
            If lErro <> SUCESSO Then gError 207763
    
            'Faz o Commit da Transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 207764
    
        Else
            gError 207765
    
        End If
    
    Else
    
        'FOX
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500019
        'FIM FOX
        
        'Faz o commit
        lErro = Transacao_Commit()
        If lErro <> SUCESSO Then gError 35647
    
    End If
    
    
    NFiscalFaturaPedido_Grava = SUCESSO
    
    Exit Function
    
Erro_NFiscalFaturaPedido_Grava:

    NFiscalFaturaPedido_Grava = gErr
    
    Select Case gErr
    
        Case 500019
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNFCLONE_DIFERE_NUMNFORIG", gErr)
        
        Case 106563, 500012, 207760, 207762, 207763, 207765
        
        Case 35647, 207764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 46598, 207761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150039)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalNumAuto(ByVal objNFiscal As ClassNFiscal) As Long
'Lê o Proximo número na tabela de Série e Coloca no objNFiscal
'Faz Lock Exclusive e atualiza o Número na Tabela de Série

Dim lErro As Long
Dim tSerie As typeSerie
Dim lComando As Long, dFator As Double
Dim lComando1 As Long, sSerieUnica As String

On Error GoTo Erro_NFiscalNumAuto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60436

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60437

    lErro = NFiscal_ObtemFatorValor(objNFiscal.iFilialEmpresa, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, dFator)
    If lErro <> SUCESSO Then Error 60438
    
    If dFator <> FATOR_PROPORCAO_100 And dFator <> FATOR_PROPORCAO_0 Then
        If ISSerieEletronica(objNFiscal.sSerie) Then
            sSerieUnica = "1-e"
        Else
            sSerieUnica = "1"
        End If
    Else
        sSerieUnica = objNFiscal.sSerie
    End If
    
    lErro = Comando_ExecutarPos(lComando, "SELECT ProxNumNFiscal FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, tSerie.lProxNumNFiscal, sSerieUnica, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 60438

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 60439
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 60440
    
    objNFiscal.lNumNotaFiscal = tSerie.lProxNumNFiscal
    
    tSerie.lProxNumNFiscal = tSerie.lProxNumNFiscal + 1
    
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Serie SET ProxNumNFiscal = ?", lComando, tSerie.lProxNumNFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 60441
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    NFiscalNumAuto = SUCESSO
    
    Exit Function
    
Erro_NFiscalNumAuto:

    NFiscalNumAuto = Err
    
    Select Case Err
    
        Case 60436, 60437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60438, 60439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", Err, objNFiscal.sSerie)

        Case 60440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", Err)
        
        Case 60441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", Err, objNFiscal.sSerie)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150041)
            
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Function NotaFiscalSaida_Excluir(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a exclusao de uma nota fiscal de Saida

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalSaida_Excluir
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 126978

    'exclui o clone se houver
    lErro = NFiscal_Exclui_Clone(objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 126969

    lErro = CF("NotaFiscalSaida_Excluir_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 126968
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 126979
    
    NotaFiscalSaida_Excluir = SUCESSO

    Exit Function

Erro_NotaFiscalSaida_Excluir:

    NotaFiscalSaida_Excluir = gErr

    Select Case gErr

        Case 126968, 126969
        
        Case 126978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 126979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150042)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalSaida_Cancelar(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Saida

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalSaida_Cancelar
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 126980

    'exclui o clone se houver
    lErro = NFiscal_Cancela_Clone(objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 126981

    lErro = CF("NotaFiscalSaida_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 126982
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 126983
    
    NotaFiscalSaida_Cancelar = SUCESSO

    Exit Function

Erro_NotaFiscalSaida_Cancelar:

    NotaFiscalSaida_Cancelar = gErr

    Select Case gErr

        Case 126980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 126981, 126982
        
        Case 126983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150043)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalSaida_Cancelar_EmTrans1(ByVal objNFiscalOficial As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Saida

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_NotaFiscalSaida_Cancelar_EmTrans1
    
    objNFiscal.lNumIntDoc = objNFiscalOficial.lNumIntDoc + 1

    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 210483
    
    If lErro <> SUCESSO Then gError 210484
    
    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 210485
    
    objNFiscal.sMotivoCancel = objNFiscalOficial.sMotivoCancel
    
    'exclui o clone se houver
    lErro = NFiscal_Cancela_Clone(objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210486
    
    lErro = CF("NotaFiscalSaida_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210487
    
    NotaFiscalSaida_Cancelar_EmTrans1 = SUCESSO

    Exit Function

Erro_NotaFiscalSaida_Cancelar_EmTrans1:

    NotaFiscalSaida_Cancelar_EmTrans1 = gErr

    Select Case gErr

        Case 210483, 210485 To 210487
        
        Case 210484
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210481)

    End Select

    Exit Function

End Function


Function NotaFiscalRemPedidoSaida_Excluir(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a exclusao de uma nota fiscal de Saida

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalRemPedidoSaida_Excluir
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 210982

    'exclui o clone se houver
    lErro = NFiscal_Exclui_Clone(objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 210983

    lErro = CF("NotaFiscalRemPedidoSaida_Excluir_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 210984
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 210985
    
    NotaFiscalRemPedidoSaida_Excluir = SUCESSO

    Exit Function

Erro_NotaFiscalRemPedidoSaida_Excluir:

    NotaFiscalRemPedidoSaida_Excluir = gErr

    Select Case gErr

        Case 210983, 210984
        
        Case 210982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 210985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210986)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

