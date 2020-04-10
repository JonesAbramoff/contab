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
Dim objSerie As New ClassSerie

On Error GoTo Erro_NFiscal_ObtemFatorValor

    If iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Or sSerie = "" Or ISSerieEletronica(sSerie) Then
    
        dFator = FATOR_PROPORCAO_100
        
    Else
        
        objTipoDocInfo.iCodigo = iTipoDocInfo
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 500007
        
        Select Case objTipoDocInfo.iTipoTrib
        
            Case TRIB_ENTRADA_FORN, TRIB_SAIDA_FORN
                If sSerie = "2" Then
                    dFator = FATOR_PROPORCAO_0
                Else
                    dFator = FATOR_PROPORCAO_100
                End If
            
            Case TRIB_SAIDA_CLI, TRIB_ENTRADA_CLI
                objSerie.iFilialEmpresa = iFilialEmpresa
                objSerie.sSerie = sSerie
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO Then gError 130403
                
                If sSerie <> "1" Then
                    dFator = objSerie.iMaxLinhasNF / 100
                Else
                    dFator = FATOR_PROPORCAO_100
                End If
                
        End Select
    
    End If
     
    NFiscal_ObtemFatorValor = SUCESSO
     
    Exit Function
    
Erro_NFiscal_ObtemFatorValor:

    NFiscal_ObtemFatorValor = gErr
     
    Select Case gErr
          
        Case 500007
        
        Case 130403
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA3", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179007)
     
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

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35961

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscal_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500007
    
'    End If
    'FIM FOX
    
'    lErro = CF("NFiscal_Grava_EmTrans", objNFiscal, objContabil)
'    If lErro <> SUCESSO Then gError 106557
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179008)

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

    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatura_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500008
        
'    End If
    'FIM FOX
    
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalFatura_Grava_EmTrans", objNFiscal, objContabil)
    End If
'    If lErro <> SUCESSO Then gError 106558
    
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
        If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500014
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
        
        Case 106558, 207681, 207685, 500008
        
        Case 42413, 207683
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 42438, 207684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179009)
            
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
    
    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalDevolucao_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500009
        
'    End If
    'FIM FOX
    
    If lErro1 = SUCESSO Then
        lErro2 = CF("NFiscalDevolucao_Grava_EmTrans", objNFiscal, objContabil)
    End If
'    If lErro <> SUCESSO Then gError 106559
       
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179010)
            
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
    
    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalRemessa_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500010
        
'    End If
    'FIM FOX
    
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
        
        Case 207720, 207722, 207723, 207725
        
        Case 42059, 207721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 42088, 207724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179011)
            
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

    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscal_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500013
        
'    End If
    'FIM FOX
    
'    lErro = CF("NFiscalPedido_Grava_EmTrans", objNFiscal, objContabil)
'    If lErro <> SUCESSO Then gError 106561
'
'    'FOX
'    If lNumNFOficial <> 0 And lNumNFOficial <> objNFiscal.lNumNotaFiscal Then gError 500017
'    'FIM FOX
'
'    'Faz o commit
'    lErro = Transacao_Commit()
'    If lErro <> SUCESSO Then gError 35647
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179012)
    
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

    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalRemessa_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500011
        
'    End If
    'FIM FOX
    
    lErro2 = CF("NFiscalRemPedido_Grava_EmTrans", objNFiscal, objBloqueioPV, objContabil)
'    If lErro <> SUCESSO Then gError 106562
           
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
        
        Case 207726, 207728, 207729, 207731
        
        Case 91499, 207727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 96027, 207730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179013)
    
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

    'FOX
'    If objNFiscal.lNumNotaFiscal = 0 Then
    
        lErro1 = NFiscal_Grava_Clone(objNFiscal, objContabil, "NFiscalFatura_Grava_EmTrans", lNumNFOficial)
'        If lErro <> SUCESSO Then gError 500012
        
'    End If
    'FIM FOX
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179014)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Public Function NFiscal_Lock_Gravacao(alComando() As Long, ByVal objNFiscal As ClassNFiscal) As Long
'executa o lock dos registros associados a nota fiscal.

Dim lErro As Long
Dim iIndice As Long
Dim sNome As String
Dim sDescricao As String
Dim iCodigo As Integer
Dim iFilialEmpresa As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim iAlmoxarifadoLocado As Integer
Dim objItemAloc As ClassItemNFAlocacao
Dim objComissaoNF As ClassComissaoNF
Dim iFilialEmpresa1 As Integer

On Error GoTo Erro_NFiscal_Lock_Gravacao

    lErro = CF("NaturezaOperacao_Lock", objNFiscal.sNaturezaOp, alComando(1))
    If lErro <> SUCESSO And lErro <> 40551 Then Error 35974
    
    If lErro = 40551 Then Error 35975
   
    lErro = CF("FilialCliente_Lock1", alComando(2), objNFiscal.lCliente, objNFiscal.iFilialCli)
    If lErro <> SUCESSO And lErro <> 43050 Then Error 35977
    
    'Não achou a Filial do Cliente --> erro
    If lErro = 43050 Then Error 43052

    lErro = CF("Serie_Lock", alComando(3), objNFiscal.sSerie)
    If lErro <> SUCESSO And lErro <> 43032 Then Error 35978
    
    'Não encontrou a Série --> erro
    If lErro = 43032 Then Error 43034
    
    If objNFiscal.iTabelaPreco > 0 Then
        lErro = CF("NFiscal_TabelaPreco_Lock", alComando(9), objNFiscal)
        If lErro <> SUCESSO Then Error 44292
    End If
    
    'Se a Transportadora estiver preenchida
    If objNFiscal.iCodTransportadora > 0 Then
        
        lErro = CF("Transportadora_Lock", alComando(4), objNFiscal.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 35979
        
        'Não encontrou a Transportadora --> erro
        If lErro = 43038 Then Error 43040
        
    End If
    
    'Se a Transportadora de redespacho estiver preenchida
    If objNFiscal.iCodTranspRedesp > 0 Then
        
        lErro = CF("Transportadora_Lock", alComando(4), objNFiscal.iCodTranspRedesp)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 35979
        
        'Não encontrou a Transportadora --> erro
        If lErro = 43038 Then Error 43040
        
    End If
    
    'Se o Canal de Venda estiver preenchido
    If objNFiscal.iCanal <> 0 Then
    
        lErro = CF("CanalVenda_Lock", alComando(5), objNFiscal.iCanal)
        If lErro <> SUCESSO And lErro <> 43026 Then Error 43022
        
        'Não encontrou o Canal de Venda --> erro
        If lErro = 43026 Then Error 43028
        
    End If
    
    For Each objComissaoNF In objNFiscal.ColComissoesNF
        lErro = CF("Vendedor_Lock", alComando(8), objComissaoNF.iCodVendedor)
        If lErro <> SUCESSO Then Error 51581
    Next
    
    'Lock shared nos Produtos e Almoxarifados
    For Each objItemNF In objNFiscal.ColItensNF
        
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Lock_Shared", alComando(6), objProduto)
        If lErro <> SUCESSO Then Error 35990

        objItemNF.iControleEstoque = objProduto.iControleEstoque
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.iApropriacaoProd = objProduto.iApropriacaoCusto

        If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            For Each objItemAloc In objItemNF.ColAlocacoes
    
                iAlmoxarifadoLocado = False
                
                If colCodigoNome.Count > 0 Then
                    For Each objCodigoNome In colCodigoNome
                        If objCodigoNome.iCodigo = objItemAloc.iAlmoxarifado Then
                            iAlmoxarifadoLocado = True
                        End If
                    Next
                End If
                    
                If iAlmoxarifadoLocado = False Then
                
                    colCodigoNome.Add objItemAloc.iAlmoxarifado, objItemAloc.sAlmoxarifado
                
                    If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
                    
                    If objNFiscal.iFilialEmpresa = 52 Then
                        iFilialEmpresa1 = 51
                    ElseIf objNFiscal.iFilialEmpresa = 2 Then
                        iFilialEmpresa1 = 1
                    Else
                        iFilialEmpresa1 = objNFiscal.iFilialEmpresa
                    End If
                    
                    lErro = Comando_ExecutarLockado(alComando(7), "SELECT FilialEmpresa FROM Almoxarifado WHERE Codigo = ? AND FilialEmpresa = ? ", iFilialEmpresa, objItemAloc.iAlmoxarifado, iFilialEmpresa1)
                    If lErro <> AD_SQL_SUCESSO Then Error 35991
        
                    lErro = Comando_BuscarPrimeiro(alComando(7))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35994
                    If lErro = AD_SQL_SEM_DADOS Then Error 35992
        
                    lErro = Comando_LockShared(alComando(7))
                    If lErro <> AD_SQL_SUCESSO Then Error 35993
            
                End If
            Next
        End If
    Next

    NFiscal_Lock_Gravacao = SUCESSO

    Exit Function

Erro_NFiscal_Lock_Gravacao:

    NFiscal_Lock_Gravacao = Err

    Select Case Err

        Case 35974, 35975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", Err, objNFiscal.sNaturezaOp)

        Case 35976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objNFiscal.sNaturezaOp)

        Case 35977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", Err, objNFiscal.sNaturezaOp)

        Case 35979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES1", Err, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case 35990, 43022, 35978, 44292, 51581

        Case 35991, 35994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", Err, objItemNF.iAlmoxarifado)

        Case 35992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objItemAloc.iAlmoxarifado)

        Case 35993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO1", Err, objItemNF.iAlmoxarifado)

        Case 43028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_CADASTRADO", Err, objNFiscal.iCanal)

        Case 43034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objNFiscal.sSerie)

        Case 43040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case 43052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179015)

    End Select

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
    If lComando = 0 Then gError 60436

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 60437

    lErro = NFiscal_ObtemFatorValor(objNFiscal.iFilialEmpresa, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, dFator)
    If lErro <> SUCESSO Then gError 60438
    
    If dFator <> FATOR_PROPORCAO_100 And dFator <> FATOR_PROPORCAO_0 Then
        sSerieUnica = "1"
    Else
        sSerieUnica = objNFiscal.sSerie
    End If
    
    lErro = Comando_ExecutarPos(lComando, "SELECT ProxNumNFiscal FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, tSerie.lProxNumNFiscal, sSerieUnica, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 60438

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60439
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60440
    
    objNFiscal.lNumNotaFiscal = tSerie.lProxNumNFiscal
    
    tSerie.lProxNumNFiscal = tSerie.lProxNumNFiscal + 1
    
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Serie SET ProxNumNFiscal = ?", lComando, tSerie.lProxNumNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 60441
    
    lErro = CF("NFiscal_Testa_Numero", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    NFiscalNumAuto = SUCESSO
    
    Exit Function
    
Erro_NFiscalNumAuto:

    NFiscalNumAuto = gErr
    
    Select Case gErr
    
        Case 60436, 60437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 60438, 60439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", gErr, objNFiscal.sSerie)

        Case 60440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", gErr)
        
        Case 60441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", gErr, objNFiscal.sSerie)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179016)
            
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179017)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179018)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Serie_Grava_Trans_Aux(ByVal objSerie As ClassSerie, ByVal iMaxLinhasNF As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Serie_Grava_Trans_Aux

    'sport
    If objSerie.iMaxLinhasNF <> iMaxLinhasNF And gsUsuario <> "supervisor" Then Error 32306
        
    Serie_Grava_Trans_Aux = SUCESSO

    Exit Function

Erro_Serie_Grava_Trans_Aux:

    Serie_Grava_Trans_Aux = Err

    Select Case Err

        Case 32306 'sport
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_MAX_LINHAS_NF", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179019)

    End Select

    Exit Function

End Function

Function EstoqueInicial_Grava(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal iAlmoxarifadoPadrao As Integer, ByVal colRastreamento As Collection) As Long
'Grava o estoque inicial do produto

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_EstoqueInicial_Grava

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 180582
    
    lErro = EstoqueInicial_Grava_Clone(objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
    If lErro <> SUCESSO Then gError 180583

    'grava o estoque inicial em transacao
    lErro = CF("EstoqueInicial_Grava1", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
    If lErro <> SUCESSO Then gError 180584
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 180585

    EstoqueInicial_Grava = SUCESSO

    Exit Function

Erro_EstoqueInicial_Grava:

    EstoqueInicial_Grava = gErr

    Select Case gErr

        Case 180582
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 180583, 180584

        Case 180585
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180586)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function EstoqueInicial_Exclui(ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'Exclui o a associacao de estoque (almoxarifado) com produto passados como parametro

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_EstoqueInicial_Exclui

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181178

    lErro = EstoqueInicial_Exclui_Clone(objEstoqueProduto)
    If lErro <> SUCESSO Then gError 181179

    'continuacao da exclusao do estoqueinicial
    lErro = CF("EstoqueInicial_Exclui0", objEstoqueProduto)
    If lErro <> SUCESSO Then gError 181180

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181181

    EstoqueInicial_Exclui = SUCESSO

    Exit Function

Erro_EstoqueInicial_Exclui:

    EstoqueInicial_Exclui = gErr

    Select Case gErr
    
        Case 181178
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 181179 To 181180

        Case 181181
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181182)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function
