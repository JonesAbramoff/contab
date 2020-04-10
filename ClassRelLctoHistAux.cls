VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelLctoHistAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gcolTransOrig As New Collection
Dim alComando(1 To 13) As Long

Private Const STRING_HISTAUX_BX_CANC_TIT_EXC = "Baixa Cancelada de Título excluído"

Function Carrega_TransOrig(colTransOrig As Collection) As Long
'Lê a Transação e o Codigo da Transacao

Dim lErro As Long
Dim lComando As Long

Dim iTransacao As Integer, iOrigem As Integer

On Error GoTo Erro_Carrega_TransOrig

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 93582
    
    'carrega todas as transacoes
    lErro = Comando_Executar(lComando, "SELECT DISTINCT TransacaoCTB.Codigo, TransacaoCTBOrigem.Codigo FROM TransacaoCTB, TransacaoCTBOrigem WHERE TransacaoCTB.OrigemLcto=TransacaoCTBOrigem.Codigo", iTransacao, iOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 93583

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93584
    
    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then gError 93585
    
    'Guarda na coleção os dados
    Do While lErro = AD_SQL_SUCESSO

        Call colTransOrig.Add(CInt(iOrigem), "X" & CStr(iTransacao))

        'Busca o próximo registro que será armazenado em colCondicoesdePagamento
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93586

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)
       
    Carrega_TransOrig = SUCESSO

    Exit Function

Erro_Carrega_TransOrig:

    Carrega_TransOrig = gErr

    Select Case gErr

        Case 93582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93583 To 93584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOCTB", gErr)
        
        Case 93585
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153189)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Lcto_GeraHistoricoAux(sHistorico As String, ByVal lNumIntDoc As Long, ByVal iTransacao As Integer) As Long
'retorna em sHistorico um historico automatico a partir do numintdoc e transacao do lcto contabil informados

Dim lErro As Long
Dim lNumNotaFiscal As Long
Dim lNumTitulo As Long
Dim lNumIntBaixa As Long
Dim iIndice As Integer
Dim iNumParcela As Integer, lNumMovto As Long

On Error GoTo Erro_Lcto_GeraHistoricoAux

    sHistorico = ""
    
    If iTransacao <> 0 Then
    
        If gcolTransOrig.Count = 0 Then
        
            lErro = Carrega_TransOrig(gcolTransOrig)
            If lErro <> SUCESSO Then gError 93587
            
            'abrir comandos
            For iIndice = LBound(alComando) To UBound(alComando)
            
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 93588
            
            Next
            
        End If
    
        '??? verificar se é outro doc
        
        Select Case gcolTransOrig.Item("X" & CStr(iTransacao))
        
            Case 2  'BaixasPag
                'retornar: "Baixa do nnnn/nn , nnnn/nn ,..." onde nnnn/nn é titulo/parcela
                
                lErro = Comando_Executar(alComando(1), "SELECT TitulosPag.NumTitulo, ParcelasPag.NumParcela, BaixasPag.NumMovCta FROM BaixasPag, BaixasParcPag, ParcelasPag, TitulosPag WHERE BaixasPag.NumIntBaixa=BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela=ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo=TitulosPag.NumIntDoc AND BaixasPag.NumIntBaixa =?", lNumTitulo, iNumParcela, lNumMovto, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93616
    
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93617
        
                Do While lErro = AD_SQL_SUCESSO
                    
                    If sHistorico = "" Then
                        sHistorico = "Baixa do Titulo "
                    Else
                        sHistorico = sHistorico & " , "
                    End If
                    sHistorico = sHistorico & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93624
    
                Loop
    
                'pesquisa em titulos baixados
                
                lErro = Comando_Executar(alComando(2), "SELECT TitulosPagBaixados.NumTitulo, ParcelasPagBaixadas.NumParcela, BaixasPag.NumMovCta FROM BaixasPag, BaixasParcPag, ParcelasPagBaixadas, TitulosPagBaixados WHERE BaixasPag.NumIntBaixa=BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela=ParcelasPagBaixadas.NumIntDoc AND ParcelasPagBaixadas.NumIntTitulo=TitulosPagBaixados.NumIntDoc AND BaixasPag.NumIntBaixa =?", lNumTitulo, iNumParcela, lNumMovto, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93625
    
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93626
                
                Do While lErro = AD_SQL_SUCESSO
                
                    If sHistorico = "" Then
                        sHistorico = "Baixa do Titulo "
                    Else
                        sHistorico = sHistorico & " , "
                    End If
                    sHistorico = sHistorico & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                        
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93627
    
                Loop
                           
                'Se não encontrou nenhum registro
                If sHistorico = "" Then sHistorico = STRING_HISTAUX_BX_CANC_TIT_EXC
            
                'se a baixa está associada a um movto de cta corrente
                If lNumMovto <> 0 Then
                
                    lErro = Lcto_GeraHistoricoAux2(sHistorico, lNumMovto, alComando(13))
                    If lErro <> SUCESSO Then gError 81714
                
                End If
                
            Case 3   'BaixasParcPag
                'retornar: "Baixa do Titulo/Parcela"
                
                'pesquisa nos titulos baixados
                lErro = Comando_Executar(alComando(3), "SELECT TitulosPagBaixados.NumTitulo, ParcelasPagBaixadas.NumParcela, BaixasPag.NumMovCta FROM BaixasPag, BaixasParcPag, ParcelasPagBaixadas, TitulosPagBaixados WHERE BaixasPag.NumIntBaixa=BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela=ParcelasPagBaixadas.NumIntDoc AND ParcelasPagBaixadas.NumIntTitulo=TitulosPagBaixados.NumIntDoc AND BaixasParcPag.NumIntDoc =?", lNumTitulo, iNumParcela, lNumMovto, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93613
    
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93614
                
                If lErro = AD_SQL_SUCESSO Then
                
                    sHistorico = "Baixa do Titulo " & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                
                Else
                
                    'pesquisa nos titulos abertos
                    lErro = Comando_Executar(alComando(4), "SELECT TitulosPag.NumTitulo, ParcelasPag.NumParcela, BaixasPag.NumMovCta FROM BaixasPag, BaixasParcPag, ParcelasPag, TitulosPag WHERE BaixasPag.NumIntBaixa=BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela=ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo=TitulosPag.NumIntDoc AND BaixasParcPag.NumIntDoc =?", lNumTitulo, iNumParcela, lNumMovto, lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 93628
    
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93629
                
                    'Se encontrou
                    If lErro = AD_SQL_SUCESSO Then sHistorico = "Baixa do Titulo " & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                    
                End If
                    
                'Se não encontrou nenhum registro
                If sHistorico = "" Then sHistorico = STRING_HISTAUX_BX_CANC_TIT_EXC
                
                'se a baixa está associada a um movto de cta corrente
                If lNumMovto <> 0 Then
                
                    lErro = Lcto_GeraHistoricoAux2(sHistorico, lNumMovto, alComando(13))
                    If lErro <> SUCESSO Then gError 81715
                
                End If
                
            Case 4   'BaixasParcRec
                'retornar: "Baixa do Titulo/Parcela"
                
                'pesquisa nos titulos baixados
                lErro = Comando_Executar(alComando(5), "SELECT TitulosRecBaixados.NumTitulo, ParcelasRecBaixadas.NumParcela FROM BaixasParcRec, ParcelasRecBaixadas, TitulosRecBaixados WHERE BaixasParcRec.NumIntParcela = ParcelasRecBaixadas.NumIntDoc AND ParcelasRecBaixadas.NumIntTitulo = TitulosRecBaixados.NumIntDoc AND BaixasParcRec.NumIntDoc =?", lNumTitulo, iNumParcela, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93610
    
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93611
    
                If lErro = AD_SQL_SUCESSO Then
                    
                    sHistorico = "Baixa do Titulo " & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                
                Else
                
                    'pesquisa nos titulos abertos
                    lErro = Comando_Executar(alComando(6), "SELECT TitulosRec.NumTitulo, ParcelasRec.NumParcela FROM BaixasParcRec, ParcelasRec, TitulosRec WHERE BaixasParcRec.NumIntParcela = ParcelasRec.NumIntDoc AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND BaixasParcRec.NumIntDoc =?", lNumTitulo, iNumParcela, lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 93591
    
                    lErro = Comando_BuscarPrimeiro(alComando(6))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93592
                                     
                    'Se encontrou
                    If lErro = AD_SQL_SUCESSO Then sHistorico = "Baixa do Titulo " & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                                      
                End If
                                   
                'Se não encontrou nenhum registro
                If sHistorico = "" Then sHistorico = STRING_HISTAUX_BX_CANC_TIT_EXC
                           
            Case 5   'BaixasRec
                'retornar: "Baixa do nnnn/nn , nnnn/nn ,..." onde nnnn/nn é titulo/parcela
                
                'pesquisa em titulos abertos
                lErro = Comando_Executar(alComando(7), "SELECT TitulosRec.NumTitulo, ParcelasRec.NumParcela FROM BaixasRec, BaixasParcRec, ParcelasRec, TitulosRec WHERE BaixasRec.NumIntBaixa = BaixasParcRec.NumIntBaixa AND BaixasParcRec.NumIntParcela = ParcelasRec.NumIntDoc AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND BaixasRec.NumIntBaixa =?", lNumTitulo, iNumParcela, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93608
    
                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93620
                                     
                Do While lErro = AD_SQL_SUCESSO
                    
                    If sHistorico = "" Then
                        sHistorico = "Baixa do Titulo "
                    Else
                        sHistorico = sHistorico & " , "
                    End If
                    sHistorico = sHistorico & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                
                    lErro = Comando_BuscarProximo(alComando(7))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93621
    
                Loop
                
                'pesquisa em titulos baixados
               
                lErro = Comando_Executar(alComando(8), "SELECT TitulosRecBaixados.NumTitulo, ParcelasRecBaixadas.NumParcela FROM BaixasRec, BaixasParcRec, ParcelasRecBaixadas, TitulosRecBaixados WHERE BaixasRec.NumIntBaixa = BaixasParcRec.NumIntBaixa AND BaixasParcRec.NumIntParcela = ParcelasRecBaixadas.NumIntDoc AND ParcelasRecBaixadas.NumIntTitulo = TitulosRecBaixados.NumIntDoc AND BaixasRec.NumIntBaixa =?", lNumTitulo, iNumParcela, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93622
    
                lErro = Comando_BuscarPrimeiro(alComando(8))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93623
    
                Do While lErro = AD_SQL_SUCESSO
                
                    If sHistorico = "" Then
                        sHistorico = "Baixa do Titulo "
                    Else
                        sHistorico = sHistorico & " , "
                    End If
                    sHistorico = sHistorico & CStr(lNumTitulo) & "/" & CStr(iNumParcela)
                        
                    lErro = Comando_BuscarProximo(alComando(8))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93589
    
                Loop
                                             
                'Se não encontrou nenhum registro
                If sHistorico = "" Then sHistorico = STRING_HISTAUX_BX_CANC_TIT_EXC
                         
            Case 10  'NFiscal
                'retornar: "Nota Fiscal No. nnnn"
                
                lErro = Comando_Executar(alComando(9), "SELECT NumNotaFiscal FROM NFiscal WHERE NumIntDoc =?", lNumNotaFiscal, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93596
    
                lErro = Comando_BuscarPrimeiro(alComando(9))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93597
        
                'Se não encontrou
                If lErro = AD_SQL_SEM_DADOS Then
                    sHistorico = "Nota Fiscal excluída"
                Else
                    sHistorico = "Nota Fiscal No. " & CStr(lNumNotaFiscal)
                End If
            
            Case 11  'NFsPag
                'retornar: "Nota Fiscal No. nnnn"
                            
                lErro = Comando_Executar(alComando(10), "SELECT NumNotaFiscal FROM nfspag_todas WHERE NumIntDoc =?", lNumNotaFiscal, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93599
    
                lErro = Comando_BuscarPrimeiro(alComando(10))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93600
        
                'Se não encontrou
                If lErro = AD_SQL_SEM_DADOS Then
                    sHistorico = "Nota Fiscal excluída"
                Else
                    sHistorico = "Nota Fiscal No. " & CStr(lNumNotaFiscal)
                End If
                                
            Case 15  'TitulosPag
                'retornar: "Titulo No. nnnn"
                            
                lErro = Comando_Executar(alComando(11), "SELECT NumTitulo FROM titulospag_todos WHERE NumIntDoc =?", lNumTitulo, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93602
    
                lErro = Comando_BuscarPrimeiro(alComando(11))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93603
        
                'Se não encontrou
                If lErro = AD_SQL_SEM_DADOS Then
                    sHistorico = "Título excluído"
                Else
                    sHistorico = "Título No. " & CStr(lNumTitulo)
                End If
                            
            Case 16  'TitulosRec
                'retornar: "Titulo No. nnnn"
                            
                lErro = Comando_Executar(alComando(12), "SELECT NumTitulo FROM titulosrectodos WHERE NumIntDoc =?", lNumTitulo, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 93605
    
                lErro = Comando_BuscarPrimeiro(alComando(12))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93606
        
                'Se não encontrou
                If lErro = AD_SQL_SEM_DADOS Then
                    sHistorico = "Título excluído"
                Else
                    sHistorico = "Título No. " & CStr(lNumTitulo)
                End If
            
            Case Else
            
                sHistorico = ""
        
        End Select
        
    End If
        
    Lcto_GeraHistoricoAux = SUCESSO
     
    Exit Function
    
Erro_Lcto_GeraHistoricoAux:

    Lcto_GeraHistoricoAux = gErr
     
    Select Case gErr
        
        Case 81714, 81715
        
        Case 93588
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 93596, 93597
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTA_FISCAL", gErr)
                   
        Case 93598
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA ", gErr)
                                 
        Case 93599, 93600
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)
                   
        Case 93601
            Call Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA ", gErr)
                        
        Case 93602, 93603
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG", gErr)
                   
        Case 93604
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOPAGAR_NAO_CADASTRADO", gErr)
                         
        Case 93605, 93606
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", gErr)
                   
        Case 93607
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_RECEBER_NAO_CADASTRADO", gErr)
                                 
        Case 93608, 93620, 93621, 93622, 93623, 93589
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXAREC", gErr)
                   
        Case 93609
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAREC_NAO_ENCONTRADA", gErr)
                                 
        Case 93610, 93611, 93591, 93592, 93593
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXAPARCREC", gErr)
                   
        Case 93612
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAPARCREC_INEXISTENTE", gErr)
                                 
        Case 93613, 93614, 93628, 93629
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXAPARCPAG", gErr)
                   
        Case 93615
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAPARCPAG_INEXISTENTE", gErr)
                                 
        Case 93616, 93617, 93624, 93625, 93626, 93627
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXAPAG", gErr)
                   
        Case 93618
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAPAG_INEXISTENTE", gErr)
                                        
        Case 93587
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_VAZIA", gErr)
             
        Case 93630
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_LCTO_HIST_AUX", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153190)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Terminate()

Dim iIndice As Integer

    'liberar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next
    
    Set gcolTransOrig = Nothing
                    
End Sub

Private Function Lcto_GeraHistoricoAux2(sHistorico As String, lNumMovto As Long, lComando As Long) As Long
'acrescenta ao historico os dados do meio de pagto

Dim lErro As Long, tMovCta As typeMovContaCorrente

On Error GoTo Erro_Lcto_GeraHistoricoAux2

    lErro = Comando_Executar(lComando, "SELECT TipoMeioPagto, Numero FROM MovimentosContaCorrente WHERE NumMovto = ?", _
        tMovCta.iTipoMeioPagto, tMovCta.lNumero, lNumMovto)
    If lErro <> AD_SQL_SUCESSO Then gError 81712
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 81713
    
    Select Case tMovCta.iTipoMeioPagto
    
        Case Cheque
            sHistorico = sHistorico & " cheque " & CStr(tMovCta.lNumero)

        Case BORDERO
            sHistorico = sHistorico & " bordero " & CStr(tMovCta.lNumero)

    End Select
    
    Lcto_GeraHistoricoAux2 = SUCESSO
     
    Exit Function
    
Erro_Lcto_GeraHistoricoAux2:

    Lcto_GeraHistoricoAux2 = gErr
     
    Select Case gErr
          
        Case 81712, 81713
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCONTACORRENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153191)
     
    End Select
     
    Exit Function

End Function
