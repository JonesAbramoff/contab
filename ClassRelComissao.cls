VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'esta classe contem funcoes auxiliares à execucao de relatorios envolvendo comissoes sobre vendas

'iTipoTitulo e lNumIntDoc identificam o docto associado a comissao
'iTipoTitulo pode ser: TIPO_NF, TIPO_PARCELA, TIPO_DEBITO, TIPO_TITULO_RECEBER
'lNumIntDoc tem o significado dependente de iTipoTitulo

'Variaveis Globais
Private lComandoNF As Long
Private lComandoParc As Long
Private lComandoDeb As Long
Private lComandoTit As Long
Private lComandoNumPVTit As Long
Private lComandoCli As Long 'Inserido por Wagner

'última "chave" lida
Private lNumIntDocUlt As Long
Private iTipoTituloUlt As Integer
'artibutos correspondentes à última "chave" lida
Private lNumeroNotaFiscalUlt As Long
Private lNumeroTituloUlt As Long
Private iNumeroParcelaUlt As Integer
Private sSiglaDocumentoUlt As String
Private dtDataUlt As Date
Private sNomeRedCliUlt As String
Private lNumeroPVUlt As Long
Private sNumeroPVCliUlt As String
Private sRazaoSocialUlt As String 'Inserido por Wagner

Private lNumIntDocUltAux As Long
Private iTipoTituloUltAux As Integer
Private sTextoDataVencUltAux As String
Private iEmBorderoUltAux As Integer
Private sTextoBaixaUltAux As String

Function Comissao_Obter_NomeRedCli(iTipoTitulo As Integer, lNumIntDoc As Long, sNomeRedCli As String, dtDataGeracao As Date) As Long
'obtem o nome reduzido do cliente

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_NomeRedCli

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then Error 59099
   
   End If

    sNomeRedCli = sNomeRedCliUlt

    Comissao_Obter_NomeRedCli = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_NomeRedCli:
    
    Comissao_Obter_NomeRedCli = Err
    
    Select Case Err
    
        Case 59099
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153056)

    End Select
        
    Exit Function
    
End Function

Function Comissao_Obter_Doc(iTipoTitulo As Integer, lNumIntDoc As Long, sDoc As String, dtDataGeracao As Date) As Long
'obtem texto p/identificar comissao
'p/nffat: "NFFAT xxxxxx" (xxxxxx numero da nf)
'para titulo: SiglaDocumento xxxxxx (xxxxxx numero do titulo)
'para parc: SiglaDocumento xxxxxx/yy (xxxxxx numero do titulo e yy o # da parcela)
'obs: para os tres acima, ver se o docto está vinculado a nf e esta por sua vez a pv
    'se estiver, concatenar " (PV xxxxxxx)" (xxxxxx numero do PV)
'p/debitorec: SiglaDocumento na tabela DebitosRecCli seguido do NumTitulo

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_Doc

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then
        
        'Obtem os Dados da comissao
        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then Error 59100
    
    End If
    
    'Verifica qual é a comissao
    Select Case iTipoTitulo
        
        Case TIPO_NF
            sDoc = "NFFAT " & lNumeroNotaFiscalUlt

        Case TIPO_PARCELA
            sDoc = sSiglaDocumentoUlt & " " & lNumeroTituloUlt & "/" & iNumeroParcelaUlt

        Case TIPO_DEBITO
            sDoc = sSiglaDocumentoUlt & " " & lNumeroTituloUlt

        Case TIPO_TITULO_RECEBER
            sDoc = sSiglaDocumentoUlt & " " & lNumeroTituloUlt

        Case TIPO_COMISSAO_LOJA
            sDoc = "Cupom Fiscal"
            
        Case Else
            sDoc = ""

    End Select

    Comissao_Obter_Doc = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_Doc:

    Comissao_Obter_Doc = Err
    
    Select Case Err
    
        Case 59100
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153057)

    End Select
        
    Exit Function
            
End Function

Function Comissao_Obter_DocOrd(iTipoTitulo As Integer, lNumIntDoc As Long, sDocOrd As String, dtDataGeracao As Date) As Long
'obtem texto alfanumerico p/ordenar comissoes no relatorio
'os xxxxx e yy abaixo tem que ter 6 e 2 de tamanho completados com zeros à esquerda
'p/nffat: "0xxxxxx" (xxxxxx numero da nf)
'para titulo: "1xxxxxx" (xxxxxx numero do titulo)
'para parc: "2xxxxxxyy" (xxxxxx numero do titulo e yy o # da parcela)
'p/debitorec: "3xxxxxx"

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_DocOrd

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then

        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then Error 59101
    
    End If

    Select Case iTipoTitulo

        Case TIPO_NF
            sDocOrd = Format(lNumeroNotaFiscalUlt, "000000") & ",0"

        Case TIPO_TITULO_RECEBER
            sDocOrd = Format(lNumeroTituloUlt, "000000") & ",1"

        Case TIPO_PARCELA
            sDocOrd = Format(lNumeroTituloUlt, "000000") & ",2" & Format(iNumeroParcelaUlt, "00")

        Case TIPO_DEBITO
            sDocOrd = Format(lNumeroTituloUlt, "000000") & ",3"

        Case Else
            sDocOrd = ""

    End Select

    Comissao_Obter_DocOrd = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_DocOrd:

    Comissao_Obter_DocOrd = Err
    
    Select Case Err
    
        Case 59101
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153058)

    End Select
        
    Exit Function

End Function

Function Comissao_Obter_DataDoc(iTipoTitulo As Integer, lNumIntDoc As Long, dtDataDoc As Date, dtDataGeracao As Date) As Long
'obtem data associada ao docto gerador da comissao
'p/nffat: data da nf
'para titulo: data do titulo
'para parc: data do titulo
'p/debitorec: data do debito

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_DataDoc

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then

        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then Error 59102
        
    End If

    dtDataDoc = dtDataUlt

    Comissao_Obter_DataDoc = SUCESSO
    
    Exit Function
        
Erro_Comissao_Obter_DataDoc:

    Comissao_Obter_DataDoc = Err
    
    Select Case Err
    
        Case 59102
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153059)

    End Select
        
    Exit Function

End Function

Function ComissaoObterDados(iTipoTitulo As Integer, lNumIntDoc As Long, dtDataGeracao As Date) As Long
'Obtem os dados para a comissao

Dim lErro As Long
Dim lNumeroNotaFiscal As Long
Dim lNumeroTitulo As Long
Dim iNumeroParcela As Integer
Dim sSiglaDocumento As String
Dim dtData As Date, lNumIntTit As Long
Dim sNomeRedCli As String
Dim lNumeroPV As Long
Dim sNumeroPVCli As String
Dim sRazaoSocial As String 'Inserido por Wagner

On Error GoTo Erro_ComissaoObterDados

    Select Case iTipoTitulo
    
        Case TIPO_NF
        
            'Le os Dados para Nota Fiscal
            lErro = CF("ComissaoObterDados_NF", lNumIntDoc, lNumeroNotaFiscal, dtData, sNomeRedCli, lNumeroPV, lComandoNF, sNumeroPVCli)
            If lErro <> SUCESSO Then gError 59103
            
        Case TIPO_PARCELA
       
            'Le os Dados para a Parcela
            lErro = CF("ComissaoObterDados_ParcelaRec", lNumIntDoc, lNumeroTitulo, iNumeroParcela, sSiglaDocumento, dtData, sNomeRedCli, lNumIntTit, dtDataGeracao, lComandoParc)
            If lErro <> SUCESSO Then gError 59104
                            
            lErro = CF("ComissaoObterDados_NumPV_NFTitRec", lNumIntTit, lNumeroPV, lComandoNumPVTit, sNumeroPVCli)
            If lErro <> SUCESSO Then gError 59126
                            
        Case TIPO_DEBITO
                
            'Le os dados para o Debito
            lErro = CF("ComissaoObterDados_DebitosRecCli", lNumIntDoc, lNumeroTitulo, sSiglaDocumento, dtData, sNomeRedCli, lComandoDeb)
            If lErro <> SUCESSO Then gError 59105
        
        Case TIPO_TITULO_RECEBER
            
            'Le os Dados para o Titulo
            lErro = CF("ComissaoObterDados_TituloRec", lNumIntDoc, lNumeroTitulo, sSiglaDocumento, dtData, sNomeRedCli, lComandoTit)
            If lErro <> SUCESSO Then gError 59106
            
            lErro = CF("ComissaoObterDados_NumPV_NFTitRec", lNumIntDoc, lNumeroPV, lComandoNumPVTit, sNumeroPVCli)
            If lErro <> SUCESSO Then gError 59127
            
    End Select
        
    'Guarda os Dados nasa Variaveis Globais
    lNumeroNotaFiscalUlt = lNumeroNotaFiscal
    lNumeroTituloUlt = lNumeroTitulo
    iNumeroParcelaUlt = iNumeroParcela
    sSiglaDocumentoUlt = sSiglaDocumento
    dtDataUlt = dtData
    sNomeRedCliUlt = sNomeRedCli
    lNumeroPVUlt = lNumeroPV
    iTipoTituloUlt = iTipoTitulo
    lNumIntDocUlt = lNumIntDoc
    sNumeroPVCliUlt = sNumeroPVCli
        
    '####################################
    'Inserido por Wagner
    lErro = CF("ComissaoObterDados_RazaoSocial", sNomeRedCli, sRazaoSocial, lComandoCli)
    If lErro <> SUCESSO Then gError 132385
    
    sRazaoSocialUlt = sRazaoSocial
    '####################################
        
    ComissaoObterDados = SUCESSO
    
    Exit Function
    
Erro_ComissaoObterDados:

    ComissaoObterDados = gErr
    
    Select Case gErr
    
        Case 59103 To 59106, 59126, 59127, 132385
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153060)

    End Select
        
    Exit Function

End Function

Private Sub Class_Terminate()

    If lComandoNF <> 0 Then Call Comando_Fechar(lComandoNF)
    If lComandoParc <> 0 Then Call Comando_Fechar(lComandoParc)
    If lComandoDeb <> 0 Then Call Comando_Fechar(lComandoDeb)
    If lComandoTit <> 0 Then Call Comando_Fechar(lComandoTit)
    If lComandoNumPVTit <> 0 Then Call Comando_Fechar(lComandoNumPVTit)
    If lComandoCli <> 0 Then Call Comando_Fechar(lComandoCli) 'Inserido por Wagner

End Sub

Function Comissao_Obter_NumPVCli(iTipoTitulo As Integer, lNumIntDoc As Long, sNumPVCli As String, dtDataGeracao As Date) As Long
'obtem o numero do pedido se houver

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_NumPVCli

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then gError 133094
   
   End If

    sNumPVCli = sNumeroPVCliUlt

    Comissao_Obter_NumPVCli = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_NumPVCli:
    
    Comissao_Obter_NumPVCli = gErr
    
    Select Case gErr
    
        Case 133094
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153061)

    End Select
        
    Exit Function
    
End Function

Function Comissao_Obter_NumPV(iTipoTitulo As Integer, lNumIntDoc As Long, lNumPV As Long, dtDataGeracao As Date) As Long
'obtem o numero do pedido se houver

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_NumPV

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then gError 133470
   
   End If

    lNumPV = lNumeroPVUlt

    Comissao_Obter_NumPV = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_NumPV:
    
    Comissao_Obter_NumPV = gErr
    
    Select Case gErr
    
        Case 133470
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153062)

    End Select
        
    Exit Function
    
End Function

'#######################################################
'Inserido por Wagner
Function Comissao_Obter_RazaoSocialCli(iTipoTitulo As Integer, lNumIntDoc As Long, sRazaoSocial As String, dtDataGeracao As Date) As Long
'obtem a razão social do cliente

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_RazaoSocialCli

    If iTipoTitulo <> iTipoTituloUlt Or lNumIntDoc <> lNumIntDocUlt Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados(iTipoTitulo, lNumIntDoc, dtDataGeracao)
        If lErro <> SUCESSO Then gError 132383
      
   End If

    sRazaoSocial = sRazaoSocialUlt

    Comissao_Obter_RazaoSocialCli = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_RazaoSocialCli:
    
    Comissao_Obter_RazaoSocialCli = gErr
    
    Select Case gErr
    
        Case 132383
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
        
    Exit Function
    
End Function
'##################################################################

Function Comissao_Obter_TextoDataVenc(iTipoTitulo As Integer, lNumIntDoc As Long, sTextoDataVenc As String) As Long
'obtem a razão social do cliente

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_TextoDataVenc

    If iTipoTitulo <> iTipoTituloUltAux Or lNumIntDoc <> lNumIntDocUltAux Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados_Aux(iTipoTitulo, lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
      
   End If

    sTextoDataVenc = sTextoDataVencUltAux

    Comissao_Obter_TextoDataVenc = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_TextoDataVenc:
    
    Comissao_Obter_TextoDataVenc = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
        
    Exit Function
    
End Function

Function Comissao_Obter_TextoBaixa(iTipoTitulo As Integer, lNumIntDoc As Long, sTextoBaixaAux As String) As Long
'obtem texto associado à baixa do titulo ligado à comissao

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_TextoBaixa

    If iTipoTitulo <> iTipoTituloUltAux Or lNumIntDoc <> lNumIntDocUltAux Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados_Aux(iTipoTitulo, lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
      
   End If

    sTextoBaixaAux = sTextoBaixaUltAux

    Comissao_Obter_TextoBaixa = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_TextoBaixa:
    
    Comissao_Obter_TextoBaixa = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
        
    Exit Function
    
End Function

Function Comissao_Obter_EmBordero(iTipoTitulo As Integer, lNumIntDoc As Long, iEmBordero As Integer) As Long
'obtem a razão social do cliente

Dim lErro As Long

On Error GoTo Erro_Comissao_Obter_EmBordero

    If iTipoTitulo <> iTipoTituloUltAux Or lNumIntDoc <> lNumIntDocUltAux Then
    
        'Obtem Dados da Comissao
        lErro = ComissaoObterDados_Aux(iTipoTitulo, lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
      
   End If

    iEmBordero = iEmBorderoUltAux

    Comissao_Obter_EmBordero = SUCESSO
    
    Exit Function
    
Erro_Comissao_Obter_EmBordero:
    
    Comissao_Obter_EmBordero = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
        
    Exit Function
    
End Function

Function ComissaoObterDados_Aux(iTipoTitulo As Integer, lNumIntDoc As Long) As Long
'Obtem os dados para a comissao

Dim lErro As Long
Dim lNumNumIntNF As Long
Dim lNumIntTitulo As Long
Dim objNF As ClassNFiscal
Dim objTitRec As ClassTituloReceber
Dim objParcRec As ClassParcelaReceber
Dim colParcRec As ColParcelaReceber
Dim sTextoDataVencAux As String
Dim iEmBorderoAux As Integer
Dim sTextoBaixaAux As String, sTextoBaixaAux2 As String
Dim lComando As Long

On Error GoTo Erro_ComissaoObterDados_Aux

    sTextoDataVencAux = ""
    iEmBorderoAux = 0
    sTextoBaixaAux = ""

    'Abre o Comando
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 95380
    
    Select Case iTipoTitulo
    
        Case TIPO_NF, TIPO_TITULO_RECEBER
        
            If iTipoTitulo = TIPO_NF Then
            
                Set objNF = New ClassNFiscal
                
                objNF.lNumIntDoc = lNumIntDoc
            
                lErro = CF("NFiscal_Le", objNF)
                If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
            
                If objNF.iClasseDocCPR = CPR_TITULO_RECEBER Then
                    lNumIntTitulo = objNF.lNumIntDocCPR
                End If
            Else
                lNumIntTitulo = lNumIntDoc
            End If
            
            If lNumIntTitulo <> 0 Then
            
                Set objTitRec = New ClassTituloReceber
                
                objTitRec.lNumIntDoc = lNumIntTitulo
            
                'Lê o Número do Título
                lErro = CF("TituloReceber_Le", objTitRec)
                If lErro <> SUCESSO And lErro <> 26061 Then gError ERRO_SEM_MENSAGEM
                
                'Se não encontrou ---> Procura em Baixados
                If lErro <> SUCESSO Then
                    
                    lErro = CF("TituloReceberBaixado_Le", objTitRec)
                    If lErro <> SUCESSO And lErro <> 56570 Then gError ERRO_SEM_MENSAGEM
                
                End If
                
                Set colParcRec = New ColParcelaReceber
                
                lErro = CF("ParcelasReceber_Le_Todas", objTitRec, colParcRec)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objParcRec In colParcRec
                    If Len(Trim(sTextoDataVencAux)) > 0 Then sTextoDataVencAux = sTextoDataVencAux & "-"
                    sTextoDataVencAux = sTextoDataVencAux & Format(objParcRec.dtDataVencimento, "dd/mm")
                    
                    If objParcRec.iCobrador <> COBRADOR_PROPRIA_EMPRESA Then iEmBorderoAux = MARCADO
                    
                Next
        
                If objTitRec.dSaldo < DELTA_VALORMONETARIO Then
                
                    sTextoBaixaAux2 = String(STRING_CONTASCORRENTESINTERNAS_NOMEREDUZIDO, 0)
                    lErro = Comando_Executar(lComando, "SELECT C.NomeReduzido FROM ContasCorrentesInternas C, BaixasParcRec BP, BaixasRec B, MovimentosContaCorrente M, ParcelasRecTodas P WHERE C.Codigo = M.CodConta AND M.NumMovto = B.NumMovCta AND BP.Status <> 5 AND BP.NumIntBaixa = B.NumIntBaixa AND BP.NumIntParcela = P.NumIntDoc AND P.NumIntTitulo = ? ORDER BY C.NomeReduzido", _
                        sTextoBaixaAux2, objTitRec.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 201541

                    lErro = Comando_BuscarProximo(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201542
                    
                    If lErro = AD_SQL_SUCESSO Then
                    
                        sTextoBaixaAux = sTextoBaixaAux2
                        
                    Else
                    
                        sTextoBaixaAux = "BAIXADA"
                    
                    End If
                    
                End If
            
            End If
        
        Case TIPO_PARCELA
        
            Set objParcRec = New ClassParcelaReceber
            
            objParcRec.lNumIntDoc = lNumIntDoc

            lErro = CF("ParcelaReceber_Le", objParcRec)
            If lErro <> SUCESSO And lErro <> 19147 Then gError ERRO_SEM_MENSAGEM
             
            'Se não encontrou ---> Procura em Baixadas
            If lErro = 19147 Then
                
                'Procura em Baixadas
                lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
                If lErro <> SUCESSO And lErro <> 58559 Then gError ERRO_SEM_MENSAGEM
                
            End If
                    
            sTextoDataVencAux = Format(objParcRec.dtDataVencimento, "dd/mm")
            If objParcRec.iCobrador <> 0 Then iEmBorderoAux = MARCADO
                                
    End Select
    
    sTextoDataVencUltAux = sTextoDataVencAux
    iEmBorderoUltAux = iEmBorderoAux
    sTextoBaixaUltAux = sTextoBaixaAux
    
    iTipoTituloUltAux = iTipoTitulo
    lNumIntDocUltAux = lNumIntDoc
        
    Call Comando_Fechar(lComando)
    
    ComissaoObterDados_Aux = SUCESSO
    
    Exit Function
    
Erro_ComissaoObterDados_Aux:

    ComissaoObterDados_Aux = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153060)

    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

