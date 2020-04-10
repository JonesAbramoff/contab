VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelNFSe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lComandoDisc As Long
Private lComandoMunic As Long
Private lComandoServ As Long

Private lNumIntDocAnt As Long
Private lTamanho As Long
Private sDiscriminacao As String
Private asDiscriminacao(1 To 100) As String

Private lCodMunicAnt As Long
Private sNomeMunicipio As String

Private lCodMunicServAnt As Long
Private sCodServAnt As String
Private sDescricaoServ As String

Private Sub Class_Initialize()

On Error GoTo Erro_Class_Initialize
        
    lComandoDisc = Comando_AbrirExt(GL_lConexaoRel)
    If lComandoDisc = 0 Then gError 209764
    
    lComandoMunic = Comando_AbrirExt(GL_lConexaoRel)
    If lComandoMunic = 0 Then gError 209765
     
    lComandoServ = Comando_AbrirExt(GL_lConexaoRel)
    If lComandoServ = 0 Then gError 209775
     
    Exit Sub
    
Erro_Class_Initialize:
     
    Select Case gErr
          
        Case 209764, 209765, 209775
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209766)
     
    End Select
     
    Call Comando_Fechar(lComandoDisc)
    Call Comando_Fechar(lComandoMunic)
    
    Exit Sub
    
End Sub

Private Sub Class_Terminate()
    If lComandoDisc <> 0 Then Call Comando_Fechar(lComandoDisc)
    If lComandoMunic <> 0 Then Call Comando_Fechar(lComandoMunic)
    If lComandoServ <> 0 Then Call Comando_Fechar(lComandoServ)
End Sub

Public Function Obter_Discriminacao(ByVal lNumIntDoc As Long, ByVal iLinha As Integer, ByVal lTam As Long, sTexto As String) As Long

Dim lErro As Long, iIndice As Integer
Dim asDisc(1 To 8) As String
Dim sAux As String, sLinha As String
Dim iPosNewLine As Integer, iPosAnt As Integer, iPosEspaco As Integer

On Error GoTo Erro_Obter_Discriminacao

    If lNumIntDoc <> lNumIntDocAnt Then
    
        For iIndice = 1 To UBound(asDisc)
            asDisc(iIndice) = String(250, 0)
        Next
        
        'ler os itens da nf
        lErro = Comando_Executar(lComandoDisc, "SELECT SUBSTRING(ServicoDiscriminacao,1,250), SUBSTRING(ServicoDiscriminacao,251,250), SUBSTRING(ServicoDiscriminacao,501,250), SUBSTRING(ServicoDiscriminacao,751,250), SUBSTRING(ServicoDiscriminacao,1001,250), SUBSTRING(ServicoDiscriminacao,1251,250), SUBSTRING(ServicoDiscriminacao,1501,250), SUBSTRING(ServicoDiscriminacao,1751,250) FROM RPSWEBProt WHERE NumIntDoc = ?", asDisc(1), asDisc(2), asDisc(3), asDisc(4), asDisc(5), asDisc(6), asDisc(7), asDisc(8), lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 209767
        
        lErro = Comando_BuscarProximo(lComandoDisc)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209768
        
        lNumIntDocAnt = lNumIntDoc
        
        sDiscriminacao = ""
        For iIndice = 1 To UBound(asDisc)
            sDiscriminacao = sDiscriminacao & asDisc(iIndice)
        Next
        'sDiscriminacao = Replace(sDiscriminacao, vbCrLf, "|")
        'sDiscriminacao = Replace(sDiscriminacao, Chr(13), "|")
        sDiscriminacao = Replace(sDiscriminacao, Chr(10), "|")

        For iIndice = 1 To UBound(asDiscriminacao)
            asDiscriminacao(iIndice) = ""
        Next
        
        lTamanho = lTam
        iIndice = 0
        sAux = sDiscriminacao
        Do While Len(Trim(sAux)) > 0
        
            iIndice = iIndice + 1
            iPosNewLine = InStr(1, sAux, "|")
            
            'Se o tamanho máximo da linha é menor que a próxima quebra de linha
            'Tem que retornar até o último espaço antes do tamanho máximo
            If lTamanho < iPosNewLine Or (iPosNewLine = 0 And lTamanho < Len(sAux)) Then
                
                'Corta a mensagem até o máximo permitido
                sLinha = left(sAux, lTamanho)
                
                'Busca o último espaço antes do corte
                iPosAnt = 0
                iPosEspaco = InStr(1, sLinha, " ")
                Do While iPosEspaco <> 0
                    iPosAnt = iPosEspaco
                    iPosEspaco = InStr(iPosAnt + 1, sLinha, " ")
                Loop
                
                'Se achou algum espaço Corta a mensagem até ele
                If iPosAnt <> 0 Then
                    sLinha = left(sLinha, iPosAnt - 1)
                End If
                
            ElseIf iPosNewLine = 0 Then
            'Se a mensagem inteira e menor que o tamanho máximo
                sLinha = sAux
            Else
            'Se não retorna até a quebra de linha
                sLinha = left(sAux, iPosNewLine - 1)
            End If
            
            asDiscriminacao(iIndice) = sLinha
            sAux = Mid(sAux, Len(sLinha) + 1)
            iPosNewLine = InStr(1, sAux, "|")
            iPosEspaco = InStr(1, sAux, " ")
            If iPosNewLine = 1 Or iPosEspaco = 1 Then sAux = Mid(sAux, 2)
            
        Loop
        
    End If

    If lTam <> lTamanho Then gError 209769 'Não pode variar o tamanho, pois senão teria que forçar a chamada ser na ordem exata de exibição
    'para o correto cálculo do texto a ser exibido
    
    sTexto = asDiscriminacao(iLinha)
    
    Obter_Discriminacao = SUCESSO
     
    Exit Function
    
Erro_Obter_Discriminacao:

    Obter_Discriminacao = gErr
     
    Select Case gErr

        Case 209767, 209768
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBPROT", gErr)
            
        Case 209769
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_OBTER_DISC_TAM_DIF", gErr, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209770)
     
    End Select
    
    Exit Function
    
End Function

Public Function Obter_NomeMunic(ByVal lCodMunic As Long, sTexto As String) As Long

Dim lErro As Long
Dim sNomeMunic As String

On Error GoTo Erro_Obter_NomeMunic

    If lCodMunic <> lCodMunicAnt Then
    
        sNomeMunic = String(STRING_MAXIMO, 0)
        
        'ler os itens da nf
        lErro = Comando_Executar(lComandoMunic, "SELECT Munic FROM IBGEMunicipios WHERE CodIBGE = ?", sNomeMunic, CStr(lCodMunic))
        If lErro <> AD_SQL_SUCESSO Then gError 209771
        
        lErro = Comando_BuscarProximo(lComandoMunic)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209772
        
        If lErro <> AD_SQL_SUCESSO Then gError 209773
        
        lCodMunicAnt = lCodMunic
        sNomeMunicipio = UCase(sNomeMunic)
        
    End If

    sTexto = sNomeMunicipio
    
    Obter_NomeMunic = SUCESSO
     
    Exit Function
    
Erro_Obter_NomeMunic:

    Obter_NomeMunic = gErr
     
    Select Case gErr

        Case 209771, 209772
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr)
            
        Case 209773
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_INEXISTENTE", gErr, CStr(lCodMunic))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209774)
     
    End Select
    
    Exit Function
    
End Function

Public Function Obter_DescServ(ByVal lCodMunic As Long, ByVal sCodServ As String, sTexto As String) As Long

Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_Obter_DescServ

    If lCodMunic <> lCodMunicServAnt Or sCodServAnt <> sCodServ Then
    
        sDescricao = String(STRING_MAXIMO, 0)
        
        'ler os itens da nf
        lErro = Comando_Executar(lComandoServ, "SELECT Descricao1 FROM CodServMun WHERE CodIBGE = ? AND CodServ = ?", sDescricao, lCodMunic, sCodServ)
        If lErro <> AD_SQL_SUCESSO Then gError 209776
        
        lErro = Comando_BuscarProximo(lComandoServ)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209777
        
        lCodMunicServAnt = lCodMunic
        sCodServAnt = sCodServ
        sDescricaoServ = sDescricao
        
    End If

    sTexto = sDescricaoServ
    
    Obter_DescServ = SUCESSO
     
    Exit Function
    
Erro_Obter_DescServ:

    Obter_DescServ = gErr
     
    Select Case gErr

        Case 209776, 209777
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RPSWEBPROT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209778)
     
    End Select
    
    Exit Function
    
End Function
