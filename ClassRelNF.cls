VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Criada para atender Notas Fiscais

Private lComando As Long

Private lUltNumIntNFMenorPag As Long
Private lMenorPaginaNF As Long

Private lUltNumIntNFPag As Long
Private lPaginaAcumulada As Long

Private lUltNumIntNFPagDanfe As Long
Private lPaginaAcumuladaDanfe As Long
Private lNumChamadasDanfe As Long

Private lUltNumIntNFPagInpal As Long
Private iPaginaAcumuladaInpal As Long

Private lUltNumIntNFPagGuedes As Long
Private iPaginaAcumuladaGuedes As Long

Private lUltNumIntNFPagHar As Long
Private iPaginaAcumuladaHar As Long

Private lUltNumIntNFPagCromaton As Long
Private iPaginaAcumuladaCromaton As Long

Private lUltNumIntNFPagPhar As Long
Private iPaginaAcumuladaPhar As Long

'para obtencao de debitos ou creditos de nfs
Private lComandoTribNF As Long
Private lComandoTribItemNF As Long
Private lUltNumIntNF As Long
Private dUltValorBase As Double
Private dUltValor As Double
Private dUltValorBaseSubst As Double
Private dUltValorSubst As Double
Private iUltimoNumItens As Integer

Private iAnoUltimo As Integer
Private iMesUltimo As Integer
Private dAliquotaICMSUltimo As Double

Private lNumIntNFCredSimples As Long
Private dUltNFCredSimples As Double
Private dUltNFBaseSimples As Double

Private sUltchNFe As String
Private colParcRec As New ColParcelaReceber
Private lComandoParcRec1 As Long
Private lComandoParcRec2 As Long

Public Function ObterSituacaoTrib(sSituacaoTrib As String, sProduto As String, iTipoTribICMS As Integer) As Long
'Para obter o Código da Situação Tributária

Dim lErro As Long

On Error GoTo Erro_ObterSituacaoTrib

    'Le a situação tributaria do Produto
    lErro = CF("Produto_Le_SituacaoTrib", sSituacaoTrib, sProduto, iTipoTribICMS, lComando)
    If lErro <> SUCESSO Then Error 58572
    
    ObterSituacaoTrib = SUCESSO
    
    Exit Function
    
Erro_ObterSituacaoTrib:

    ObterSituacaoTrib = Err
    
    Select Case Err

        Case 58572 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153283)

    End Select

    Exit Function

End Function


Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

    If lComandoTribNF <> 0 Then Call Comando_Fechar(lComandoTribNF)
    If lComandoTribItemNF <> 0 Then Call Comando_Fechar(lComandoTribItemNF)
    If lComandoParcRec1 <> 0 Then Call Comando_Fechar(lComandoParcRec1)
    If lComandoParcRec2 <> 0 Then Call Comando_Fechar(lComandoParcRec2)

End Sub

Public Function TemDebitoOuCredito(ByVal lNumIntNF As Long, iCredito As Integer) As Long
'retorna em iCredito: 0:nem credito nem debito, 1:tem credito ou debito

Dim lErro As Long

On Error GoTo Erro_TemDebitoOuCredito

    lErro = ObterInfoICMS(lNumIntNF)
    If lErro <> SUCESSO Then Error 59436
    
    If (dUltValor <> 0 Or dUltValorSubst <> 0) Then
        iCredito = 1
    Else
        iCredito = 0
    End If
    
    TemDebitoOuCredito = SUCESSO
     
    Exit Function
    
Erro_TemDebitoOuCredito:

    TemDebitoOuCredito = Err
     
    Select Case Err
          
        Case 59436
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153284)
     
    End Select
     
    Exit Function

End Function

Public Function ObterICMSBase(ByVal lNumIntNF As Long, dValorBase As Double) As Long
'retorna em dValorBase a base de calculo de icms que gerou debito ou credito

Dim lErro As Long

On Error GoTo Erro_ObterICMSBase

    lErro = ObterInfoICMS(lNumIntNF)
    If lErro <> SUCESSO Then Error 59437
    
    dValorBase = dUltValorBase
    
    ObterICMSBase = SUCESSO
     
    Exit Function
    
Erro_ObterICMSBase:

    ObterICMSBase = Err
     
    Select Case Err
          
        Case 59437
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153285)
     
    End Select
     
    Exit Function

End Function

Public Function ObterICMSValor(ByVal lNumIntNF As Long, dValor As Double) As Long
'retorna em dValor o valor de icms que gerou debito ou credito

Dim lErro As Long

On Error GoTo Erro_ObterICMSValor

    lErro = ObterInfoICMS(lNumIntNF)
    If lErro <> SUCESSO Then Error 59438
    
    dValor = dUltValor
    
    ObterICMSValor = SUCESSO
     
    Exit Function
    
Erro_ObterICMSValor:

    ObterICMSValor = Err
     
    Select Case Err
          
        Case 59438
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153286)
     
    End Select
     
    Exit Function

End Function

Public Function ObterICMSSubstBase(ByVal lNumIntNF As Long, dValorBaseSubst As Double) As Long
'retorna em dValorBaseSubst a base de calculo de icms de substituicao que gerou debito ou credito

Dim lErro As Long

On Error GoTo Erro_ObterICMSSubstBase

    lErro = ObterInfoICMS(lNumIntNF)
    If lErro <> SUCESSO Then Error 59439
    
    dValorBaseSubst = dUltValorBaseSubst
    
    ObterICMSSubstBase = SUCESSO
     
    Exit Function
    
Erro_ObterICMSSubstBase:

    ObterICMSSubstBase = Err
     
    Select Case Err
          
        Case 59439
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153287)
     
    End Select
     
    Exit Function

End Function

Public Function ObterICMSSubstValor(ByVal lNumIntNF As Long, dValorSubst As Double) As Long
'retorna em dValorSubst o valor de icms de substituicao que gerou debito ou credito

Dim lErro As Long

On Error GoTo Erro_ObterICMSSubstValor

    lErro = ObterInfoICMS(lNumIntNF)
    If lErro <> SUCESSO Then Error 59440
    
    dValorSubst = dUltValorSubst
    
    ObterICMSSubstValor = SUCESSO
     
    Exit Function
    
Erro_ObterICMSSubstValor:

    ObterICMSSubstValor = Err
     
    Select Case Err
          
        Case 59440
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153288)
     
    End Select
     
    Exit Function

End Function

Private Function ObterInfoICMS(ByVal lNumIntNF As Long) As Long
'obtem informacoes sobre valores de credito ou debito de icms de uma nota fiscal
'Os resultados sao armazenados nas variaveis globais à esta classe

Dim lErro As Long, dValorBase As Double, dValor As Double, dValorBaseSubst As Double, dValorSubst As Double

On Error GoTo Erro_ObterInfoICMS

    If lNumIntNF <> lUltNumIntNF Then
    
        lErro = CF("NFiscal_ObterDebCredICMS", lNumIntNF, dValorBase, dValor, dValorBaseSubst, dValorSubst, lComandoTribNF, lComandoTribItemNF)
        If lErro <> SUCESSO Then Error 59424
        
        'armazena valores para nao ter que repetir a pesquisa
        dUltValorBase = dValorBase
        dUltValor = dValor
        dUltValorBaseSubst = dValorBaseSubst
        dUltValorSubst = dValorSubst
        
        lUltNumIntNF = lNumIntNF
    
    End If
    
    ObterInfoICMS = SUCESSO
     
    Exit Function
    
Erro_ObterInfoICMS:

    ObterInfoICMS = Err
     
    Select Case Err
          
        Case 59424
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153289)
     
    End Select
     
    Exit Function

End Function

Public Function ObterPaginasNFGuedes(lNumIntDoc As Long, iMaxLinhasPagina As Integer, iNumPaginas As Integer, iMaxCaracLinha As Integer, iLinhasDesconto As Integer, iLinhasSuframa As Integer) As Long
'Para obter o Número de Páginas

Dim lErro As Long, sProduto As String, sDescricaoItem As String, iLinhasItem As Integer
Dim lComando As Long, iPaginas As Integer, iLinhasAcum As Integer
Dim sProdCurto As String

On Error GoTo Erro_ObterPaginasNFGuedes

    If lNumIntDoc <> lUltNumIntNFPagGuedes Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 81810
                
        iPaginas = 1
        iLinhasAcum = 0 + iLinhasDesconto + iLinhasSuframa
        
        sProduto = String(STRING_PRODUTO, 0)
        sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        
        'ler os itens da nf
        lErro = Comando_Executar(lComando, "SELECT Produto, DescricaoItem FROM ItensNFiscal WHERE NumIntNF = ?", sProduto, sDescricaoItem, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 81811
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81812
    
        Do While lErro = AD_SQL_SUCESSO
        
            'obter numero de linhas do item
            If Len(sDescricaoItem) = 0 Then
                iLinhasItem = 1
            Else
                iLinhasItem = SymUp(CDbl(Len(sDescricaoItem)) / iMaxCaracLinha)
            End If
            
            If iLinhasAcum + iLinhasItem > iMaxLinhasPagina Then
            
                iPaginas = iPaginas + 1
                iLinhasAcum = iLinhasAcum + iLinhasItem - iMaxLinhasPagina
            
            Else
            
                iLinhasAcum = iLinhasAcum + iLinhasItem
            
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81813
    
        Loop
    
        lUltNumIntNFPagGuedes = lNumIntDoc
        iPaginaAcumuladaGuedes = iPaginas
        
        Call Comando_Fechar(lComando)
    
    End If
        
    iNumPaginas = iPaginaAcumuladaGuedes
    
    ObterPaginasNFGuedes = SUCESSO
    
    Exit Function
    
Erro_ObterPaginasNFGuedes:

    ObterPaginasNFGuedes = gErr
    
    Select Case gErr

        Case 69087 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153290)

    End Select

    Exit Function

End Function

Public Function ObterPaginasNF(lNumIntDoc As Long, lNumeroMaximoPagina As Long, lNumPaginas As Long) As Long
'Para obter o Número de Páginas

Dim lErro As Long
Dim lPaginaTemp As Double
Dim iResto As Integer
Dim iNumItens As Integer

On Error GoTo Erro_ObterPaginasNF

    If lNumIntDoc <> lUltNumIntNFPag Then
    
        'Le a Quantidade de Itens da Nota
        lErro = CF("NF_Obter_NumeroItens", lNumIntDoc, iNumItens, lComando)
        If lErro <> SUCESSO Then gError 69087
        
        iResto = Resto(CDbl(iNumItens), CInt(lNumeroMaximoPagina))
        
        If iResto = 0 Then
            lPaginaTemp = iNumItens / lNumeroMaximoPagina
        Else
            lPaginaTemp = Fix(iNumItens / lNumeroMaximoPagina) + 1
        End If
        
        lUltNumIntNFPag = lNumIntDoc
        lPaginaAcumulada = lPaginaTemp
    
    End If
        
    lNumPaginas = lPaginaAcumulada
    
    ObterPaginasNF = SUCESSO
    
    Exit Function
    
Erro_ObterPaginasNF:

    ObterPaginasNF = gErr
    
    Select Case gErr

        Case 69087 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153291)

    End Select

    Exit Function

End Function

Function SymUp(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

Dim temp As Double
     
     temp = Fix(X * Factor)
     SymUp = (temp + IIf(X = temp, 0, Sgn(X))) / Factor
     
End Function

Public Function ObterPaginasNFInpal(lNumIntDoc As Long, iNumPaginas As Integer, iMaxLinhasPagina As Integer, iMaxCaracLinha As Integer) As Long
'Para obter o Número de Páginas para a impressao de uma nota fiscal identificada por lNumIntDoc.
'Leva em conta o maximo de linhas por pagina e o numero maximo de caracteres p/impressao de descricao de item da nf, p/linha

Dim lErro As Long, sProduto As String, sDescricaoItem As String, iLinhasItem As Integer
Dim lComando As Long, iPaginas As Integer, iLinhasAcum As Integer, sDescricaoAux As String
Dim X As New ClassRelAux, sProdCurto As String

On Error GoTo Erro_ObterPaginasNFInpal

    If lNumIntDoc <> lUltNumIntNFPagInpal Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 81810
                
        iPaginas = 1
        iLinhasAcum = 0
        
        sProduto = String(STRING_PRODUTO, 0)
        sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        
        'ler os itens da nf
        lErro = Comando_Executar(lComando, "SELECT Produto, DescricaoItem FROM ItensNFiscal WHERE NumIntNF = ?", sProduto, sDescricaoItem, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 81811
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81812
    
        Do While lErro = AD_SQL_SUCESSO
        
'            Call X.EsvaziaProduto_INPAL(sProdCurto, sProduto)
'            If Len(sProdCurto) > 4 Then
'
'                sDescricaoAux = sProdCurto & " - " & sDescricaoItem
'
'            Else
                
                sDescricaoAux = sDescricaoItem
            
'            End If
            
            'obter numero de linhas do item
            If Len(sDescricaoAux) = 0 Then
                iLinhasItem = 1
            Else
                iLinhasItem = SymUp(CDbl(Len(sDescricaoAux)) / iMaxCaracLinha)
            End If
            
            If iLinhasAcum + iLinhasItem > iMaxLinhasPagina Then
            
                iPaginas = iPaginas + 1
                iLinhasAcum = iLinhasAcum + iLinhasItem - iMaxLinhasPagina
            
            Else
            
                iLinhasAcum = iLinhasAcum + iLinhasItem
            
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81813
    
        Loop
    
        lUltNumIntNFPagInpal = lNumIntDoc
        iPaginaAcumuladaInpal = iPaginas
        
        Call Comando_Fechar(lComando)
    
    End If
        
    iNumPaginas = iPaginaAcumuladaInpal
    
    ObterPaginasNFInpal = SUCESSO
     
    Exit Function
    
Erro_ObterPaginasNFInpal:

    ObterPaginasNFInpal = gErr
     
    Select Case gErr
          
        Case 81810
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 81811, 81812, 81813
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENS_NF_REL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153292)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'??? ERRO_LEITURA_ITENS_NF_REL

Public Function ObterPagNF(lNumIntDoc As Long, iNumPag As Integer, lNumPagAtualEmissao As Long) As Long
'obtem o numero da pagina corrente da nota fiscal (1a pagina da nota, 2a,...)
'lNumPagAtualEmissao é o numero da pagina na emissao de notas sendo executada

Dim lErro As Long

On Error GoTo Erro_ObterPagNF

    If lNumIntDoc <> lUltNumIntNFMenorPag Then
    
        lMenorPaginaNF = lNumPagAtualEmissao
        lUltNumIntNFMenorPag = lNumIntDoc
    
    End If
    
    iNumPag = lNumPagAtualEmissao - lMenorPaginaNF + 1
    
    ObterPagNF = SUCESSO
     
    Exit Function
    
Erro_ObterPagNF:

    ObterPagNF = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153293)
     
    End Select
     
    Exit Function

End Function

Public Function ObterNumLoteItensNFiscal(sNumLoteRetorno As String, ByVal lNumIntItemNF As Long) As Long
'Função que vai ao BD e lê os Lotes para um determinado produto passados como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDocMovEst As Long
Dim iIndice As Integer
Dim sRetornoCodLote As String
Dim sSQL As String

On Error GoTo Erro_ObterNumLoteItensNFiscal
 
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111301
    
    'Inicializar a String
    sRetornoCodLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
    'Lê os Lotes Fazendo o Join das tabelas ItensNiscal , Movimentos Estoque , RastreamentoMovto , RastreamentoLote
    lErro = Comando_Executar(lComando, "SELECT Distinct RastreamentoLote.Lote FROM RastreamentoMovto, RastreamentoLote, MovimentoEstoque, ItensNFiscal " _
    & "WHERE ItensNFiscal.NumIntDoc = ? AND ItensNFiscal.NumIntDoc = MovimentoEstoque.NumIntDocOrigem AND MovimentoEstoque.TipoNumIntDocOrigem = 1 AND RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoMovto.TipoDocOrigem = 0 AND " _
    & "RastreamentoMovto.NumIntDocOrigem = MovimentoEstoque.NumIntDoc", sRetornoCodLote, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 111302
     
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111303
    
    
    Do While lErro = SUCESSO
    
        'Passa o Numero do Lote para a função de retorno
        sNumLoteRetorno = sNumLoteRetorno & sRetornoCodLote & "/"
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111305
    
    Loop

    If Len(sNumLoteRetorno) > 0 Then
        
        sNumLoteRetorno = left(sNumLoteRetorno, Len(sNumLoteRetorno) - 1)
    
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
       
    ObterNumLoteItensNFiscal = SUCESSO

    Exit Function

Erro_ObterNumLoteItensNFiscal:

    ObterNumLoteItensNFiscal = gErr

    Select Case gErr

        Case 111301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111302, 111303, 111305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE_VAZIO_NF", gErr, lNumIntItemNF)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153294)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ObterSituacaoTrib1(sSituacaoTrib As String, sProduto As String, iTipoTribICMS As Integer, iTipoTribIPI As Integer) As Long
'Para obter o Código da Situação Tributária

Dim lErro As Long

On Error GoTo Erro_ObterSituacaoTrib1

    'Le a situação tributaria do Produto
    lErro = CF("Define_SituacaoTributaria_Produto", sSituacaoTrib, sProduto, iTipoTribICMS, iTipoTribIPI)
    If lErro <> SUCESSO Then gError 108765
    
    ObterSituacaoTrib1 = SUCESSO
    
    Exit Function
    
Erro_ObterSituacaoTrib1:

    ObterSituacaoTrib1 = gErr
    
    Select Case gErr

        Case 108765 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153295)

    End Select

End Function

Public Function ObterPaginasNFCromaton(lNumIntDoc As Long, iMaxLinhasPagina As Integer, iNumPaginas As Integer, iMaxCaracLinha As Integer, iLinhasDesconto As Integer, iLinhasSuframa As Integer) As Long
'Para obter o Número de Páginas

Dim lErro As Long, sProduto As String, sDescricaoItem As String, iLinhasItem As Integer
Dim lComando As Long, iPaginas As Integer, iLinhasAcum As Integer
Dim sProdCurto As String, iQuantLotes As Integer, lNumIntItemNF As Long

On Error GoTo Erro_ObterPaginasNFCromaton

    If lNumIntDoc <> lUltNumIntNFPagCromaton Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 81810
                
        iPaginas = 1
        iLinhasAcum = 0 + iLinhasDesconto + iLinhasSuframa
        
        sProduto = String(STRING_PRODUTO, 0)
        sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        
        'ler os itens da nf
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, DescricaoItem FROM ItensNFiscal WHERE NumIntNF = ?", lNumIntItemNF, sProduto, sDescricaoItem, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 81811
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81812
    
        Do While lErro = AD_SQL_SUCESSO
        
            'obter numero de linhas do item
            If Len(sDescricaoItem) = 0 Then
                iLinhasItem = 1
            Else
                iLinhasItem = SymUp(CDbl(Len(sDescricaoItem)) / iMaxCaracLinha)
            End If
            
            lErro = ObterQuantLoteItensNFiscal(lNumIntItemNF, iQuantLotes)
            If lErro <> SUCESSO Then gError 130337
            
            If iQuantLotes > 1 Then
                If iLinhasItem < (iQuantLotes + 1) Then iLinhasItem = iQuantLotes + 1
            End If
            
            If iLinhasAcum + iLinhasItem > iMaxLinhasPagina Then
            
                iPaginas = iPaginas + 1
                iLinhasAcum = iLinhasAcum + iLinhasItem - iMaxLinhasPagina
            
            Else
            
                iLinhasAcum = iLinhasAcum + iLinhasItem
            
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81813
    
        Loop
    
        lUltNumIntNFPagCromaton = lNumIntDoc
        iPaginaAcumuladaCromaton = iPaginas
        
        Call Comando_Fechar(lComando)
    
    End If
        
    iNumPaginas = iPaginaAcumuladaCromaton
    
    ObterPaginasNFCromaton = SUCESSO
    
    Exit Function
    
Erro_ObterPaginasNFCromaton:

    ObterPaginasNFCromaton = gErr
    
    Select Case gErr

        Case 130337, 69087 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153296)

    End Select

    Exit Function

End Function

Private Function ObterQuantLoteItensNFiscal(ByVal lNumIntItemNF As Long, iQuantLotes As Integer) As Long
'Função que vai ao BD e lê a qtde de Lotes para um determinado item de nf passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDocMovEst As Long
Dim iIndice As Integer
Dim sRetornoCodLote As String
Dim sSQL As String

On Error GoTo Erro_ObterQuantLoteItensNFiscal
 
    iQuantLotes = 0
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111301
    
    'Inicializar a String
    sRetornoCodLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
    'Lê os Lotes Fazendo o Join das tabelas ItensNiscal , Movimentos Estoque , RastreamentoMovto , RastreamentoLote
    lErro = Comando_Executar(lComando, "SELECT Distinct RastreamentoLote.Lote FROM RastreamentoMovto, RastreamentoLote, MovimentoEstoque, ItensNFiscal " _
    & "WHERE ItensNFiscal.NumIntDoc = ? AND ItensNFiscal.NumIntDoc = MovimentoEstoque.NumIntDocOrigem AND MovimentoEstoque.TipoNumIntDocOrigem = 1 AND RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoMovto.TipoDocOrigem = 0 AND " _
    & "RastreamentoMovto.NumIntDocOrigem = MovimentoEstoque.NumIntDoc", sRetornoCodLote, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 111302
     
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111303
    
    Do While lErro = SUCESSO
    
        iQuantLotes = iQuantLotes + 1
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111305
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)
       
    ObterQuantLoteItensNFiscal = SUCESSO

    Exit Function

Erro_ObterQuantLoteItensNFiscal:

    ObterQuantLoteItensNFiscal = gErr

    Select Case gErr

        Case 111301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111302, 111303, 111305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE_VAZIO_NF", gErr, lNumIntItemNF)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153297)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ObterPaginasNFHar(lNumIntDoc As Long, iMaxLinhasPagina As Integer, iNumPaginas As Integer, iMaxCaracLinha As Integer, iLinhasDesconto As Integer, iLinhasSuframa As Integer) As Long
'Para obter o Número de Páginas

Dim lErro As Long, sDescricaoItem As String, iLinhasItem As Integer
Dim lComando As Long, iPaginas As Integer, iLinhasAcum As Integer
Dim sProdCurto As String, sSegmento1 As String, sSegmento1Ant As String

On Error GoTo Erro_ObterPaginasNFHar

    If lNumIntDoc <> lUltNumIntNFPagHar Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 181815
                
        sSegmento1Ant = ""
        iPaginas = 1
        iLinhasAcum = 0 + iLinhasDesconto + iLinhasSuframa
        
        sSegmento1 = String(255, 0)
        sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        
        'ler os itens da nf
        lErro = Comando_Executar(lComando, "SELECT LEFT(Produtos.NomeReduzido,7) + SUBSTRING(Produtos.NomeReduzido,16,2), DescricaoItem FROM NFiscais_FornCli, Produtos WHERE NFiscais_FornCli.Produto = Produtos.Codigo AND NumIntDoc = ? ORDER BY LEFT(Produtos.NomeReduzido,7) + SUBSTRING(Produtos.NomeReduzido,16,2), Item", _
            sSegmento1, sDescricaoItem, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 181816
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181817
    
        Do While lErro = AD_SQL_SUCESSO
        
            If sSegmento1Ant <> sSegmento1 Then
            
                'obter numero de linhas do item
                If Len(sDescricaoItem) = 0 Then
                    iLinhasItem = 1
                Else
                    iLinhasItem = SymUp(CDbl(Len(sDescricaoItem)) / iMaxCaracLinha)
                End If
                
                If iLinhasAcum + iLinhasItem > iMaxLinhasPagina Then
                
                    iPaginas = iPaginas + 1
                    iLinhasAcum = iLinhasAcum + iLinhasItem - iMaxLinhasPagina
                
                Else
                
                    iLinhasAcum = iLinhasAcum + iLinhasItem
                
                End If
            
                sSegmento1Ant = sSegmento1
            
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181818
    
        Loop
    
        lUltNumIntNFPagHar = lNumIntDoc
        iPaginaAcumuladaHar = iPaginas
        
        Call Comando_Fechar(lComando)
    
    End If
        
    iNumPaginas = iPaginaAcumuladaHar
    
    ObterPaginasNFHar = SUCESSO
    
    Exit Function
    
Erro_ObterPaginasNFHar:

    ObterPaginasNFHar = gErr
    
    Select Case gErr

        Case 181815
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 181816, 181817, 181818
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181819)

    End Select

    Exit Function

End Function

Public Function ObterListaCFOP(lNumIntDoc As Long, sListaCFOP As String) As Long
'Para obter o Número de Páginas

Dim lErro As Long
Dim lComando As Long
Dim sNatOP As String
Dim sListaCFOPAux As String

On Error GoTo Erro_ObterListaCFOP
    
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 194852
    
    sNatOP = String(STRING_NATUREZAOP_CODIGO, 0)
    
    'ler as naturezas de operação ligadas as máquinas
    lErro = Comando_Executar(lComando, "SELECT DISTINCT NaturezaOP FROM TributacaoItemNF WHERE NumIntNF = ? ORDER BY NaturezaOP", sNatOP, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194853
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194854

    Do While lErro = AD_SQL_SUCESSO
    
        If sListaCFOPAux <> "" Then
            sListaCFOPAux = sListaCFOPAux & "\" & sNatOP
        Else
            sListaCFOPAux = sNatOP
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194855

    Loop
    
    sListaCFOP = sListaCFOPAux
        
    Call Comando_Fechar(lComando)
    
    ObterListaCFOP = SUCESSO
    
    Exit Function
    
Erro_ObterListaCFOP:

    ObterListaCFOP = gErr
    
    Select Case gErr

        Case 194852
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 194853, 194854, 194855
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194856)

    End Select

    Exit Function

End Function

Public Function ObterPaginasNFPhar(lNumIntDoc As Long, lNumeroMaximoPagina As Long, lNumPaginas As Long) As Long
'Para obter o Número de Páginas

Dim lErro As Long
Dim lPaginaTemp As Double
Dim iResto As Integer
Dim iNumItens As Integer

On Error GoTo Erro_ObterPaginasNFPhar

    If lNumIntDoc <> lUltNumIntNFPagPhar Then
    
        'Le a Quantidade de Itens da Nota
        lErro = NF_Obter_NumeroItens_Phar(lNumIntDoc, iNumItens, lComando)
        If lErro <> SUCESSO Then gError 69087
        
        iResto = Resto(CDbl(iNumItens), CInt(lNumeroMaximoPagina))
        
        If iResto = 0 Then
            lPaginaTemp = iNumItens / lNumeroMaximoPagina
        Else
            lPaginaTemp = Fix(iNumItens / lNumeroMaximoPagina) + 1
        End If
        
        lUltNumIntNFPagPhar = lNumIntDoc
        iPaginaAcumuladaPhar = lPaginaTemp
    
    End If
        
    lNumPaginas = iPaginaAcumuladaPhar
    
    ObterPaginasNFPhar = SUCESSO
    
    Exit Function
    
Erro_ObterPaginasNFPhar:

    ObterPaginasNFPhar = gErr
    
    Select Case gErr

        Case 69087 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153291)

    End Select

    Exit Function

End Function

Private Function NF_Obter_NumeroItens_Phar(ByVal lNumIntDoc As Long, iNumItens As Integer, lComando As Long) As Long
'OBS: Função para o cálculo de Número de Itens de Nota Fiscal para Relatorio
'Se o Comando vier zerado vai ser aberto e não será fechado

Dim lErro As Long

On Error GoTo Erro_NF_Obter_NumeroItens_Phar

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 69088
    End If

    lErro = Comando_Executar(lComando, "SELECT COUNT(*)+(SELECT COUNT(*) FROM ItensNFiscal ITN WHERE ITN.ProdutoKit <> '' AND ITN.SeqNoPack=1 AND ITN.NumIntNF = ?) FROM ItensNFiscal INF WHERE NumIntNF = ?", iNumItens, lNumIntDoc, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69089

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69090

    NF_Obter_NumeroItens_Phar = SUCESSO
    
    Exit Function
    
Erro_NF_Obter_NumeroItens_Phar:

    NF_Obter_NumeroItens_Phar = gErr
    
    Select Case gErr
    
        Case 69088
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 69089, 69090
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150187)
     
    End Select
     
    Exit Function

End Function

Public Function Cliente_Possui_CatItem(ByVal lCliente As Long, ByVal iFilial As Integer, ByVal sCategoria As String, ByVal sItem As String, iPossuiCatItem As Double) As Long
'retorna em iPossuiCatItem: 0:A filial não possui a categoria ou item informado, 1:A filial  possui o item de categoria informado

Dim lErro As Long
Dim objFilialCliCategoria As New ClassFilialCliCategoria

On Error GoTo Erro_Cliente_Possui_CatItem

    objFilialCliCategoria.lCliente = lCliente
    objFilialCliCategoria.iFilial = iFilial
    objFilialCliCategoria.sCategoria = sCategoria
    objFilialCliCategoria.sItem = sItem

    lErro = CF("FilialClienteCategorias_Le", objFilialCliCategoria)
    If lErro <> SUCESSO And lErro <> 199401 Then gError 200964
    If lErro <> SUCESSO Then
        iPossuiCatItem = DESMARCADO
    Else
        iPossuiCatItem = MARCADO
    End If
        
    Cliente_Possui_CatItem = SUCESSO
     
    Exit Function
    
Erro_Cliente_Possui_CatItem:

    Cliente_Possui_CatItem = gErr
     
    Select Case gErr
          
        Case 200964
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200965)
     
    End Select
     
    Exit Function

End Function

Public Function Obtem_AliquotaICMSDAS(ByVal dtDataEmissaoNF As Date, dAliquota As Double) As Long
'retorna em dAliquota: a alíquota de ICMS da DAS que pode ser creditada na Nf

Dim lErro As Long
Dim objDASAliquotas As New ClassDASAliquotas

On Error GoTo Erro_Obtem_AliquotaICMSDAS

    objDASAliquotas.iAno = Year(DateAdd("m", -1, dtDataEmissaoNF))
    objDASAliquotas.iMes = Month(DateAdd("m", -1, dtDataEmissaoNF))

    If objDASAliquotas.iAno <> iAnoUltimo Or objDASAliquotas.iMes <> iMesUltimo Then
    
        lErro = CF("DASAliquotas_Le", objDASAliquotas)
        If lErro <> SUCESSO And lErro <> 200911 Then gError 200966
        
        If lErro = 200911 Then gError 200967
        
        iAnoUltimo = objDASAliquotas.iAno
        iMesUltimo = objDASAliquotas.iMes
        dAliquotaICMSUltimo = objDASAliquotas.dAliquotaICMS
    
    End If
        
    dAliquota = dAliquotaICMSUltimo
        
    Obtem_AliquotaICMSDAS = SUCESSO
     
    Exit Function
    
Erro_Obtem_AliquotaICMSDAS:

    Obtem_AliquotaICMSDAS = gErr
     
    Select Case gErr
          
        Case 200966
        
        Case 200967
            Call Rotina_Erro(vbOKOnly, "ERRO_DASALIQUOTAS_NAO_CADASTRADO", gErr, objDASAliquotas.iAno, objDASAliquotas.iMes)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200968)
     
    End Select
     
    Exit Function

End Function

Public Function ObterBaseICMSSaidaSimples(dValor As Double, ByVal lNumIntNF As Long) As Long
Dim dVal As Double, lErro As Long
    
    lErro = ObterCredICMSSaidaSimples(dVal, lNumIntNF, DATA_NULA)
    If lErro = SUCESSO Then dValor = dUltNFBaseSimples
    
    ObterBaseICMSSaidaSimples = lErro
    
End Function

Public Function ObterCredICMSSaidaSimples(dValor As Double, ByVal lNumIntNF As Long, ByVal dtDataEmissaoNF As Date) As Long
'retorna em dValor o credito do ICMS para emissao de nf de venda para empresas no simples

Dim lErro As Long, dAliquota As Double, dVal As Double, lComando As Long, dAcumSemSubst As Double, dAcumComSubst As Double
Dim dDespesas As Double, dValorTotal As Double, dQuantidade As Double, dPrecoUnitario As Double, dICMSSubstValor As Double
Dim sNaturezaOp As String, dtDataEmissao As Date, dValorNaoFat As Double, dValorBase As Double

On Error GoTo Erro_ObterCredICMSSaidaSimples

    If lNumIntNFCredSimples <> lNumIntNF Then
    
        dValorBase = 0
        dValor = 0
        dAcumSemSubst = 0
        dAcumComSubst = 0
        
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 81810
        
        sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        lErro = Comando_Executar(lComando, "SELECT ValorFrete+ValorSeguro+ValorOutrasDespesas, ValorTotal, Quantidade, PrecoUnitario, ICMSSubstValor, TributacaoItemNF.NaturezaOp, NFiscal.DataEmissao FROM NFiscal, ItensNFiscal, TributacaoItemNF WHERE ItensNFiscal.NumIntNF = TributacaoItemNF.NumIntNF AND ItensNFiscal.NumIntNF = ? AND ItensNFiscal.Item = TributacaoItemNF.Item AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF", _
            dDespesas, dValorTotal, dQuantidade, dPrecoUnitario, dICMSSubstValor, sNaturezaOp, dtDataEmissao, lNumIntNF)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194854
    
        Do While lErro = AD_SQL_SUCESSO
        
            Select Case sNaturezaOp
            
                Case "5901", "5902", "6901", "6902", "6925", "5925"
                    dValorNaoFat = dValorNaoFat + Arredonda_Moeda(dQuantidade * dPrecoUnitario)
                    
                Case Else
                    If dICMSSubstValor <> 0 Then
                        dAcumComSubst = dAcumComSubst + Arredonda_Moeda(dQuantidade * dPrecoUnitario)
                    Else
                        dAcumSemSubst = dAcumSemSubst + Arredonda_Moeda(dQuantidade * dPrecoUnitario)
                    End If
            
            End Select
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194855
    
        Loop
        
        Call Comando_Fechar(lComando)
        
        lErro = Obtem_AliquotaICMSDAS(IIf(dtDataEmissaoNF = DATA_NULA, dtDataEmissao, dtDataEmissaoNF), dAliquota)
        If lErro <> SUCESSO Then gError 201038
        
        If dAcumSemSubst <> 0 Then
        
            If dAcumComSubst = 0 Then
                dValorBase = Arredonda_Moeda(dValorTotal - dValorNaoFat)
            Else
                dValorBase = Arredonda_Moeda((dDespesas * (dAcumSemSubst / (dAcumSemSubst + dAcumComSubst)) + dAcumSemSubst))
            End If
            dValor = Arredonda_Moeda(dAliquota * dValorBase)
            
        End If
        
        dUltNFCredSimples = dValor
        dUltNFBaseSimples = dValorBase
        lNumIntNFCredSimples = lNumIntNF
        
    End If
    
    dValor = dUltNFCredSimples
        
    ObterCredICMSSaidaSimples = SUCESSO
     
    Exit Function
    
Erro_ObterCredICMSSaidaSimples:

    ObterCredICMSSaidaSimples = gErr
     
    Select Case gErr
          
        Case 201038
        
        Case 81810
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201037)
     
    End Select
     
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ObterCredICMSSaidaSimples2(dValor As Double, ByVal lNumIntNF As Long, ByVal dtDataEmissaoNF As Date) As Long
'retorna em dValor o credito do ICMS para emissao de nf de venda para empresas no simples

Dim lErro As Long, dAliquota As Double, dVal As Double, lComando As Long, dAcumSemSubst As Double, dAcumComSubst As Double
Dim dDespesas As Double, dValorTotal As Double, dQuantidade As Double, dPrecoUnitario As Double, dICMSSubstValor As Double
Dim sNaturezaOp As String, dtDataEmissao As Date, dValorNaoFat As Double, dValorBase As Double

On Error GoTo Erro_ObterCredICMSSaidaSimples2

    If lNumIntNFCredSimples <> lNumIntNF Then
    
        dValorBase = 0
        dValor = 0
        dAcumSemSubst = 0
        dAcumComSubst = 0
        
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 81810
        
        sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        lErro = Comando_Executar(lComando, "SELECT ValorFrete+ValorSeguro+ValorOutrasDespesas, ValorTotal, Quantidade, PrecoUnitario, ICMSSubstValor, TributacaoItemNF.NaturezaOp, NFiscal.DataEmissao FROM NFiscal, ItensNFiscal, TributacaoItemNF WHERE ItensNFiscal.NumIntNF = TributacaoItemNF.NumIntNF AND ItensNFiscal.NumIntNF = ? AND ItensNFiscal.Item = TributacaoItemNF.Item AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF", _
            dDespesas, dValorTotal, dQuantidade, dPrecoUnitario, dICMSSubstValor, sNaturezaOp, dtDataEmissao, lNumIntNF)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194854
    
        Do While lErro = AD_SQL_SUCESSO
        
            Select Case sNaturezaOp
            
                Case "5101", "6101", "5124", "6124"
                    If dICMSSubstValor <> 0 Then
                        dAcumComSubst = dAcumComSubst + Arredonda_Moeda(dQuantidade * dPrecoUnitario)
                    Else
                        dAcumSemSubst = dAcumSemSubst + Arredonda_Moeda(dQuantidade * dPrecoUnitario)
                    End If
            
                Case Else
                    dValorNaoFat = dValorNaoFat + Arredonda_Moeda(dQuantidade * dPrecoUnitario)
                    
            End Select
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194855
    
        Loop
        
        Call Comando_Fechar(lComando)
        
        lErro = Obtem_AliquotaICMSDAS(IIf(dtDataEmissaoNF = DATA_NULA, dtDataEmissao, dtDataEmissaoNF), dAliquota)
        If lErro <> SUCESSO Then gError 201038
        
        If dAcumSemSubst <> 0 Then
        
            If dAcumComSubst = 0 Then
                dValorBase = Arredonda_Moeda(dValorTotal - dValorNaoFat)
            Else
                dValorBase = Arredonda_Moeda((dDespesas * (dAcumSemSubst / (dAcumSemSubst + dAcumComSubst)) + dAcumSemSubst))
            End If
            dValor = Arredonda_Moeda(dAliquota * dValorBase)
            
        End If
        
        dUltNFCredSimples = dValor
        dUltNFBaseSimples = dValorBase
        lNumIntNFCredSimples = lNumIntNF
        
    End If
    
    dValor = dUltNFCredSimples
        
    ObterCredICMSSaidaSimples2 = SUCESSO
     
    Exit Function
    
Erro_ObterCredICMSSaidaSimples2:

    ObterCredICMSSaidaSimples2 = gErr
     
    Select Case gErr
          
        Case 201038
        
        Case 81810
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201037)
     
    End Select
     
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ObterBaseICMSSaidaSimples2(dValor As Double, ByVal lNumIntNF As Long) As Long
Dim dVal As Double, lErro As Long
    
    lErro = ObterCredICMSSaidaSimples2(dVal, lNumIntNF, DATA_NULA)
    If lErro = SUCESSO Then dValor = dUltNFBaseSimples
    
    ObterBaseICMSSaidaSimples2 = lErro
    
End Function

Public Function ObterTransp(sRetorno As String, iCodTransp As Integer) As Long
'Para obter o Código da Situação Tributária

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_ObterTransp

    sRetorno = 0

    If iCodTransp <> 0 Then

        objTransportadora.iCodigo = iCodTransp

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM

        If lErro <> SUCESSO Then gError 208095
        
        'Carrega lEndereco em objTransportadora
        objEndereco.lCodigo = objTransportadora.lEndereco
    
        'Lê o endereço a partir do Código
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
    
        If lErro = 12309 Then gError 208096
        
        sRetorno = "Redespacho: " + objTransportadora.sNome + " - " + objEndereco.sEndereco + " - " + objEndereco.sBairro + " - " + objEndereco.sCidade + " - " + objEndereco.sSiglaEstado + " - " + Format(objEndereco.sCEP, "#####-###")

    End If
    
    ObterTransp = SUCESSO
    
    Exit Function
    
Erro_ObterTransp:

    ObterTransp = gErr
    
    Select Case gErr

        Case 208095
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr)

        Case 208096
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208097)

    End Select

    Exit Function

End Function

Public Function Obter_Campo_Danfe(sRetorno As String, ByVal lNumIntItemNF As Long, ByVal iTipoCampo As Integer, ByVal sNomeCampo As String) As Long

Dim lErro As Long, lComando As Long, vCampo As Variant
Dim sUF As String, sUF1 As String, sUF2 As String, iFilialEmpresa As Integer, lComando2 As Long

On Error GoTo Erro_Obter_Campo_Danfe

    sRetorno = ""
          
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 208291
    
    Select Case sNomeCampo
    
        Case "FilEmpIEST"
               
            sUF1 = String(STRING_ESTADO_SIGLA, 0)
            sUF2 = String(STRING_ESTADO_SIGLA, 0)
    
            lErro = Comando_Executar(lComando, "SELECT FilialCliEstado, FilialFornSiglaEstado, NFFilialEmpresa FROM Danfe WHERE ItemNFNumIntDoc = ?", sUF1, sUF2, iFilialEmpresa, lNumIntItemNF)
            If lErro <> AD_SQL_SUCESSO Then gError 208292
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208293
    
            If lErro = AD_SQL_SUCESSO Then
            
                lComando2 = Comando_AbrirExt(GL_lConexaoRel)
                If lComando2 = 0 Then gError 208291
            
                sUF = IIf(Len(Trim(sUF1)) = 0, sUF2, sUF1)
            
                vCampo = String(255, 0)
            
                lErro = Comando_Executar(lComando2, "SELECT InscricaoEstadual FROM ContribuinteSubstituto WHERE FilialEmpresa = ? AND UF = ? ", vCampo, iFilialEmpresa, sUF)
                If lErro <> AD_SQL_SUCESSO Then gError 208292
            
                lErro = Comando_BuscarProximo(lComando2)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208293
            
                If lErro = AD_SQL_SUCESSO Then
                    sRetorno = vCampo
                End If
            
                Call Comando_Fechar(lComando2)
            
            End If
       
        Case Else

            '1=SmallInt 2=Integer 3=Double 4=VarChar 6=Date
            Select Case iTipoCampo
                Case 1
                    vCampo = CInt(0)
                Case 2
                    vCampo = CLng(0)
                Case 3
                    vCampo = CDbl(0)
                Case 4
                    vCampo = String(255, 0)
                Case 5
                    vCampo = CDbl(0)
                Case 6
                    vCampo = DATA_NULA
            End Select
                
            lErro = Comando_Executar(lComando, "SELECT " & sNomeCampo & " FROM Danfe WHERE ItemNFNumIntDoc = ?", vCampo, lNumIntItemNF)
            If lErro <> AD_SQL_SUCESSO Then gError 208292
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208293
                
            If lErro = AD_SQL_SUCESSO Then
                
                Select Case iTipoCampo
                    Case 1, 2
                        sRetorno = CStr(vCampo)
                    Case 3
                        sRetorno = Format(vCampo, "STANDARD")
                    Case 4
                        sRetorno = vCampo
                    Case 5
                        sRetorno = Format(CDate(vCampo), "HH:MM")
                    Case 6
                        If vCampo = DATA_NULA Then
                            sRetorno = ""
                        Else
                            sRetorno = Format(vCampo, "dd/mm/yyyy")
                        End If
                End Select
            
            End If
            
    End Select
    
        
    Call Comando_Fechar(lComando)
        
    Obter_Campo_Danfe = SUCESSO
     
    Exit Function
    
Erro_Obter_Campo_Danfe:

    Obter_Campo_Danfe = gErr
     
    Select Case gErr
       
        Case 208291
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 208292, 208293
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208294)
     
    End Select
     
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    If lComando2 <> 0 Then Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Function ObterPaginasDanfe(lNumIntDoc As Long, lNumeroMaximoPaginaPri As Long, lNumeroMaximoPaginaOut As Long, sPagina As String) As Long
'Para obter o Número de Páginas

Dim lErro As Long
Dim lPaginaTemp As Double
Dim iResto As Integer, iIndice As Integer
Dim iNumItens As Integer
Dim iNumTotalAreasDet As Integer
Dim objNFiscal As New ClassNFiscal
Dim iNumQuebrasExibidas As Integer
Dim objItemNF As ClassItemNF
Dim objMsg As New ClassRelMsg
Dim sAux As String

On Error GoTo Erro_ObterPaginasDanfe

    If lNumIntDoc <> lUltNumIntNFPagDanfe Then
    
        lNumChamadasDanfe = 0
    
        objNFiscal.lNumIntDoc = lNumIntDoc
    
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        iNumItens = objNFiscal.colItens.Count
        
        'Atualmente se for maior que 38 exibe na linha seguinte
        'Se mudar o tsk tem que mudar essa rotina
        For Each objItemNF In objNFiscal.colItens
        
            If Len(objItemNF.sDescricaoItem) > 50 Then
                iNumQuebrasExibidas = iNumQuebrasExibidas + 1
            End If
            
            'Se mudar o tamanho da mensagem exibida por linha também tem que alterar essa função
            'Se mudar o número de linha também
            'Para cada linha que pode ser exibida
            For iIndice = 1 To 3
                sAux = ""
                lErro = objMsg.ObterMsg(sAux, REGRAMSG_TIPODOC_ITEMNF, objItemNF.lNumIntDoc, REGRAMSG_TIPOMSG_NORMAL, 150, iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If sAux <> "" Then
                    iNumQuebrasExibidas = iNumQuebrasExibidas + 1
                End If
            Next
            
        Next
        
        'Se mudar o tamanho da mensagem exibida por linha também tem que alterar essa função
        'Se mudar o número de linha também
        'Para cada linha que pode ser exibida
        For iIndice = 1 To 5
            sAux = ""
            lErro = objMsg.ObterMsg(sAux, REGRAMSG_TIPODOC_NF, objNFiscal.lNumIntDoc, REGRAMSG_TIPOMSG_CORPO, 150, iIndice)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If sAux <> "" Then
                iNumQuebrasExibidas = iNumQuebrasExibidas + 1
            End If
        Next
        
        iNumTotalAreasDet = iNumItens + iNumQuebrasExibidas
        
        'Se cabe tudo na primeira página
        If lNumeroMaximoPaginaPri > iNumTotalAreasDet Then
            lPaginaTemp = 1
        Else
        'Senão abate o que já foi na primeira e calcula com o máximo das demais
            iNumTotalAreasDet = iNumTotalAreasDet - lNumeroMaximoPaginaPri
        
            iResto = Resto(CDbl(iNumTotalAreasDet), CInt(lNumeroMaximoPaginaOut))
        
            If iResto = 0 Then
                lPaginaTemp = iNumTotalAreasDet / lNumeroMaximoPaginaOut + 1
            Else
                lPaginaTemp = Fix(iNumTotalAreasDet / lNumeroMaximoPaginaOut) + 2
            End If
        
        End If
        
        lUltNumIntNFPagDanfe = lNumIntDoc
        lPaginaAcumuladaDanfe = lPaginaTemp
    
    End If
    
    lNumChamadasDanfe = lNumChamadasDanfe + 1
    
    sPagina = Format(lNumChamadasDanfe, "00") & SEPARADOR & Format(lPaginaAcumuladaDanfe, "00")
    
    ObterPaginasDanfe = SUCESSO
    
    Exit Function
    
Erro_ObterPaginasDanfe:

    ObterPaginasDanfe = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153291)

    End Select

    Exit Function

End Function

Public Function Obter_Campo_Config(sRetorno As String, ByVal sTabela As String, ByVal sCodigo As String) As Long

Dim lErro As Long, lComando As Long, sConteudo As String

On Error GoTo Erro_Obter_Campo_Config

    sRetorno = ""
          
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 208403
    
    sConteudo = String(255, 0)
        
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM " & sTabela & " WHERE Codigo = ? AND FilialEmpresa = ? ", sConteudo, sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 208404

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208405
        
    If lErro <> AD_SQL_SUCESSO Then
        
        sConteudo = String(255, 0)
        
        lErro = Comando_Executar(lComando, "SELECT Conteudo FROM " & sTabela & " WHERE Codigo = ? AND FilialEmpresa = ? ", sConteudo, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then gError 208406
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208407
        
    End If
    
    sRetorno = sConteudo
        
    Call Comando_Fechar(lComando)
        
    Obter_Campo_Config = SUCESSO
     
    Exit Function
    
Erro_Obter_Campo_Config:

    Obter_Campo_Config = gErr
     
    Select Case gErr
       
        Case 208403
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 208404 To 208407
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", gErr, sTabela, sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208408)
     
    End Select
     
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function NF_Possui_ParcRec_Bordero(dFlag As Double, ByVal lNumIntNF As Long, ByVal lNumBordero As Long) As Long

Dim lErro As Long, lComando As Long, lAux As Long
Dim sSQL As String

On Error GoTo Erro_NF_Possui_ParcRec_Bordero
          
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 211684
    
    sSQL = "SELECT T.NumIntDoc FROM ParcelasRecTodas AS P, TitulosRecTodos AS T, NFiscal As N, OcorrenciasRemParcRec AS O"
    sSQL = sSQL & " " & "WHERE N.ClasseDocCPR = 2 AND T.NumIntDoc = N.NumIntDocCPR AND T.NumIntDoc = P.NumIntTitulo AND"
    sSQL = sSQL & " " & "T.Status <> 5 AND O.NumIntParc = P.NumIntDoc AND O.NumBordero = ? AND N.NumIntDoc = ?"
        
    lErro = Comando_Executar(lComando, sSQL, lAux, lNumBordero, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 211685

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211686
        
    If lErro <> AD_SQL_SUCESSO Then
        dFlag = 0
    Else
        dFlag = 1
    End If
        
    Call Comando_Fechar(lComando)
        
    NF_Possui_ParcRec_Bordero = SUCESSO
     
    Exit Function
    
Erro_NF_Possui_ParcRec_Bordero:

    NF_Possui_ParcRec_Bordero = gErr
     
    Select Case gErr
       
        Case 211684
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211685 To 211686
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211687)
     
    End Select
     
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ParcRec_Obtem_Bordero(dNumBordero As Double, ByVal lNumIntParc As Long) As Long

Dim lErro As Long, lComando As Long, lAux As Long
Dim sSQL As String

On Error GoTo Erro_ParcRec_Obtem_Bordero
          
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 211688
    
    sSQL = "SELECT NumBordero FROM OcorrenciasRemParcRec WHERE CodOcorrencia = 1 AND NumIntParc = ? ORDER BY NumIntDoc DESC"
        
    lErro = Comando_Executar(lComando, sSQL, lAux, lNumIntParc)
    If lErro <> AD_SQL_SUCESSO Then gError 211689

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211690
        
    If lErro <> AD_SQL_SUCESSO Then
        dNumBordero = 0
    Else
        dNumBordero = lAux
    End If
        
    Call Comando_Fechar(lComando)
        
    ParcRec_Obtem_Bordero = SUCESSO
     
    Exit Function
    
Erro_ParcRec_Obtem_Bordero:

    ParcRec_Obtem_Bordero = gErr
     
    Select Case gErr
       
        Case 211688
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211689 To 211690
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OCORRENCIASREMPARCREC", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211691)
     
    End Select
     
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ObterDataParcRecNFe(dtDataParcRec As Date, ByVal schNFe As String, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter a data de vencimento da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterDataParcRecNFe

    If schNFe <> sUltchNFe Then
        
        Set colParcRec = New ColParcelaReceber
        
        If lNumIntTitulo <> 0 Then
            
            lErro = CF("ParcelasRec_Le_DataValor_NFe", schNFe, lNumIntTitulo, colParcRec, lComandoParcRec1, lComandoParcRec2)
            If lErro <> SUCESSO Then gError 58223
        
        End If
        
        sUltchNFe = schNFe
        
    End If
    
    If iParcela <= colParcRec.Count Then
        dtDataParcRec = colParcRec(iParcela).dtDataVencimento
    Else
        dtDataParcRec = DATA_NULA
    End If
    
    
    ObterDataParcRecNFe = SUCESSO
    
    Exit Function
    
Erro_ObterDataParcRecNFe:

    ObterDataParcRecNFe = gErr
    
    Select Case gErr

        Case 58223 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153377)

    End Select

    Exit Function

End Function

Public Function ObterValorParcRecNFe(dValor As Double, ByVal schNFe As String, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterValorParcRecNFe

    If schNFe <> sUltchNFe Then
        
        Set colParcRec = New ColParcelaReceber
        
        If lNumIntTitulo <> 0 Then
        
            lErro = CF("ParcelasRec_Le_DataValor_NFe", schNFe, lNumIntTitulo, colParcRec, lComandoParcRec1, lComandoParcRec2)
            If lErro <> SUCESSO Then gError 58224
            
        End If
        
        sUltchNFe = schNFe
        
    End If
    
    If iParcela <= colParcRec.Count Then
        dValor = colParcRec(iParcela).dValorOriginal
    Else
        dValor = 0
    End If
        
    ObterValorParcRecNFe = SUCESSO
    
    Exit Function
    
Erro_ObterValorParcRecNFe:

    ObterValorParcRecNFe = gErr
    
    Select Case gErr

        Case 58224 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153378)

    End Select

    Exit Function

End Function
