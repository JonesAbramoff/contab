VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMapaCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Criada para atender a mapa de cotacao

Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long

Dim sProdAnt As String
Dim iQtdFornAnt As Integer

Dim sProdAntPC As String
Dim iQtdPCAnt As Integer

Dim sProdAntRC As String
Dim iQtdRCAnt As Integer

Dim colFornAnt As Collection
Dim colPCAnt As Collection
Dim colRCAnt As Collection

'------------------------------------------------------------------------
'CODIGO DEFAULT DAS CLASSES DE RELATORIO
'------------------------------------------------------------------------
Private Sub Class_Terminate()

    If lComando1 <> 0 Then Call Comando_Fechar(lComando1)
    If lComando2 <> 0 Then Call Comando_Fechar(lComando2)
    If lComando3 <> 0 Then Call Comando_Fechar(lComando3)
    If lComando4 <> 0 Then Call Comando_Fechar(lComando4)
        
End Sub
'------------------------------------------------------------------------


'------------------------------------------------------------------------
'TRATAMENTO DE FORNECEDORES DO RELATORIO
'------------------------------------------------------------------------
Public Function MapaCot_ObterFornCod(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterFornCod

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111864
    
    If Not (objForn Is Nothing) Then
        sTexto = objForn.lCodigo
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterFornCod = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterFornCod:

    MapaCot_ObterFornCod = gErr
    
    Select Case gErr
        
        Case 111864
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153192)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornContato(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterFornContato

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111863
    
    If Not (objForn Is Nothing) Then
        sTexto = objForn.sContato
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterFornContato = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterFornContato:

    MapaCot_ObterFornContato = gErr
    
    Select Case gErr
        
        Case 111863
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153193)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornNome(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterFornNome

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111862
    
    If Not (objForn Is Nothing) Then
        sTexto = objForn.sNome
    Else
        sTexto = ""
    End If
        
    MapaCot_ObterFornNome = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterFornNome:

    MapaCot_ObterFornNome = gErr
    
    Select Case gErr
        
        Case 111862
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153194)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornTel1(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterFornTel1

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111861
    
    If Not (objForn Is Nothing) Then
        sTexto = objForn.sTel1
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterFornTel1 = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterFornTel1:

    MapaCot_ObterFornTel1 = gErr
    
    Select Case gErr
        
        Case 111861
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153195)

    End Select

    Exit Function

End Function

 Public Function MapaCot_ObterFornTel2(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterFornTel2

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111860
    
    If Not (objForn Is Nothing) Then
        sTexto = objForn.sTel2
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterFornTel2 = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterFornTel2:

    MapaCot_ObterFornTel2 = gErr
    
    Select Case gErr
        
        Case 111860
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153196)

    End Select

    Exit Function

End Function

Public Function ObterDadosFornMapaCotacao(ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer, objForn As ClassFornMapaCot) As Long

Dim lErro As Long

On Error GoTo Erro_ObterDadosFornMapaCotacao
 
    'se o produto passado como parametro mudou
    If sProduto <> sProdAnt Or iQtdFornAnt <> iQtd Then
               
        'le as informacoes relevantes dos fornecedores do produto passado como parametro
        lErro = DadosForn_Le_Produto_PedCot(sProduto, lNumIntMapaCot, iQtd, colFornAnt, lComando1, lComando2)
        If lErro <> SUCESSO Then gError 111860
                        
        'seta como corrente o novo produto
        sProdAnt = sProduto
        iQtdFornAnt = iQtd
    
    End If
    
    'Retorna o fornecedor
    If iSeq > colFornAnt.Count Then
        Set objForn = Nothing
    Else
        Set objForn = colFornAnt(iSeq)
    End If
        
    ObterDadosFornMapaCotacao = SUCESSO

    Exit Function

Erro_ObterDadosFornMapaCotacao:

    ObterDadosFornMapaCotacao = gErr

    Select Case gErr
        
        Case 111860
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153197)

    End Select

    Exit Function

End Function

Function DadosForn_Le_Produto_PedCot(sCodProduto As String, lNumIntMapaCot As Long, iQtdLeitura As Integer, colFornecedores As Collection, lComando1 As Long, lComando2 As Long) As Long
'Le os fornecedores do produto em questao
'A qtde de fornecedores retornadas na colecao nao ultrapassara o valor contido em iQtdLeitura
'Essa funcao garante que o primeiro fornecedor da colecao sera o fornecedor do pedido de cotacao em questao, caso o mesmo exista

'sCodProduto (INPUT) eh o filtro do produto para a leitura
'lNumIntPedCot (INPUT) eu vou utilizar para garantir que o primeiro dos fornecedores
'seja o fornecedor do pedido em questao
'iQtdLeitura (INPUT) eh a qtd de fornecedores que eu vou ler do produto em questao
'pode ser que a quantidade seja menor, caso existam menos fornecedores
'mantem os indices da colecao relacionados com o sequencial do pedido
'lComando1 e 2 (INPUT)

Dim lErro As Long, dQuantidade As Double, lCodPedCotacao As Long
Dim colForn As New Collection
Dim sNome, sTel1, sTel2, sContato As String
Dim lCodigo As Long
Dim iIndice As Integer
Dim lFornecedor As Long
Dim iFilial As Integer
Dim bAchou As Boolean
Dim objFornMapaCot As ClassFornMapaCot

On Error GoTo Erro_DadosForn_Le_Produto_PedCot

    'se o comando ainda nao foi aberto... abre...
    If lComando1 = 0 Then
        lComando1 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando1 = 0 Then gError 111850
    End If
    
    If lComando2 = 0 Then
        lComando2 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando2 = 0 Then gError 111850
    End If
    
    'inicializa as strings
    sNome = String(STRING_BUFFER_MAX_TEXTO, 0)
    sTel1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    sTel2 = String(STRING_BUFFER_MAX_TEXTO, 0)
    sContato = String(STRING_BUFFER_MAX_TEXTO, 0)

    'obtem o conjunto de fornecedores associados ao produto
    lErro = Comando_Executar(lComando1, "SELECT MapaCot_FornProdAux.Quantidade, MapaCot_FornProdAux.CodPedCotacao, MapaCot_FornProd.Codigo, MapaCot_FornProd.Nome, MapaCot_FornProd.Telefone1, MapaCot_FornProd.Telefone2, MapaCot_FornProd.Contato, MapaCot_FornProd.CodFornecedor, MapaCot_FornProd.CodFilial FROM MapaCot_FornProd LEFT OUTER JOIN MapaCot_FornProdAux ON " & _
            " (MapaCot_FornProd.CodFornecedor = MapaCot_FornProdAux.Fornecedor AND MapaCot_FornProd.CodFilial = MapaCot_FornProdAux.Filial AND MapaCot_FornProd.Produto = MapaCot_FornProdAux.Produto) " & _
            " WHERE MapaCot_FornProd.Produto = ? AND (MapaCot_FornProdAux.NumIntDoc IS NULL OR MapaCot_FornProdAux.NumIntDoc = ?) AND MapaCot_FornProd.FilialEmpresa = ? ORDER BY MapaCot_FornProdAux.Produto DESC, CodPedCotacao", dQuantidade, lCodPedCotacao, lCodigo, sNome, sTel1, sTel2, sContato, lFornecedor, iFilial, sCodProduto, lNumIntMapaCot, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 111851
    
    'le o primeiro fornecedor
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111852
    
    'enquanto for achando fornecedores
    Do While lErro = AD_SQL_SUCESSO
    
        'instancia novo obj
        Set objFornMapaCot = New ClassFornMapaCot
        
        'preenche o obj com os dados lidos
        objFornMapaCot.lCodigo = lCodigo
        objFornMapaCot.sNome = sNome
        objFornMapaCot.sTel1 = sTel1
        objFornMapaCot.sTel2 = sTel2
        objFornMapaCot.sContato = sContato
        objFornMapaCot.lCodFornecedor = lFornecedor
        objFornMapaCot.iCodFilial = iFilial
        objFornMapaCot.dQuantidade = dQuantidade
        objFornMapaCot.lCodPedCotacao = lCodPedCotacao
        
        'se ainda nao leu iQtdLeitura fornecedores
        If iIndice < iQtdLeitura Then
        
            'adiciona o fornecedor na colecao principal
            colForn.Add objFornMapaCot
            
        'se ja leu, para de ler registros
        Else
        
            Exit Do
        
        End If
    
        'busca o proximo fornecedor
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111853
        
        'incrementa o contador de fornecedores lidos
        iIndice = iIndice + 1
    
    Loop
    
    'retorna a colecao de fornecedores
    Set colFornecedores = colForn
        
    DadosForn_Le_Produto_PedCot = SUCESSO
    
    Exit Function
 
Erro_DadosForn_Le_Produto_PedCot:

    DadosForn_Le_Produto_PedCot = gErr

    Select Case gErr
    
        Case 111850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 111851, 111852, 111853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOT_FORNPROD", gErr, sCodProduto)
    
        Case 111854, 111855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOT_PEDIDOCOTACAON", gErr, lNumIntMapaCot)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153198)
    
    End Select
    
    Exit Function

End Function
'------------------------------------------------------------------------


'------------------------------------------------------------------------
'TRATAMENTO DE PEDIDO DE COMPRA DO RELATORIO
'------------------------------------------------------------------------

Public Function MapaCot_ObterPCCod(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCCod

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111875
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.lCodigo
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCCod = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCCod:

    MapaCot_ObterPCCod = gErr
    
    Select Case gErr
        
        Case 111875
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153199)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCData(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCData

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111876
    
    If Not (objPC Is Nothing) Then
        sTexto = IIf(objPC.dtData <> DATA_NULA, objPC.dtData, "")
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCData = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCData:

    MapaCot_ObterPCData = gErr
    
    Select Case gErr
        
        Case 111876
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153200)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCForn(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCForn

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111877
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.sFornecedor
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCForn = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCForn:

    MapaCot_ObterPCForn = gErr
    
    Select Case gErr
        
        Case 111877
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153201)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCQuant(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCQuant

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111878
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.dQuantidade
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCQuant = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCQuant:

    MapaCot_ObterPCQuant = gErr
    
    Select Case gErr
        
        Case 111878
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153202)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCPreco(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCPreco

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111879
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.dPreco
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCPreco = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCPreco:

    MapaCot_ObterPCPreco = gErr
    
    Select Case gErr
        
        Case 111879
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153203)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCDesconto(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCDesconto

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111880
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.dDesconto
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCDesconto = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCDesconto:

    MapaCot_ObterPCDesconto = gErr
    
    Select Case gErr
        
        Case 1118880
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153204)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCIPI(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCIPI

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111881
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.dIPI
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCIPI = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCIPI:

    MapaCot_ObterPCIPI = gErr
    
    Select Case gErr
        
        Case 111881
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153205)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCCondPagto(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCCondPagto

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111882
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.sCondPagto
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCCondPagto = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCCondPagto:

    MapaCot_ObterPCCondPagto = gErr
    
    Select Case gErr
        
        Case 111882
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153206)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCTaxaMoeda(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCTaxaMoeda

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111883
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.dTaxa
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCTaxaMoeda = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCTaxaMoeda:

    MapaCot_ObterPCTaxaMoeda = gErr
    
    Select Case gErr
        
        Case 111883
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153207)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCDataEntrega(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCDataEntrega

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111884
    
    If Not (objPC Is Nothing) Then
        sTexto = IIf(objPC.dtDataEntrega <> DATA_NULA, objPC.dtDataEntrega, "")
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCDataEntrega = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCDataEntrega:

    MapaCot_ObterPCDataEntrega = gErr
    
    Select Case gErr
        
        Case 111884
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153208)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCStatus(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCStatus

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError 111885
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.sStatus
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCStatus = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCStatus:

    MapaCot_ObterPCStatus = gErr
    
    Select Case gErr
        
        Case 111885
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153209)

    End Select

    Exit Function

End Function

Public Function ObterDadosPCMapaCotacao(ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer, objPC As ClassPCMapaCot) As Long

Dim lErro As Long

On Error GoTo Erro_ObterDadosPCMapaCotacao
 
    'se o produto passado como parametro mudou
    If sProduto <> sProdAntPC Or iQtdPCAnt <> iQtd Then
               
        'le as informacoes relevantes dos pedidos de compra do produto passado como parametro
        lErro = DadosPC_Le_Produto(sProduto, iQtd, colPCAnt, lComando3)
        If lErro <> SUCESSO Then gError 120020
                        
        'seta como corrente o novo produto
        sProdAntPC = sProduto
        iQtdPCAnt = iQtd
    
    End If
    
    'Retorna o pc
    If iSeq > colPCAnt.Count Then
        Set objPC = Nothing
    Else
        Set objPC = colPCAnt(iSeq)
    End If
        
    ObterDadosPCMapaCotacao = SUCESSO

    Exit Function

Erro_ObterDadosPCMapaCotacao:

    ObterDadosPCMapaCotacao = gErr

    Select Case gErr
        
        Case 120020
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153210)

    End Select

    Exit Function

End Function

Function DadosPC_Le_Produto(sCodProduto As String, iQtdLeitura As Integer, colPC As Collection, lComando3 As Long) As Long
'Le os pedidos de compra do produto em questao retornando iQtd ou menos pedidos na colecao passada como parametro
'sCodProduto (INPUT) traz o produto que eh o filtro para a leitura dos pedidos
'iQtdLeitura (INPUT) eh a qtde de pedidos que serao armazenados na colecao
'colRC (OUTPUT) eh a colecao de retorno, que retornara iQtd ou menos pcs lidos
'lComando3 (INPUT) serah utilizado para as operacoes de bd


Dim lErro As Long
Dim colPedCompra As New Collection
Dim dDesconto As Double
Dim lCodigo As Long
Dim dQuantidade As Double, dPreco As Double, dIPI As Double, dTaxa As Double
Dim dtData As Date, dtDataEntrega As Date
Dim sCondPagto As String, sStatus As String, sFornecedor As String
Dim iStatus As Integer

Dim iIndice As Integer

'???tuliodebug.. a quantidade nao deveria ser double????????
'ja acertado.. estava dando erro...
Dim iQuantidade As Integer

Dim objPCMapaCot As ClassPCMapaCot
Dim dValorFrete As Double, dValorProdutos As Double, sTipoFrete As String

On Error GoTo Erro_DadosPC_Le_Produto

    'se a conexao ainda nao foi aberta... abre...
    If lComando3 = 0 Then
        lComando3 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando3 = 0 Then gError 120014
    End If
    
    'inicializa as strings
    sCondPagto = String(STRING_BUFFER_MAX_TEXTO, 0)
    sFornecedor = String(STRING_BUFFER_MAX_TEXTO, 0)
    sTipoFrete = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'obtem o conjunto de pedidos associados ao produto
    lErro = Comando_Executar(lComando3, "SELECT Codigo, Data, Fornecedor, Quantidade, PrecoUnitario, ValorDesconto, ValorIPI, CondPagto, Taxa, DataEntrega, Status, ValorFrete, ValorProdutos, TipoFrete FROM MapaCot_PedCompra WHERE Produto = ? ORDER BY Data DESC, NumIntDoc DESC", lCodigo, dtData, sFornecedor, dQuantidade, dPreco, dDesconto, dIPI, sCondPagto, dTaxa, dtDataEntrega, iStatus, dValorFrete, dValorProdutos, sTipoFrete, sCodProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 120015
    
    'le o primeiro pedido
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 120016
    
    'enquanto for achando pedidos
    Do While lErro = AD_SQL_SUCESSO
    
        'instancia novo obj
        Set objPCMapaCot = New ClassPCMapaCot
        
        'preenche o obj com os dados lidos
        With objPCMapaCot
        
            .dIPI = dIPI
            .dPreco = dPreco
            .dQuantidade = dQuantidade
            .dTaxa = dTaxa
            .dtData = dtData
            .dtDataEntrega = dtDataEntrega
            .lCodigo = lCodigo
            .sFornecedor = sFornecedor
            .sCondPagto = sCondPagto
            .dDesconto = dDesconto
            .sStatus = IIf(iStatus = 1, STRING_STATUS_BAIXADO, STRING_STATUS_PENDENTE)
            
            If sTipoFrete = "0" Then
                .sFreteTipo = "CIF"
            ElseIf sTipoFrete = "1" Then
                .sFreteTipo = "FOB"
            Else
                .sFreteTipo = ""
            End If
            
            .dFrete = (dValorFrete / dValorProdutos) * dPreco
                    
        End With
                
        'se ainda nao leu iQtdLeitura pedidos
        If iIndice < iQtdLeitura Then
        
            'adiciona o pedido na colecao principal
            colPedCompra.Add objPCMapaCot
            
        'se ja leu, sai do loop
        Else
        
            Exit Do
        
        End If
    
        'busca o proximo pedido
        lErro = Comando_BuscarProximo(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 120017
        
        'incrementa o contador de pedidos lidos
        iIndice = iIndice + 1
    
    Loop
    
    'retorna a colecao de pedidos
    Set colPC = colPedCompra
        
    DadosPC_Le_Produto = SUCESSO
    
    Exit Function
 
Erro_DadosPC_Le_Produto:

    DadosPC_Le_Produto = gErr

    Select Case gErr
    
        Case 120014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 120015 To 120017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOT_PEDCOMPRA", gErr, sCodProduto)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153211)
    
    End Select
    
    Exit Function

End Function



'------------------------------------------------------------------------




'------------------------------------------------------------------------
'TRATAMENTO DE REQUISICAO DE COMPRA DO RELATORIO
'------------------------------------------------------------------------
Public Function MapaCot_ObterRCCodigo(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCCodigo

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111886
    
    If Not (objRC Is Nothing) Then
        sTexto = objRC.lCodigo
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCCodigo = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCCodigo:

    MapaCot_ObterRCCodigo = gErr
    
    Select Case gErr
        
        Case 111886
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153212)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCRequisitante(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCRequisitante

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111887
    
    If Not (objRC Is Nothing) Then
        sTexto = objRC.sRequisitante
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCRequisitante = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCRequisitante:

    MapaCot_ObterRCRequisitante = gErr
    
    Select Case gErr
        
        Case 111887
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153213)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCCCL(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCCCL

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111888
    
    If Not (objRC Is Nothing) Then
        sTexto = objRC.sCcl
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCCCL = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCCCL:

    MapaCot_ObterRCCCL = gErr
    
    Select Case gErr
        
        Case 111888
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153214)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCQuant(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCQuant

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111889
    
    If Not (objRC Is Nothing) Then
        sTexto = objRC.dQuantidade
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCQuant = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCQuant:

    MapaCot_ObterRCQuant = gErr
    
    Select Case gErr
        
        Case 111889
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153215)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCPrazo(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCPrazo

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111890
    
    If Not (objRC Is Nothing) Then
        sTexto = IIf(objRC.dtPrazo <> DATA_NULA, objRC.dtPrazo, "")
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCPrazo = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCPrazo:

    MapaCot_ObterRCPrazo = gErr
    
    Select Case gErr
        
        Case 111890
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153216)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCUrgente(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCUrgente

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111891
    
    If Not (objRC Is Nothing) Then
        sTexto = objRC.sUrgente
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCUrgente = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCUrgente:

    MapaCot_ObterRCUrgente = gErr
    
    Select Case gErr
        
        Case 111891
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153217)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCObs(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objRC As ClassRCMapaCot

On Error GoTo Erro_MapaCot_ObterRCObs

    lErro = ObterDadosRCMapaCotacao(sProduto, iSeq, iQtd, objRC)
    If lErro <> SUCESSO Then gError 111892
    
    If Not (objRC Is Nothing) Then
        sTexto = objRC.sOBS
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterRCObs = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterRCObs:

    MapaCot_ObterRCObs = gErr
    
    Select Case gErr
        
        Case 111892
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153218)

    End Select

    Exit Function

End Function

Public Function ObterDadosRCMapaCotacao(ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer, objRC As ClassRCMapaCot) As Long

Dim lErro As Long

On Error GoTo Erro_ObterDadosRCMapaCotacao
 
    'se o produto passado como parametro mudou
    If sProduto <> sProdAntPC Or iQtdRCAnt <> iQtd Then
               
        'le as informacoes relevantes das requisicoes de compra do produto passado como parametro
        lErro = DadosRC_Le_Produto(sProduto, iQtd, colRCAnt, lComando4)
        If lErro <> SUCESSO Then gError 120021
                        
        'seta como corrente o novo produto
        sProdAntRC = sProduto
        iQtdRCAnt = iQtd
    
    End If
    
    'Retorna a rc
    If iSeq > colRCAnt.Count Then
        Set objRC = Nothing
    Else
        Set objRC = colRCAnt(iSeq)
    End If
        
    ObterDadosRCMapaCotacao = SUCESSO

    Exit Function

Erro_ObterDadosRCMapaCotacao:

    ObterDadosRCMapaCotacao = gErr

    Select Case gErr
        
        Case 120021
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153219)

    End Select

    Exit Function

End Function

Function DadosRC_Le_Produto(sCodProduto As String, iQtdLeitura As Integer, colRC As Collection, lComando4 As Long) As Long
'Le as requisicoes de compra do produto em questao retornando iQtd ou menos Requisicoes na colecao passada como parametro
'sCodProduto (INPUT) traz o produto que eh o filtro para a leitura das requisicoes
'iQtdLeitura (INPUT) eh a qtde de requisicoes que serao armazenadas na colecao
'colRC (OUTPUT) eh a colecao de retorno, que retornara iQtd ou menos requisicoes lidas
'lComando4 (INPUT) serah utilizado para as operacoes de bd

Dim lErro As Long
Dim colReqCompra As New Collection
Dim sCcl As String, sOBS As String
Dim lCodigo As Long
Dim dQuantidade As Double
Dim dtPrazo As Date
Dim sRequisitante As String
Dim iStatus As Integer
Dim iUrgente As Integer

Dim iIndice As Integer

Dim objRCMapaCot As ClassRCMapaCot

On Error GoTo Erro_DadosRC_Le_Produto

    'se a conexao ainda nao foi aberta... abre...
    If lComando4 = 0 Then
        lComando4 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando4 = 0 Then gError 120022
    End If
    
    'inicializa as strings
    sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
    sOBS = String(STRING_BUFFER_MAX_TEXTO, 0)
    sRequisitante = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'obtem o conjunto de requisicoes associadas ao produto
    lErro = Comando_Executar(lComando4, "SELECT codigo, nome, ccl, quantidade, observacao, urgente, datalimite FROM MapaCot_ReqCompra WHERE Produto = ? ORDER BY Data DESC", lCodigo, sRequisitante, sCcl, dQuantidade, sOBS, iUrgente, dtPrazo, sCodProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 120023
    
    'le a 1a rc
    lErro = Comando_BuscarPrimeiro(lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 120024
    
    'enquanto for achando requisicoes
    Do While lErro = AD_SQL_SUCESSO
    
        'instancia novo obj
        Set objRCMapaCot = New ClassRCMapaCot
        
        'preenche o obj com os dados lidos
        With objRCMapaCot
        
            .dQuantidade = dQuantidade
            .dtPrazo = dtPrazo
            .lCodigo = lCodigo
            .sCcl = sCcl
            .sOBS = sOBS
            .sRequisitante = sRequisitante
            .sUrgente = IIf(iUrgente = 1, ADM_SUBTIPO_SIM, ADM_SUBTIPO_NAO)
                    
        End With
                
        'se ainda nao leu iQtdLeitura requisicoes
        If iIndice < iQtdLeitura Then
        
            'adiciona a requisicao na colecao principal
            colReqCompra.Add objRCMapaCot
            
        'se ja leu, sai do loop
        Else
        
            Exit Do
        
        End If
    
        'busca a proxima req.
        lErro = Comando_BuscarProximo(lComando4)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 120025
        
        'incrementa o contador de rcs lidas
        iIndice = iIndice + 1
    
    Loop
    
    'retorna a colecao de rcs
    Set colRC = colReqCompra
        
    DadosRC_Le_Produto = SUCESSO
    
    Exit Function
 
Erro_DadosRC_Le_Produto:

    DadosRC_Le_Produto = gErr

    Select Case gErr
    
        Case 120022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 120023 To 120025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOT_REQCOMPRA", gErr, sCodProduto)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153220)
    
    End Select
    
    Exit Function

End Function


'------------------------------------------------------------------------
Public Function MapaCot_ObterPedCot(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterPedCot

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111864
    
    If Not (objForn Is Nothing) Then
        sTexto = IIf(objForn.lCodPedCotacao <> 0, CStr(objForn.lCodPedCotacao), "")
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPedCot = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPedCot:

    MapaCot_ObterPedCot = gErr
    
    Select Case gErr
        
        Case 111864
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153221)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPedCotQtd(sTexto As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objForn As ClassFornMapaCot

On Error GoTo Erro_MapaCot_ObterPedCotQtd

    lErro = ObterDadosFornMapaCotacao(sProduto, lNumIntMapaCot, iSeq, iQtd, objForn)
    If lErro <> SUCESSO Then gError 111864
    
    If Not (objForn Is Nothing) Then
        sTexto = IIf(objForn.dQuantidade <> 0, Format(objForn.dQuantidade, "###,###,##0.0####"), "")
        If right(sTexto, 2) = ",0" Then sTexto = Format(objForn.dQuantidade, "###,###,##0")
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPedCotQtd = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPedCotQtd:

    MapaCot_ObterPedCotQtd = gErr
    
    Select Case gErr
        
        Case 111864
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153222)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCFrete(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCFrete

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not (objPC Is Nothing) Then
        sTexto = Format(objPC.dFrete, "STANDARD")
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCFrete = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCFrete:

    MapaCot_ObterPCFrete = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153203)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCFreteTipo(sTexto As String, ByVal sProduto As String, ByVal iSeq As Integer, ByVal iQtd As Integer) As Long

Dim lErro As Long
Dim objPC As ClassPCMapaCot

On Error GoTo Erro_MapaCot_ObterPCFreteTipo

    lErro = ObterDadosPCMapaCotacao(sProduto, iSeq, iQtd, objPC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not (objPC Is Nothing) Then
        sTexto = objPC.sFreteTipo
    Else
        sTexto = ""
    End If
    
    MapaCot_ObterPCFreteTipo = SUCESSO
    
    Exit Function

Erro_MapaCot_ObterPCFreteTipo:

    MapaCot_ObterPCFreteTipo = gErr
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153203)

    End Select

    Exit Function

End Function
