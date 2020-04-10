VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelNFProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Criada para atender a mapa de cotacao

Const TIPOPRODUTO_SERVICO = 15

Dim lComando As Long

Dim lNumIntDocAnt As Long

Dim gcolNFProd As Collection
Dim gcolNFServ As Collection

'------------------------------------------------------------------------
'CODIGO DEFAULT DAS CLASSES DE RELATORIO
'------------------------------------------------------------------------
Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
        
End Sub
'------------------------------------------------------------------------

Public Function NF_ObterCodProd(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Código do Produto

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterCodProd

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128126
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = objNFProd.sProduto
    End If
    
    NF_ObterCodProd = SUCESSO
    
    Exit Function

Erro_NF_ObterCodProd:

    NF_ObterCodProd = gErr
    
    Select Case gErr
        
        Case 128126
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153298)

    End Select

    Exit Function

End Function

Public Function NF_ObterDescProd(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Descrição do Produto

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterDescProd

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128127
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = objNFProd.sDescricao
    End If

    NF_ObterDescProd = SUCESSO
    
    Exit Function

Erro_NF_ObterDescProd:

    NF_ObterDescProd = gErr
    
    Select Case gErr
        
        Case 128127
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153299)

    End Select

    Exit Function

End Function

Public Function NF_ObterCF(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Classificação Fiscal

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterCF

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128128

    'Retorna o NF
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = objNFProd.sClassificacaoFiscal
    End If

    NF_ObterCF = SUCESSO
    
    Exit Function

Erro_NF_ObterCF:

    NF_ObterCF = gErr
    
    Select Case gErr
        
        Case 128128
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153300)

    End Select

    Exit Function

End Function

Public Function NF_ObterST(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Situação Tributária

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterST

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128129
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = objNFProd.sSituacaoTributaria
    End If

    
    NF_ObterST = SUCESSO
    
    Exit Function

Erro_NF_ObterST:

    NF_ObterST = gErr
    
    Select Case gErr
        
        Case 128129
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153301)

    End Select

    Exit Function

End Function

Public Function NF_ObterUM(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Unidade de Medida

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterUM

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128130
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = objNFProd.sUnidadeMed
    End If

    NF_ObterUM = SUCESSO
    
    Exit Function

Erro_NF_ObterUM:

    NF_ObterUM = gErr
    
    Select Case gErr
        
        Case 128130
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153302)

    End Select

    Exit Function

End Function

Public Function NF_ObterQuantidade(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Quantidade

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterQuantidade

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128131
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = objNFProd.dQuatidade
    End If

    NF_ObterQuantidade = SUCESSO
    
    Exit Function

Erro_NF_ObterQuantidade:

    NF_ObterQuantidade = gErr
    
    Select Case gErr
        
        Case 128131
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153303)

    End Select

    Exit Function

End Function

Public Function NF_ObterPrecoUnitario(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Preço Unitário

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterPrecoUnitario

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128132
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = Format(objNFProd.dPrecoUnitario, "Standard")
    End If

    NF_ObterPrecoUnitario = SUCESSO
    
    Exit Function

Erro_NF_ObterPrecoUnitario:

    NF_ObterPrecoUnitario = gErr
    
    Select Case gErr
        
        Case 128132
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153304)

    End Select

    Exit Function

End Function

Public Function NF_ObterPrecoTotal(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Preço Total

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterPrecoTotal

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128133
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = Format(objNFProd.dPrecoTotal, "Standard")
    End If

    NF_ObterPrecoTotal = SUCESSO
    
    Exit Function

Erro_NF_ObterPrecoTotal:

    NF_ObterPrecoTotal = gErr
    
    Select Case gErr
        
        Case 128133
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153305)

    End Select

    Exit Function

End Function

Public Function NF_ObterICMSAliquota(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Alíquota ICMS

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterICMSAliquota

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128134
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = CStr(Round(objNFProd.dICMSAliquota * 100, 1))
    End If

    NF_ObterICMSAliquota = SUCESSO
    
    Exit Function

Erro_NF_ObterICMSAliquota:

    NF_ObterICMSAliquota = gErr
    
    Select Case gErr
        
        Case 128134
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153306)

    End Select

    Exit Function

End Function

Public Function NF_ObterIPIAliquota(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a alíquota IPI

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterIPIAliquota

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128135
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = CStr(Round(objNFProd.dIPIAliquota * 100, 1))
    End If

    NF_ObterIPIAliquota = SUCESSO
    
    Exit Function

Erro_NF_ObterIPIAliquota:

    NF_ObterIPIAliquota = gErr
    
    Select Case gErr
        
        Case 128135
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153307)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorIPI(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Valor IPI

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterValorIPI

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128136
    
    If iLinha > gcolNFProd.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFProd(iLinha)
        sTexto = CStr(Round(objNFProd.dValorIPI, 2))
    End If

    NF_ObterValorIPI = SUCESSO
    
    Exit Function

Erro_NF_ObterValorIPI:

    NF_ObterValorIPI = gErr
    
    Select Case gErr
        
        Case 128136
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153308)

    End Select

    Exit Function

End Function

Public Function NF_ObterDescServ(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve a Descrição do Serviço

Dim lErro As Long
Dim objNFServ As ClassNFProd

On Error GoTo Erro_NF_ObterDescServ

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128137
    
    If iLinha > gcolNFServ.Count Then
        Set objNFServ = Nothing
        sTexto = ""
    Else
        Set objNFServ = gcolNFServ(iLinha)
        sTexto = objNFServ.sDescricao
    End If

    NF_ObterDescServ = SUCESSO
    
    Exit Function

Erro_NF_ObterDescServ:

    NF_ObterDescServ = gErr
    
    Select Case gErr
        
        Case 128137
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153309)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorServ(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Valor do Serviço

Dim lErro As Long
Dim objNFServ As ClassNFProd

On Error GoTo Erro_NF_ObterValorServ

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128138
    
    If iLinha > gcolNFServ.Count Then
        Set objNFServ = Nothing
        sTexto = ""
    Else
        Set objNFServ = gcolNFServ(iLinha)
        sTexto = CStr(Round(objNFServ.dPrecoTotal, 2))
    End If

    NF_ObterValorServ = SUCESSO
    
    Exit Function

Erro_NF_ObterValorServ:

    NF_ObterValorServ = gErr
    
    Select Case gErr
        
        Case 128138
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153310)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorISS(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Valor do ISS

Dim lErro As Long
Dim objNFServ As ClassNFProd

On Error GoTo Erro_NF_ObterValorISS

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128139
    
    If iLinha > gcolNFServ.Count Then
        Set objNFServ = Nothing
        sTexto = ""
    Else
        Set objNFServ = gcolNFServ(iLinha)
        sTexto = CStr(Round(objNFServ.dValorISS, 2))
    End If

    NF_ObterValorISS = SUCESSO
    
    Exit Function

Erro_NF_ObterValorISS:

    NF_ObterValorISS = gErr
    
    Select Case gErr
        
        Case 128139
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153311)

    End Select

    Exit Function

End Function

Public Function ObterDadosNFProdServ(ByVal lNumIntDoc As Long) As Long
'Verifica se está realizando a leitura de uma Nota Fiscal Diferente

Dim lErro As Long

On Error GoTo Erro_ObterDadosNFProdServ
 
    'se o NumIntDoc passado como parametro mudou
    If lNumIntDoc <> lNumIntDocAnt Then
               
        'le as informacoes relevantes dos IntensNF do NumIntDoc passado como parametro
        lErro = DadosNF_Le_Itens(lNumIntDoc, gcolNFProd, gcolNFServ, lComando)
        If lErro <> SUCESSO Then gError 128140
                        
        'seta como corrente o novo numIntDoc
        lNumIntDocAnt = lNumIntDoc
        
    End If
    
    ObterDadosNFProdServ = SUCESSO

    Exit Function

Erro_ObterDadosNFProdServ:

    ObterDadosNFProdServ = gErr

    Select Case gErr
        
        Case 128140
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153312)

    End Select

    Exit Function

End Function

Function DadosNF_Le_Itens(ByVal lNumIntDoc As Long, colNFProd As Collection, colNFServ As Collection, lComando As Long) As Long
'Le os itens da Nota Fiscal pedidos na colecao passada como parametro
'lNumIntDoc (INPUT) traz o NFiscal que é o filtro para a leitura dos Itens
'colNFProd (OUTPUT) é a colecao de retorno, que retornara os produtos lidos
'colNFServ (OUTPUT) é a colecao de retorno, que retornara os serviços lidos
'lComando (INPUT/OUTPUT) será utilizado para as operações de bd

Dim lErro As Long
Dim objNotaFiscal As ClassNFProd
Dim sProduto As String
Dim sDescricao As String
Dim dPrecoUnitario As Double
Dim dPrecoTotal As Double
Dim dValorDesconto As Double
Dim dQuantidade As Double
Dim dIPIValor As Double
Dim dValorISS As Double
Dim dIPIAliquota As Double
Dim dICMSAliquota As Double
Dim sClassificacaoFiscal As String
Dim iTipoTribCST As Integer
Dim iTipoProduto As Integer
Dim iOrigemMercadoria As Integer
Dim sUnidadeMed As String

On Error GoTo Erro_DadosNF_Le_Itens

    Set colNFProd = New Collection
    Set colNFServ = New Collection
        
    'se a conexao ainda nao foi aberta... abre...
    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 128141
    End If
    
    'inicializa as strings
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sClassificacaoFiscal = String(STRING_PRODUTO_IPI_CODIGO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    'obtém o conjunto de itens associados ao numintdoc
    lErro = Comando_Executar(lComando, "SELECT Produtos.Codigo, ItensNFiscal.DescricaoItem, ItensNFiscal.Quantidade, ItensNFiscal.PrecoUnitario, ItensNFiscal.ValorDesconto, TributacaoItemNF.IPIAliquota, TributacaoItemNF.ICMSAliquota, TributacaoItemNF.IPIValor, Produtos.Tipo, Produtos.IPICodigo, TiposTribICMS.TipoTribCST, ItensNFiscal.UnidadeMed, Produtos.OrigemMercadoria, TributacaoNF.ISSValor FROM ItensNFiscal, Produtos, TributacaoItemNF, TributacaoNF, TiposTribICMS WHERE TiposTribICMS.Tipo = TributacaoItemNF.ICMSTipo AND ItensNFiscal.NumIntNF = ? AND ItensNFiscal.Produto = Produtos.Codigo AND ItensNFiscal.NumIntDoc = TributacaoItemNF.NumIntDoc AND ItensNFiscal.NumIntNF = TributacaoNF.NumIntDoc ORDER BY ItensNFiscal.NumIntNF, ItensNFiscal.Item", _
        sProduto, sDescricao, dQuantidade, dPrecoUnitario, dValorDesconto, dIPIAliquota, dICMSAliquota, dIPIValor, iTipoProduto, sClassificacaoFiscal, iTipoTribCST, sUnidadeMed, iOrigemMercadoria, dValorISS, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 128142
    
    'le o primeiro pedido
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128143
    
    'enquanto for achando itens
    Do While lErro = AD_SQL_SUCESSO
    
        'instancia novo obj
        Set objNotaFiscal = New ClassNFProd
        
        'preenche o obj com os dados lidos
        With objNotaFiscal
        
            .sProduto = sProduto
            .sDescricao = sDescricao
            .dQuatidade = dQuantidade
            .dPrecoTotal = Arredonda_Moeda(dPrecoUnitario * dQuantidade - dValorDesconto)
            If dValorDesconto = 0 Or dQuantidade = 0 Then
                .dPrecoUnitario = dPrecoUnitario
            Else
                .dPrecoUnitario = Arredonda_Moeda(.dPrecoTotal / dQuantidade)
            End If
            .dIPIAliquota = dIPIAliquota
            .dICMSAliquota = dICMSAliquota
            .dValorIPI = dIPIValor
            .dValorISS = dValorISS
            .iTipoProduto = iTipoProduto
            .sClassificacaoFiscal = sClassificacaoFiscal
            .sSituacaoTributaria = CStr(iOrigemMercadoria) & Format(iTipoTribCST, "00")
            .iOrigemMercadoria = iOrigemMercadoria
            .sUnidadeMed = sUnidadeMed
            
        End With
            
        If iTipoProduto = TIPOPRODUTO_SERVICO Then
        
            'adiciona o item na colecao Serviço
            colNFServ.Add objNotaFiscal
            
        Else
            
            'adiciona o item na colecao Produto
            colNFProd.Add objNotaFiscal
            
        End If
        
        'busca o proximo pedido
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128144
       
    Loop
    
    DadosNF_Le_Itens = SUCESSO
    
    Exit Function
 
Erro_DadosNF_Le_Itens:

    DadosNF_Le_Itens = gErr

    Select Case gErr
    
        Case 128141
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 128142 To 128144
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153313)
    
    End Select
    
    Exit Function

End Function

Public Function NF_ObterValorUnitServ(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Valor Unitario do Serviço

Dim lErro As Long
Dim objNFServ As ClassNFProd

On Error GoTo Erro_NF_ObterValorUnitServ

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128138
    
    If iLinha > gcolNFServ.Count Then
        Set objNFServ = Nothing
        sTexto = ""
    Else
        Set objNFServ = gcolNFServ(iLinha)
        sTexto = CStr(Round(objNFServ.dPrecoUnitario, 2))
    End If

    NF_ObterValorUnitServ = SUCESSO
    
    Exit Function

Erro_NF_ObterValorUnitServ:

    NF_ObterValorUnitServ = gErr
    
    Select Case gErr
        
        Case 128138
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153314)

    End Select

    Exit Function

End Function

Public Function NF_ObterQtdeServ(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Valor Unitario do Serviço

Dim lErro As Long
Dim objNFServ As ClassNFProd

On Error GoTo Erro_NF_ObterQtdeServ

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128138
    
    If iLinha > gcolNFServ.Count Then
        Set objNFServ = Nothing
        sTexto = ""
    Else
        Set objNFServ = gcolNFServ(iLinha)
        sTexto = CStr(Round(objNFServ.dQuatidade, 0))
    End If

    NF_ObterQtdeServ = SUCESSO
    
    Exit Function

Erro_NF_ObterQtdeServ:

    NF_ObterQtdeServ = gErr
    
    Select Case gErr
        
        Case 128138
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153315)

    End Select

    Exit Function

End Function

Public Function NF_ObterUMServ(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Valor Unitario do Serviço

Dim lErro As Long
Dim objNFServ As ClassNFProd

On Error GoTo Erro_NF_ObterUMServ

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 128138
    
    If iLinha > gcolNFServ.Count Then
        Set objNFServ = Nothing
        sTexto = ""
    Else
        Set objNFServ = gcolNFServ(iLinha)
        sTexto = objNFServ.sUnidadeMed
    End If

    NF_ObterUMServ = SUCESSO
    
    Exit Function

Erro_NF_ObterUMServ:

    NF_ObterUMServ = gErr
    
    Select Case gErr
        
        Case 128138
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153316)

    End Select

    Exit Function

End Function

'###########################################
'Inserido por Wagner
Public Function NF_ObterCodServ(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve o Código do Produto

Dim lErro As Long
Dim objNFProd As ClassNFProd

On Error GoTo Erro_NF_ObterCodServ

    lErro = ObterDadosNFProdServ(lNumIntDoc)
    If lErro <> SUCESSO Then gError 138384
    
    If iLinha > gcolNFServ.Count Then
        Set objNFProd = Nothing
        sTexto = ""
    Else
        Set objNFProd = gcolNFServ(iLinha)
        sTexto = objNFProd.sProduto
    End If
    
    NF_ObterCodServ = SUCESSO
    
    Exit Function

Erro_NF_ObterCodServ:

    NF_ObterCodServ = gErr
    
    Select Case gErr
        
        Case 138384
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153317)

    End Select

    Exit Function

End Function
'###########################################

