VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelProdFatFilCliPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sUltProduto As String
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
Private dtUltData As Date

Private iUltFilialEmpresa As Integer
Private iUltFilialCli As Integer
Private lUltCliente As Long

Private dUltQuantidadeFaturada As Double
Private dUltQuantidadeDevolvida As Double
Private dUltValorDevolvido As Double
Private dUltValorFaturado As Double

Private lComando As Long

Public Function ObterQtdeDevolvidaFatProdFilCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaFatProdFilCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Then
        
        lErro = ObterRelMesProdFatFilCli(iFilialEmpresa, iFilialCli, lCliente, sProduto, dtDataInicial, dtDataFinal)
        If lErro <> SUCESSO Then gError 90191
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvida
    
    ObterQtdeDevolvidaFatProdFilCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevolvidaFatProdFilCliPer:

    ObterQtdeDevolvidaFatProdFilCliPer = gErr
    
    Select Case gErr

        Case 90191
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153320)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaFatProdFilCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaFatProdFilCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Or iUltFilialCli <> iFilialCli Or lUltCliente <> lCliente Then
        
        lErro = ObterRelMesProdFatFilCli(iFilialEmpresa, iFilialCli, lCliente, sProduto, dtDataInicial, dtDataFinal)
        If lErro <> SUCESSO Then gError 90192
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturada
    
    ObterQtdeFaturadaFatProdFilCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFaturadaFatProdFilCliPer:

    ObterQtdeFaturadaFatProdFilCliPer = gErr
    
    Select Case gErr

        Case 90192
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153321)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoFatProdFilCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dValorDevolvido As Double) As Long
'Para obter valor devolvido por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorDevolvidoFatProdFilCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Then
        
        lErro = ObterRelMesProdFatFilCli(iFilialEmpresa, iFilialCli, lCliente, sProduto, dtDataInicial, dtDataFinal)
        If lErro <> SUCESSO Then gError 90685
    
    End If
    
    dValorDevolvido = dUltValorDevolvido
    
    ObterValorDevolvidoFatProdFilCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorDevolvidoFatProdFilCliPer:

    ObterValorDevolvidoFatProdFilCliPer = gErr
    
    Select Case gErr

        Case 90685
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153322)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoFatProdFilCliPer(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dValorFaturado As Double) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoFatProdFilCliPer
    
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iFilialEmpresa <> iUltFilialEmpresa Or iUltFilialCli <> iFilialCli Or lUltCliente <> lCliente Then
        
        lErro = ObterRelMesProdFatFilCli(iFilialEmpresa, iFilialCli, lCliente, sProduto, dtDataInicial, dtDataFinal)
        If lErro <> SUCESSO Then gError 90686
    
    End If
    
    dValorFaturado = dUltValorFaturado
    
    ObterValorFaturadoFatProdFilCliPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorFaturadoFatProdFilCliPer:

    ObterValorFaturadoFatProdFilCliPer = gErr
    
    Select Case gErr

        Case 90686
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153323)

    End Select

    Exit Function

End Function

Private Function ObterRelMesProdFatFilCli(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, sProduto As String, dtDataIni As Date, dtDataFim As Date) As Long
'Para obter quantidade faturada e devolvida, valor faturado e devolvido e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double

On Error GoTo Erro_ObterRelMesProdFatFilCli
 
    lErro = Obter_QtdeDevFat_ValorDevFat_CliProdFilCliPer(iFilialEmpresa, iFilialCli, lCliente, sProduto, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90157
    
    iUltFilialEmpresa = iFilialEmpresa
    iUltFilialCli = iFilialCli
    lUltCliente = lCliente
    sUltProduto = sProduto
    dtUltDataInicial = dtDataIni
    dtUltDataFinal = dtDataFim
    
    dUltQuantidadeFaturada = dQuantidadeFaturada
    dUltQuantidadeDevolvida = dQuantidadeDevolvida
    dUltValorDevolvido = dValorDevolvido
    dUltValorFaturado = dValorFaturado
    
    ObterRelMesProdFatFilCli = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesProdFatFilCli:

    ObterRelMesProdFatFilCli = gErr
    
    Select Case gErr

        Case 90157
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153324)

    End Select

    Exit Function
    
End Function
Function Preenche_SldMesFatFilCli(ByVal iFilialEmpresa As Integer, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long
'Povoa a tabela SldMesFatFilCli para gerar relatório

Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim adValor(1 To 48) As Double
Dim lErro As Long
Dim lErroBusca As Long
Dim lTransacao As Long
Dim dFator As Double
Dim dQtd As Double

'Variáveis auxiliares
Dim iAnoAux As Integer
Dim sCodProdutoAux As String
Dim lClienteAux As Long
Dim iFilialCliAux As Integer
Dim iFilialEmpresaAux As Integer

'Variáveis utilizadas no Select
Dim dQuantidade As Double
Dim dValorTotal As Double
Dim sCodProduto As String
Dim lCliente As Long
Dim iFilialCli As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim iFaturamento As Integer
Dim sUnidadeMed As String
Dim iClasseUM As Integer
Dim iCodTipoDocInfo As Integer
Dim sSiglaUMVenda As String

'Variáveis Auxiliares para comandos SQL
Dim sSelect As String
Dim sWhere As String

On Error GoTo Erro_Preenche_SldMesFatFilCli

    'Abertura da Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 110215

    'Abre comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 110210
    Next
    
    'Aloca Espaço na String para trazer dados do banco
    sCodProduto = String(STRING_PRODUTO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sSiglaUMVenda = String(STRING_UM_SIGLA, 0)
    
    sSelect = "SELECT N.FilialEmpresa, P.Codigo,N.Cliente, N.FilialCli, SUM(I.Quantidade), SUM(I.Quantidade * I.PrecoUnitario - I.ValorDesconto), T.Faturamento, YEAR( N.dataemissao),MONTH( N.dataemissao), I.UnidadeMed, P.ClasseUM, T.Codigo, P.SiglaUMVenda " & _
              "FROM TiposDocInfo AS T,ItensNFiscal AS I,NFiscal AS N,Produtos AS P "
    sWhere = " N.DataEmissao >= ? AND N.DataEmissao <= ? AND N.NumIntDoc = I.NumIntNF AND T.Codigo = N.TipoNFiscal AND (T.Faturamento IN (1, 2))AND I.Produto = P.Codigo AND N.Status <> 7 AND (T.Complementar = 0 OR T.Codigo IN (?,?)) " & _
              "GROUP BY N.FilialEmpresa, N.Cliente, N.FilialCli, P.Codigo, YEAR( N.dataemissao),MONTH( N.dataemissao), T.Faturamento, I.UnidadeMed, P.ClasseUM, T.Codigo, P.SiglaUMVenda " & _
              "ORDER BY N.FilialEmpresa, N.Cliente, N.FilialCli, P.Codigo, YEAR( N.dataemissao),MONTH( N.dataemissao), T.Faturamento, I.UnidadeMed, P.ClasseUM, T.Codigo, P.SiglaUMVenda"
    
    If iFilialEmpresa = EMPRESA_TODA Then
        sWhere = "WHERE " & sWhere
        lErro = Comando_Executar(alComando(1), sSelect & sWhere, iFilialEmpresa, sCodProduto, lCliente, iFilialCli, dQuantidade, dValorTotal, iFaturamento, iAno, iMes, sUnidadeMed, iClasseUM, iCodTipoDocInfo, sSiglaUMVenda, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    Else
        'Filtra por FilialEmpresa
        sWhere = "WHERE N.FilialEmpresa = ? AND " & sWhere
        lErro = Comando_Executar(alComando(1), sSelect & sWhere, iFilialEmpresa, sCodProduto, lCliente, iFilialCli, dQuantidade, dValorTotal, iFaturamento, iAno, iMes, sUnidadeMed, iClasseUM, iCodTipoDocInfo, sSiglaUMVenda, iFilialEmpresa, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 110214
    
    'Busca Primeira Tupla
    lErroBusca = Comando_BuscarPrimeiro(alComando(1))
    If lErroBusca <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110211
    
    'Aux Contém a Chave Primária anterior a mudança
    iFilialEmpresaAux = iFilialEmpresa
    iAnoAux = iAno
    sCodProdutoAux = sCodProduto
    lClienteAux = lCliente
    iFilialCliAux = iFilialCli
    
    lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, IIf(iClasseUM = 2, "un", sSiglaUMVenda), dFator)
    If lErro <> SUCESSO Then gError 110218
                
    If iCodTipoDocInfo <> DOCINFO_NFISCP And iCodTipoDocInfo <> DOCINFO_NFISFCP Then
        
        If iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
            adValor(iMes) = dQuantidade * dFator
        Else
            adValor(iMes + 12) = dQuantidade * dFator
        End If
                        
    End If
       
    'Preenche advalor
    If iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
        adValor(iMes + 24) = dValorTotal
    Else
        adValor(iMes + 36) = dValorTotal
    End If

    Do While lErroBusca <> AD_SQL_SEM_DADOS
       
        'Busca Próxima tupla
        lErroBusca = Comando_BuscarProximo(alComando(1))
        If lErroBusca <> AD_SQL_SUCESSO And lErroBusca <> AD_SQL_SEM_DADOS Then gError 110213
        
        'Verifica se mudou a Chave Primária
        If lErroBusca = AD_SQL_SEM_DADOS Or iFilialEmpresa <> iFilialEmpresaAux Or iAno <> iAnoAux Or sCodProduto <> sCodProdutoAux Or lCliente <> lClienteAux Or iFilialCli <> iFilialCliAux Then
                
            lErro = Comando_Executar(alComando(2), "INSERT INTO SldMesFatFilCli (FilialEmpresa, Ano, Produto, Cliente, Filial, QtdFaturada1, QtdFaturada2, QtdFaturada3, " & _
                                    "QtdFaturada4, QtdFaturada5, QtdFaturada6, QtdFaturada7, QtdFaturada8, QtdFaturada9, " & _
                                    "QtdFaturada10, QtdFaturada11, QtdFaturada12, QtdDevolvida1, QtdDevolvida2, QtdDevolvida3, " & _
                                    "QtdDevolvida4, QtdDevolvida5, QtdDevolvida6, QtdDevolvida7, QtdDevolvida8, QtdDevolvida9, " & _
                                    "QtdDevolvida10, QtdDevolvida11, QtdDevolvida12, ValorFaturado1, ValorFaturado2, ValorFaturado3, " & _
                                    "ValorFaturado4, ValorFaturado5, ValorFaturado6, ValorFaturado7, ValorFaturado8, ValorFaturado9, " & _
                                    "ValorFaturado10, ValorFaturado11, ValorFaturado12, ValorDevolvido1, ValorDevolvido2, " & _
                                    "ValorDevolvido3, ValorDevolvido4, ValorDevolvido5, ValorDevolvido6, ValorDevolvido7, " & _
                                    "ValorDevolvido8, ValorDevolvido9, ValorDevolvido10, ValorDevolvido11, ValorDevolvido12) " & _
                                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                    iFilialEmpresaAux, iAnoAux, sCodProdutoAux, lClienteAux, iFilialCliAux, adValor(1), adValor(2), adValor(3), adValor(4), adValor(5), adValor(6), _
                                    adValor(7), adValor(8), adValor(9), adValor(10), adValor(11), adValor(12), adValor(13), adValor(14), adValor(15), adValor(16), adValor(17), adValor(18), _
                                    adValor(19), adValor(20), adValor(21), adValor(22), adValor(23), adValor(24), adValor(25), adValor(26), adValor(27), adValor(28), adValor(29), adValor(30), _
                                    adValor(31), adValor(32), adValor(33), adValor(34), adValor(35), adValor(36), adValor(37), adValor(38), adValor(39), adValor(40), adValor(41), adValor(42), _
                                    adValor(43), adValor(44), adValor(45), adValor(46), adValor(47), adValor(48))
            If lErro <> AD_SQL_SUCESSO Then gError 110212
            
            'Limpa adValor
            For iIndice = LBound(adValor) To UBound(adValor)
                adValor(iIndice) = 0
            Next
            
            iFilialEmpresaAux = iFilialEmpresa
            iAnoAux = iAno
            sCodProdutoAux = sCodProduto
            lClienteAux = lCliente
            iFilialCliAux = iFilialCli
            
        End If
        
        lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, IIf(iClasseUM = 2, "un", sSiglaUMVenda), dFator)
        If lErro <> SUCESSO Then gError 110217
                
        If iCodTipoDocInfo <> DOCINFO_NFISCP And iCodTipoDocInfo <> DOCINFO_NFISFCP Then
            
            If iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
                adValor(iMes) = adValor(iMes) + dQuantidade * dFator
            Else
                adValor(iMes + 12) = adValor(iMes + 12) + dQuantidade * dFator
            End If
                            
        End If
           
        'Preenche advalor
        If iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
            adValor(iMes + 24) = adValor(iMes + 24) + dValorTotal
        Else
            adValor(iMes + 36) = adValor(iMes + 36) + dValorTotal
        End If
    
    Loop
    
    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 110216
    
    Preenche_SldMesFatFilCli = SUCESSO
    
    Exit Function
    
Erro_Preenche_SldMesFatFilCli:

    Preenche_SldMesFatFilCli = gErr
    
    Select Case gErr
    
        Case 110210
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 110212
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFATFILCLI", gErr, iAnoAux, iFilialEmpresaAux, sCodProduto, lCliente, iFilialCli)
            
        Case 110211, 110213, 110214
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
     
        Case 110215
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 110216
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
            
        Case 110217, 110218
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153325)

    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Cancela Alterações
    Call Transacao_Rollback

    Exit Function

End Function


Function Obter_QtdeDevFat_ValorDevFat_CliProdFilCliPer(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, sProduto As String, dtDataIni As Date, dtDataFim As Date, dQuantidadeDevolvida As Double, dQuantidadeFaturada As Double, dValorDevolvido As Double, dValorFaturado As Double, lComando As Long) As Long
'Obtem a Quantidade devolvida e faturada, Valor devolvido e faturado no Dia, Mes e Ano, de uma FilialEmpresa(passada po parametro),
'de um Cliente(passado por parametro),de um FiliaCli(passado por parametro), de um sProduto(passado por parametro), de uma data(passada por parametro)

Dim lErro As Long
Dim dFator As Double
Dim dQuantidade As Double
Dim dValorTotal As Double
Dim iFaturamento As Integer
Dim sUnidadeMed As String, sUMBase As String
Dim iClasseUM As Integer
Dim iCodTipoDocInfo As Integer

On Error GoTo Erro_Obter_QtdeDevFat_ValorDevFat_CliProdFilCliPer

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90300
    End If

    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sUMBase = String(STRING_UM_SIGLA, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        'Acumula a Quantidade e o Valor, devolvido e Faturado do mes da data passada por parametro, nao comparando a Filial Empresa.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
        "FROM TiposDocInfo, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM " & _
        "WHERE NFiscal.FilialCli = ? AND NFiscal.Cliente = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?  AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7 AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
        "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
        dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iFilialCli, lCliente, sProduto, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)

    Else
        'Acumula a Quantidade e o Valor, devolvido e Faturado do mes da data passada por parametro.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
        "FROM TiposDocInfo, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM " & _
        "WHERE NFiscal.FilialEmpresa = ?  AND NFiscal.FilialCli = ? AND NFiscal.Cliente = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?  AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
        "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
        dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iFilialEmpresa, iFilialCli, lCliente, sProduto, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)

    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90301

    Do While lErro = AD_SQL_SUCESSO

        'Faturada
        If iFaturamento = TIPODOCINFO_FATURAMENTO_SIM Then
            
            dValorFaturado = dValorFaturado + dValorTotal

'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90302
                
                If iCodTipoDocInfo <> DOCINFO_NFISCP And iCodTipoDocInfo <> DOCINFO_NFISFCP Then
                    dQuantidadeFaturada = dQuantidadeFaturada + (dQuantidade * dFator)
                End If
'            End If

        Else
            'Devolvida
            dValorDevolvido = dValorDevolvido + dValorTotal

'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90303

                dQuantidadeDevolvida = dQuantidadeDevolvida + (dQuantidade * dFator)
'            End If

        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90304

    Loop

    Obter_QtdeDevFat_ValorDevFat_CliProdFilCliPer = SUCESSO

    Exit Function

Erro_Obter_QtdeDevFat_ValorDevFat_CliProdFilCliPer:

    Obter_QtdeDevFat_ValorDevFat_CliProdFilCliPer = gErr

    Select Case gErr

        Case 90300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90301, 90304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 90302, 90303

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153326)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

