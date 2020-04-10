VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFatFilCliPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iUltFilialEmpresa As Integer
Private iUltFilialCli As Integer
Private lUltCliente As Long
Private iUltMes As Integer

Private dUltQuantidadeFaturada As Double
Private dUltQuantidadeDevolvida As Double
Private dUltValorDevolvido As Double
Private dUltValorFaturado As Double

Private lComando As Long

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

Public Function ObterQtdeDevFilCliMes(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, iMes As Integer, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida no mês

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevFilCliMes
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or iMes <> iUltMes Then
        
        lErro = ObterRelFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes)
        If lErro <> SUCESSO Then gError 90280
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvida
    
    ObterQtdeDevFilCliMes = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevFilCliMes:

    ObterQtdeDevFilCliMes = gErr
    
    Select Case gErr

        Case 90280
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153101)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obterqtdedevfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterQtdeDevFilCliMes2(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dtDataFim As Date, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida no mês

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevFilCliMes2
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or Month(dtDataFim) <> iUltMes Then
        
        lErro = ObterRelFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim)
        If lErro <> SUCESSO Then gError 111706
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvida
    
    ObterQtdeDevFilCliMes2 = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevFilCliMes2:

    ObterQtdeDevFilCliMes2 = gErr
    
    Select Case gErr

        Case 111706
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153102)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFatFilCliMes(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, iMes As Integer, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada no mês

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFatFilCliMes
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or iMes <> iUltMes Then
        
        lErro = ObterRelFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes)
        If lErro <> SUCESSO Then gError 90281
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturada
    
    ObterQtdeFatFilCliMes = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFatFilCliMes:

    ObterQtdeFatFilCliMes = gErr
    
    Select Case gErr

        Case 90281
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153103)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obterqtdefatfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterQtdeFatFilCliMes2(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dtDataFim As Date, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada no mês

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFatFilCliMes2
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or Month(dtDataFim) <> iUltMes Then
        
        lErro = ObterRelFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim)
        If lErro <> SUCESSO Then gError 111707
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturada
    
    ObterQtdeFatFilCliMes2 = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFatFilCliMes2:

    ObterQtdeFatFilCliMes2 = gErr
    
    Select Case gErr

        Case 111707
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153104)

    End Select

    Exit Function

End Function

Public Function ObterValorDevFilCliMes(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, iMes As Integer, dValorDevolvido As Double) As Long
'Para obter o valor devolvido no mês

Dim lErro As Long

On Error GoTo Erro_ObterValorDevFilCliMes
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or iMes <> iUltMes Then
        
        lErro = ObterRelFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes)
        If lErro <> SUCESSO Then gError 90282
    
    End If
    
    dValorDevolvido = dUltValorDevolvido
    
    ObterValorDevFilCliMes = SUCESSO
    
    Exit Function
    
Erro_ObterValorDevFilCliMes:

    ObterValorDevFilCliMes = gErr
    
    Select Case gErr

        Case 90282
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153105)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obtervalordevfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterValorDevFilCliMes2(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dtDataFim As Date, dValorDevolvido As Double) As Long
'Para obter o valor devolvido no mês

Dim lErro As Long

On Error GoTo Erro_ObterValorDevFilCliMes2
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or Month(dtDataFim) <> iUltMes Then
        
        lErro = ObterRelFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim)
        If lErro <> SUCESSO Then gError 111708
    
    End If
    
    dValorDevolvido = dUltValorDevolvido
    
    ObterValorDevFilCliMes2 = SUCESSO
    
    Exit Function
    
Erro_ObterValorDevFilCliMes2:

    ObterValorDevFilCliMes2 = gErr
    
    Select Case gErr

        Case 111708
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153106)

    End Select

    Exit Function

End Function

Public Function ObterValorFatFilCliMes(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, iMes As Integer, dValorFaturado As Double) As Long
'Para obter o valor faturado no mês

Dim lErro As Long

On Error GoTo Erro_ObterValorFatFilCliMes
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or iMes <> iUltMes Then
        
        lErro = ObterRelFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes)
        If lErro <> SUCESSO Then gError 90283
    
    End If
    
    dValorFaturado = dUltValorFaturado
    
    ObterValorFatFilCliMes = SUCESSO
    
    Exit Function
    
Erro_ObterValorFatFilCliMes:

    ObterValorFatFilCliMes = gErr
    
    Select Case gErr

        Case 90283
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153107)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obtervalorfatfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterValorFatFilCliMes2(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dtDataFim As Date, dValorFaturado As Double) As Long
'Para obter o valor faturado no mês

Dim lErro As Long

On Error GoTo Erro_ObterValorFatFilCliMes2
    
    If iFilialEmpresa <> iUltFilialEmpresa Or iFilialCli <> iUltFilialCli Or lCliente <> lUltCliente Or Month(dtDataFim) <> iUltMes Then
        
        lErro = ObterRelFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim)
        If lErro <> SUCESSO Then gError 111709
    
    End If
    
    dValorFaturado = dUltValorFaturado
    
    ObterValorFatFilCliMes2 = SUCESSO
    
    Exit Function
    
Erro_ObterValorFatFilCliMes2:

    ObterValorFatFilCliMes2 = gErr
    
    Select Case gErr

        Case 111709
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153108)

    End Select

    Exit Function

End Function

Private Function ObterRelFilCliMes(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, iMes As Integer) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dtDataIni As Date
Dim dtDataFim As Date

Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double


On Error GoTo Erro_ObterRelFilCliMes

    dtDataIni = CDate("01/" & iMes & "/" & IIf(iMes = 12, Year(gdtDataHoje) - 1, Year(gdtDataHoje)))
    dtDataFim = DateAdd("m", 1, dtDataIni) - 1

    lErro = Obter_QtdeDevFat_ValorDevFat_FilCliMes(iFilialEmpresa, iFilialCli, lCliente, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90284
    
    iUltFilialEmpresa = iFilialEmpresa
    iUltFilialCli = iFilialCli
    lUltCliente = lCliente
    iUltMes = iMes
    
    dUltQuantidadeDevolvida = dQuantidadeDevolvida
    dUltQuantidadeFaturada = dQuantidadeFaturada
    dUltValorDevolvido = dValorDevolvido
    dUltValorFaturado = dValorFaturado

    ObterRelFilCliMes = SUCESSO
    
    Exit Function
    
Erro_ObterRelFilCliMes:

    ObterRelFilCliMes = gErr
    
    Select Case gErr

        Case 90284
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153109)

    End Select

    Exit Function
    
End Function

'copiada por tulio271202
'a versao anterior (obterrelfilclimes) soh recebia o mes e
'calculava a data de maneira meio tricky...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Private Function ObterRelFilCliMes2(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dtDataFim As Date) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dtDataIni As Date

Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double


On Error GoTo Erro_ObterRelFilCliMes2

    dtDataIni = CDate("01/" & Month(dtDataFim) & "/" & Year(dtDataFim))
    
    lErro = Obter_QtdeDevFat_ValorDevFat_FilCliMes(iFilialEmpresa, iFilialCli, lCliente, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 111705
    
    iUltFilialEmpresa = iFilialEmpresa
    iUltFilialCli = iFilialCli
    lUltCliente = lCliente
    iUltMes = Month(dtDataFim)
    
    dUltQuantidadeDevolvida = dQuantidadeDevolvida
    dUltQuantidadeFaturada = dQuantidadeFaturada
    dUltValorDevolvido = dValorDevolvido
    dUltValorFaturado = dValorFaturado

    ObterRelFilCliMes2 = SUCESSO
    
    Exit Function
    
Erro_ObterRelFilCliMes2:

    ObterRelFilCliMes2 = gErr
    
    Select Case gErr

        Case 111705
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153110)

    End Select

    Exit Function
    
End Function

Function Obter_QtdeDevFat_ValorDevFat_FilCliMes(iFilialEmpresa As Integer, iFilialCli As Integer, lCliente As Long, dtDataIni As Date, dtDataFim As Date, dQuantidadeDevolvida As Double, dQuantidadeFaturada As Double, dValorDevolvido As Double, dValorFaturado As Double, lComando As Long) As Long
'Obtem a Quantidade devolvida e faturada, Valor devolvido e faturado de uma FilialEmpresa(passada po parametro),
'em uma iFlilialCli(passada por parametro), em uma lCliente(passada por parametro),em uma regiao(passada por parametro), em um mes(passado por parametro)

Dim lErro As Long
Dim dFator As Double

Dim dQuantidade As Double
Dim dValorTotal As Double
Dim iFaturamento As Integer
Dim sUnidadeMed As String, sUMBase As String
Dim iClasseUM As Integer
Dim iCodTipoDocInfo As Integer

On Error GoTo Erro_Obter_QtdeDevFat_ValorDevFat_FilCliMes

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90285
    End If
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sUMBase = String(STRING_UM_SIGLA, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
       
        'Acumula a Quantidade e o Valor, devolvido e Faturado do mes da data passada por parametro, nao comparando a Filial Empresa.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
        "FROM TiposDocInfo, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM WHERE NFiscal.FilialCli = ? AND NFiscal.Cliente = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?  AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
        "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
        dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iFilialCli, lCliente, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    Else
        'Acumula a Quantidade e o Valor, devolvido e Faturado do mes da data passada por parametro.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
        "FROM TiposDocInfo, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM WHERE NFiscal.FilialEmpresa = ? AND NFiscal.FilialCli = ? AND NFiscal.Cliente = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?  AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
        "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
        dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iFilialEmpresa, iFilialCli, lCliente, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90286

    Do While lErro = AD_SQL_SUCESSO
        
        'Faturada
        If iFaturamento = 1 Then
            dValorFaturado = dValorFaturado + dValorTotal
            
'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90287

                If iCodTipoDocInfo <> DOCINFO_NFISCP And iCodTipoDocInfo <> DOCINFO_NFISFCP Then
                    dQuantidadeFaturada = dQuantidadeFaturada + (dQuantidade * dFator)
                End If
'            End If

        Else
            'Devolvida
            dValorDevolvido = dValorDevolvido + dValorTotal

'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90288
                
                dQuantidadeDevolvida = dQuantidadeDevolvida + (dQuantidade * dFator)
'            End If
        
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90289
           
    Loop
          
    Obter_QtdeDevFat_ValorDevFat_FilCliMes = SUCESSO
    
    Exit Function
    
Erro_Obter_QtdeDevFat_ValorDevFat_FilCliMes:

    Obter_QtdeDevFat_ValorDevFat_FilCliMes = gErr
    
    Select Case gErr

        Case 90285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 90286, 90289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 90287, 90288
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153111)

    End Select

    Call Comando_Fechar(lComando)
        
    Exit Function
    
End Function


