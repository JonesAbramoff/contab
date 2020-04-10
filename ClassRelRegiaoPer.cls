VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRegiaoPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iUltFilialEmpresaPer As Integer
Private iUltCodRegiaoPer As Integer
Private iUltFilialEmpresaData As Integer
Private iUltCodRegiaoData As Integer
Private iUltMes As Integer
Private dtUltData As Date

Private dUltQuantidadeFaturadaPer As Double
Private dUltQuantidadeDevolvidaPer As Double
Private dUltValorDevolvidoPer As Double
Private dUltValorFaturadoPer As Double

Private dUltQuantidadeFaturadaData As Double
Private dUltQuantidadeDevolvidaData As Double
Private dUltValorDevolvidoData As Double
Private dUltValorFaturadoData As Double

Private lComando As Long

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

Public Function ObterQtdeDevolvidaRegiaoPer(iFilialEmpresa As Integer, iCodRegiao As Integer, iMes As Integer, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida por uma Região no mês

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaRegiaoPer
    
    If iFilialEmpresa <> iUltFilialEmpresaPer Or iCodRegiao <> iUltCodRegiaoPer Or iMes <> iUltMes Then
        
        lErro = ObterRelMesFatRegiaoPer(iFilialEmpresa, iCodRegiao, iMes)
        If lErro <> SUCESSO Then gError 90250
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvidaPer
    
    ObterQtdeDevolvidaRegiaoPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevolvidaRegiaoPer:

    ObterQtdeDevolvidaRegiaoPer = gErr
    
    Select Case gErr

        Case 90250
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153327)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaRegiaoPer(iFilialEmpresa As Integer, iCodRegiao As Integer, iMes As Integer, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada por uma Região no mês.

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaRegiaoPer
    
    If iFilialEmpresa <> iUltFilialEmpresaPer Or iCodRegiao <> iUltCodRegiaoPer Or iMes <> iUltMes Then
        
        lErro = ObterRelMesFatRegiaoPer(iFilialEmpresa, iCodRegiao, iMes)
        If lErro <> SUCESSO Then gError 90251
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturadaPer
    
    ObterQtdeFaturadaRegiaoPer = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFaturadaRegiaoPer:

    ObterQtdeFaturadaRegiaoPer = gErr
    
    Select Case gErr

        Case 90251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153328)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoRegiaoPer(iFilialEmpresa As Integer, iCodRegiao As Integer, iMes As Integer, dValorDevolvido As Double) As Long
'Para obter o valor devolvido por uma Região no mês

Dim lErro As Long

On Error GoTo Erro_ObterValorDevolvidoRegiaoPer
    
    If iFilialEmpresa <> iUltFilialEmpresaPer Or iCodRegiao <> iUltCodRegiaoPer Or iMes <> iUltMes Then
        
        lErro = ObterRelMesFatRegiaoPer(iFilialEmpresa, iCodRegiao, iMes)
        If lErro <> SUCESSO Then gError 90252
    
    End If
    
    dValorDevolvido = dUltValorDevolvidoPer
    
    ObterValorDevolvidoRegiaoPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorDevolvidoRegiaoPer:

    ObterValorDevolvidoRegiaoPer = gErr
    
    Select Case gErr

        Case 90252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153329)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoRegiaoPer(iFilialEmpresa As Integer, iCodRegiao As Integer, iMes As Integer, dValorFaturado As Double) As Long
'Para obter o valor faturado por uma Região no mês.

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoRegiaoPer
    
    If iFilialEmpresa <> iUltFilialEmpresaPer Or iCodRegiao <> iUltCodRegiaoPer Or iMes <> iUltMes Then
        
        lErro = ObterRelMesFatRegiaoPer(iFilialEmpresa, iCodRegiao, iMes)
        If lErro <> SUCESSO Then gError 90253
    
    End If
    
    dValorFaturado = dUltValorFaturadoPer
    
    ObterValorFaturadoRegiaoPer = SUCESSO
    
    Exit Function
    
Erro_ObterValorFaturadoRegiaoPer:

    ObterValorFaturadoRegiaoPer = gErr
    
    Select Case gErr

        Case 90253
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153330)

    End Select

    Exit Function

End Function

Private Function ObterRelMesFatRegiaoPer(iFilialEmpresa As Integer, iCodRegiao As Integer, iMes As Integer) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long

Dim dtDataIni As Date
Dim dtDataFim As Date

Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double

On Error GoTo Erro_ObterRelMesFatRegiaoPer
    
    dtDataIni = CDate("01/" & iMes & "/" & IIf(iMes = 12, Year(gdtDataHoje) - 1, Year(gdtDataHoje)))
    dtDataFim = DateAdd("m", 1, dtDataIni) - 1
    
    lErro = Obter_QtdeDevFat_ValorDevFat_Regiao(iFilialEmpresa, iCodRegiao, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90254
    
    iUltFilialEmpresaPer = iFilialEmpresa
    iUltCodRegiaoPer = iCodRegiao
    iUltMes = iMes
    
    dUltQuantidadeDevolvidaPer = dQuantidadeDevolvida
    dUltQuantidadeFaturadaPer = dQuantidadeFaturada
    dUltValorDevolvidoPer = dValorDevolvido
    dUltValorFaturadoPer = dValorFaturado

    ObterRelMesFatRegiaoPer = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesFatRegiaoPer:

    ObterRelMesFatRegiaoPer = gErr
    
    Select Case gErr

        Case 90254
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153331)

    End Select

    Exit Function
    
End Function

'Maristela - Regiao por Data

Public Function ObterQtdeDevolvidaRegiaoAteData(iFilialEmpresa As Integer, iCodRegiao As Integer, dtData As Date, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida por uma Região

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaRegiaoAteData
    
    If iFilialEmpresa <> iUltFilialEmpresaData Or iCodRegiao <> iUltCodRegiaoData Or dtData <> dtUltData Then
        
        lErro = ObterRelMesFatRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData)
        If lErro <> SUCESSO Then gError 90260
    
    End If
    
    dQuantidadeDevolvida = dUltQuantidadeDevolvidaData
    
    ObterQtdeDevolvidaRegiaoAteData = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeDevolvidaRegiaoAteData:

    ObterQtdeDevolvidaRegiaoAteData = gErr
    
    Select Case gErr

        Case 90260
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153332)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaRegiaoAteData(iFilialEmpresa As Integer, iCodRegiao As Integer, dtData As Date, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada por uma Região no mês.

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaRegiaoAteData
    
    If iFilialEmpresa <> iUltFilialEmpresaData Or iCodRegiao <> iUltCodRegiaoData Or dtData <> dtUltData Then
        
        lErro = ObterRelMesFatRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData)
        If lErro <> SUCESSO Then gError 90261
    
    End If
    
    dQuantidadeFaturada = dUltQuantidadeFaturadaData
    
    ObterQtdeFaturadaRegiaoAteData = SUCESSO
    
    Exit Function
    
Erro_ObterQtdeFaturadaRegiaoAteData:

    ObterQtdeFaturadaRegiaoAteData = gErr
    
    Select Case gErr

        Case 90261
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153333)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoRegiaoAteData(iFilialEmpresa As Integer, iCodRegiao As Integer, dtData As Date, dValorDevolvido As Double) As Long
'Para obter o valor devolvido por uma Região no mês

Dim lErro As Long

On Error GoTo Erro_ObterValorDevolvidoRegiaoAteData
    
    If iFilialEmpresa <> iUltFilialEmpresaData Or iCodRegiao <> iUltCodRegiaoData Or dtData <> dtUltData Then
        
        lErro = ObterRelMesFatRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData)
        If lErro <> SUCESSO Then gError 90262
    
    End If
    
    dValorDevolvido = dUltValorDevolvidoData
    
    ObterValorDevolvidoRegiaoAteData = SUCESSO
    
    Exit Function
    
Erro_ObterValorDevolvidoRegiaoAteData:

    ObterValorDevolvidoRegiaoAteData = gErr
    
    Select Case gErr

        Case 90262
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153334)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoRegiaoAteData(iFilialEmpresa As Integer, iCodRegiao As Integer, dtData As Date, dValorFaturado As Double) As Long
'Para obter o valor faturado por uma Região no mês.

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoRegiaoAteData
    
    If iFilialEmpresa <> iUltFilialEmpresaData Or iCodRegiao <> iUltCodRegiaoData Or dtData <> dtUltData Then
        
        lErro = ObterRelMesFatRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData)
        If lErro <> SUCESSO Then gError 90263
    
    End If
    
    dValorFaturado = dUltValorFaturadoData
    
    ObterValorFaturadoRegiaoAteData = SUCESSO
    
    Exit Function
    
Erro_ObterValorFaturadoRegiaoAteData:

    ObterValorFaturadoRegiaoAteData = gErr
    
    Select Case gErr

        Case 90263
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153335)

    End Select

    Exit Function

End Function

Private Function ObterRelMesFatRegiaoAteData(iFilialEmpresa As Integer, iCodRegiao As Integer, dtData As Date) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dtDataIni As Date
Dim dtDataFim As Date

Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double


On Error GoTo Erro_ObterRelMesFatRegiaoAteData

    dtDataIni = CDate("01/" & Month(dtData) & "/" & Year(dtData))
    dtDataFim = dtData

    lErro = Obter_QtdeDevFat_ValorDevFat_Regiao(iFilialEmpresa, iCodRegiao, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90264
    
    iUltFilialEmpresaData = iFilialEmpresa
    iUltCodRegiaoData = iCodRegiao
    dtUltData = dtData
    
    dUltQuantidadeDevolvidaData = dQuantidadeDevolvida
    dUltQuantidadeFaturadaData = dQuantidadeFaturada
    dUltValorDevolvidoData = dValorDevolvido
    dUltValorFaturadoData = dValorFaturado

    ObterRelMesFatRegiaoAteData = SUCESSO
    
    Exit Function
    
Erro_ObterRelMesFatRegiaoAteData:

    ObterRelMesFatRegiaoAteData = gErr
    
    Select Case gErr

        Case 90264
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153336)

    End Select

    Exit Function
    
End Function

Function Obter_QtdeDevFat_ValorDevFat_Regiao(iFilialEmpresa As Integer, iCodRegiao As Integer, dtDataIni As Date, dtDataFim As Date, dQuantidadeDevolvida As Double, dQuantidadeFaturada As Double, dValorDevolvido As Double, dValorFaturado As Double, lComando As Long) As Long
'Obtem a Quantidade devolvida e faturada, Valor devolvido e faturado de uma FilialEmpresa(passada po parametro),
'em uma regiao(passada por parametro), em um mes(passado por parametro)


Dim lErro As Long
Dim dFator As Double
Dim dQuantidade As Double
Dim dValorTotal As Double
Dim iFaturamento As Integer
Dim sUnidadeMed As String, sUMBase As String
Dim iClasseUM As Integer
Dim iCodTipoDocInfo As Integer

On Error GoTo Erro_Obter_QtdeDevFat_ValorDevFat_Regiao

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90259
    End If
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sUMBase = String(STRING_UM_SIGLA, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
            
        'Acumula a Quantidade e o Valor Faturado dentro do mes da data passada por parametro, nao comparando a Filial Empresa.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
            "FROM TiposDocInfo, FiliaisClientes, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM " & _
            "WHERE FiliaisClientes.Regiao = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ? AND FiliaisClientes.CodFilial = NFiscal.FilialCli AND FiliaisClientes.CodCliente = NFiscal.Cliente AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7 AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
            "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
            dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iCodRegiao, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    Else
    
        'Acumula a Quantidade e o Valor Faturado dentro do mes da data passada por parametro.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
            "FROM TiposDocInfo, FiliaisClientes, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM " & _
            "WHERE NFiscal.FilialEmpresa = ? AND FiliaisClientes.Regiao = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ? AND FiliaisClientes.CodFilial = NFiscal.FilialCli AND FiliaisClientes.CodCliente = NFiscal.Cliente AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7 AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
            "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
            dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iFilialEmpresa, iCodRegiao, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90255

    Do While lErro = AD_SQL_SUCESSO
        
        'Faturada
        If iFaturamento = 1 Then
            dValorFaturado = dValorFaturado + dValorTotal
            
'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90256
                
                If iCodTipoDocInfo <> DOCINFO_NFISCP And iCodTipoDocInfo <> DOCINFO_NFISFCP Then
                    dQuantidadeFaturada = dQuantidadeFaturada + (dQuantidade * dFator)
                End If
'            End If
        
        Else
            'Devolvida
            dValorDevolvido = dValorDevolvido + dValorTotal
            
'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90257
                
                dQuantidadeDevolvida = dQuantidadeDevolvida + (dQuantidade * dFator)
'            End If
        
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 90258
           
    Loop
          
    Obter_QtdeDevFat_ValorDevFat_Regiao = SUCESSO
    
    Exit Function
    
Erro_Obter_QtdeDevFat_ValorDevFat_Regiao:

    Obter_QtdeDevFat_ValorDevFat_Regiao = gErr
    
    Select Case gErr

        Case 90259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 90255, 90258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 90256, 90257
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153337)

    End Select

    Call Comando_Fechar(lComando)
        
    Exit Function
    
End Function


