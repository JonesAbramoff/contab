VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFatCliPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iUltFilialEmpresaDia As Integer
Private lUltClienteDia As Long
Private sUltProdutoDia As String
Private dtUltDataDia As Date
Private iUltFilialEmpresaMes As Integer
Private lUltClienteMes As Long
Private sUltProdutoMes As String
Private dtUltDataMes As Date
Private iUltFilialEmpresaAno As Integer
Private lUltClienteAno As Long
Private sUltProdutoAno As String
Private dtUltDataAno As Date

Private dUltQuantidadeFaturadaDia As Double
Private dUltQuantidadeDevolvidaDia As Double
Private dUltValorDevolvidoDia As Double
Private dUltValorFaturadoDia As Double
Private dUltQuantidadeFaturadaMes As Double
Private dUltQuantidadeDevolvidaMes As Double
Private dUltValorDevolvidoMes As Double
Private dUltValorFaturadoMes As Double
Private dUltQuantidadeFaturadaAno As Double
Private dUltQuantidadeDevolvidaAno As Double
Private dUltValorDevolvidoAno As Double
Private dUltValorFaturadoAno As Double

Private lComando As Long

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

'***** 05 e 06/06 Maristela - filialempresa, cliente, produto, data

'***** No dia
Public Function ObterQtdeDevolvidaCliProdDia(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida no dia

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaCliProdDia

    If iFilialEmpresa <> iUltFilialEmpresaDia Or lCliente <> lUltClienteDia Or sProduto <> sUltProdutoDia Or dtData <> dtUltDataDia Then

        lErro = ObterRelCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90317

    End If

    dQuantidadeDevolvida = dUltQuantidadeDevolvidaDia

    ObterQtdeDevolvidaCliProdDia = SUCESSO

    Exit Function

Erro_ObterQtdeDevolvidaCliProdDia:

    ObterQtdeDevolvidaCliProdDia = gErr

    Select Case gErr

        Case 90317

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153085)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaCliProdDia(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada no dia

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaCliProdDia

    If iFilialEmpresa <> iUltFilialEmpresaDia Or lCliente <> lUltClienteDia Or sProduto <> sUltProdutoDia Or dtData <> dtUltDataDia Then

        lErro = ObterRelCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90318

    End If

    dQuantidadeFaturada = dUltQuantidadeFaturadaDia

    ObterQtdeFaturadaCliProdDia = SUCESSO

    Exit Function

Erro_ObterQtdeFaturadaCliProdDia:

    ObterQtdeFaturadaCliProdDia = gErr

    Select Case gErr

        Case 90318

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153086)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoCliProdDia(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dValorDevolvido As Double) As Long
'Para obter a valor devolvido no dia

Dim lErro As Long

On Error GoTo Erro_ObterValorDevolvidoCliProdDia

    If iFilialEmpresa <> iUltFilialEmpresaDia Or lCliente <> lUltClienteDia Or sProduto <> sUltProdutoDia Or dtData <> dtUltDataDia Then

        lErro = ObterRelCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90319

    End If

    dValorDevolvido = dUltValorDevolvidoDia

    ObterValorDevolvidoCliProdDia = SUCESSO

    Exit Function

Erro_ObterValorDevolvidoCliProdDia:

    ObterValorDevolvidoCliProdDia = gErr

    Select Case gErr

        Case 90319

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153087)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoCliProdDia(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dValorFaturado As Double) As Long
'Para obter a valor faturado no dia

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoCliProdDia

    If iFilialEmpresa <> iUltFilialEmpresaDia Or lCliente <> lUltClienteDia Or sProduto <> sUltProdutoDia Or dtData <> dtUltDataDia Then

        lErro = ObterRelCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90320

    End If

    dValorFaturado = dUltValorFaturadoDia

    ObterValorFaturadoCliProdDia = SUCESSO

    Exit Function

Erro_ObterValorFaturadoCliProdDia:

    ObterValorFaturadoCliProdDia = gErr

    Select Case gErr

        Case 90320

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153088)

    End Select

    Exit Function

End Function

'********No mes
Public Function ObterQtdeDevolvidaCliProdMes(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida no mes

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaCliProdMes

    If iFilialEmpresa <> iUltFilialEmpresaMes Or lCliente <> lUltClienteMes Or sProduto <> sUltProdutoMes Or dtData <> dtUltDataMes Then

        lErro = ObterRelCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90321

    End If

    dQuantidadeDevolvida = dUltQuantidadeDevolvidaMes

    ObterQtdeDevolvidaCliProdMes = SUCESSO

    Exit Function

Erro_ObterQtdeDevolvidaCliProdMes:

    ObterQtdeDevolvidaCliProdMes = gErr

    Select Case gErr

        Case 90321

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153089)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaCliProdMes(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada no mes

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaCliProdMes

    If iFilialEmpresa <> iUltFilialEmpresaMes Or lCliente <> lUltClienteMes Or sProduto <> sUltProdutoMes Or dtData <> dtUltDataMes Then

        lErro = ObterRelCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90322

    End If

    dQuantidadeFaturada = dUltQuantidadeFaturadaMes

    ObterQtdeFaturadaCliProdMes = SUCESSO

    Exit Function

Erro_ObterQtdeFaturadaCliProdMes:

    ObterQtdeFaturadaCliProdMes = gErr

    Select Case gErr

        Case 90322

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153090)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoCliProdMes(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dValorDevolvido As Double) As Long
'Para obter o valor devolvido no mes

Dim lErro As Long

On Error GoTo Erro_ObterValorDevolvidoCliProdMes
    
    If iFilialEmpresa <> iUltFilialEmpresaMes Or lCliente <> lUltClienteMes Or sProduto <> sUltProdutoMes Or dtData <> dtUltDataMes Then

        lErro = ObterRelCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90323

    End If

    dValorDevolvido = dUltValorDevolvidoMes

    ObterValorDevolvidoCliProdMes = SUCESSO

    Exit Function

Erro_ObterValorDevolvidoCliProdMes:

    ObterValorDevolvidoCliProdMes = gErr

    Select Case gErr

        Case 90323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153091)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoCliProdMes(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dValorFaturado As Double) As Long
'Para obter o valor faturado no mes

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoCliProdMes

    If iFilialEmpresa <> iUltFilialEmpresaMes Or lCliente <> lUltClienteMes Or sProduto <> sUltProdutoMes Or dtData <> dtUltDataMes Then

        lErro = ObterRelCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90324

    End If

    dValorFaturado = dUltValorFaturadoMes

    ObterValorFaturadoCliProdMes = SUCESSO

    Exit Function

Erro_ObterValorFaturadoCliProdMes:

    ObterValorFaturadoCliProdMes = gErr

    Select Case gErr

        Case 90324

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153092)

    End Select

    Exit Function

End Function

'*********No ano
Public Function ObterQtdeDevolvidaCliProdAno(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dQuantidadeDevolvida As Double) As Long
'Para obter a qtde devolvida no ano

Dim lErro As Long

On Error GoTo Erro_ObterQtdeDevolvidaCliProdAno
    
    If iFilialEmpresa <> iUltFilialEmpresaAno Or lCliente <> lUltClienteAno Or sProduto <> sUltProdutoAno Or dtData <> dtUltDataAno Then

        lErro = ObterRelCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90325

    End If

    dQuantidadeDevolvida = dUltQuantidadeDevolvidaAno

    ObterQtdeDevolvidaCliProdAno = SUCESSO

    Exit Function

Erro_ObterQtdeDevolvidaCliProdAno:

    ObterQtdeDevolvidaCliProdAno = gErr

    Select Case gErr

        Case 90325

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153093)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaCliProdAno(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dQuantidadeFaturada As Double) As Long
'Para obter a qtde faturada no ano

Dim lErro As Long

On Error GoTo Erro_ObterQtdeFaturadaCliProdAno

    If iFilialEmpresa <> iUltFilialEmpresaAno Or lCliente <> lUltClienteAno Or sProduto <> sUltProdutoAno Or dtData <> dtUltDataAno Then

        lErro = ObterRelCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90326

    End If

    dQuantidadeFaturada = dUltQuantidadeFaturadaAno

    ObterQtdeFaturadaCliProdAno = SUCESSO

    Exit Function

Erro_ObterQtdeFaturadaCliProdAno:

    ObterQtdeFaturadaCliProdAno = gErr

    Select Case gErr

        Case 90326

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153094)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoCliProdAno(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dValorDevolvido As Double) As Long
'Para obter a valor devolvido no ano

Dim lErro As Long

On Error GoTo Erro_ObterValorDevolvidoCliProdAno

    If iFilialEmpresa <> iUltFilialEmpresaAno Or lCliente <> lUltClienteAno Or sProduto <> sUltProdutoAno Or dtData <> dtUltDataAno Then

        lErro = ObterRelCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90327

    End If

    dValorDevolvido = dUltValorDevolvidoAno

    ObterValorDevolvidoCliProdAno = SUCESSO

    Exit Function

Erro_ObterValorDevolvidoCliProdAno:

    ObterValorDevolvidoCliProdAno = gErr

    Select Case gErr

        Case 90327

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153095)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoCliProdAno(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date, dValorFaturado As Double) As Long
'Para obter a valor faturado no ano

Dim lErro As Long

On Error GoTo Erro_ObterValorFaturadoCliProdAno
    
    If iFilialEmpresa <> iUltFilialEmpresaAno Or lCliente <> lUltClienteAno Or sProduto <> sUltProdutoAno Or dtData <> dtUltDataAno Then

        lErro = ObterRelCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData)
        If lErro <> SUCESSO Then gError 90328

    End If

    dValorFaturado = dUltValorFaturadoAno

    ObterValorFaturadoCliProdAno = SUCESSO

    Exit Function

Erro_ObterValorFaturadoCliProdAno:

    ObterValorFaturadoCliProdAno = gErr

    Select Case gErr

        Case 90328

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153096)

    End Select

    Exit Function

End Function

Private Function ObterRelCliProdDia(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double
Dim dtDataIni As Date
Dim dtDataFim As Date

On Error GoTo Erro_ObterRelCliProdDia

    dtDataIni = dtData
    dtDataFim = dtData

    lErro = Obter_QtdeDevFat_ValorDevFat_CliProdPer(iFilialEmpresa, lCliente, sProduto, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90329

    iUltFilialEmpresaDia = iFilialEmpresa
    lUltClienteDia = lCliente
    sUltProdutoDia = sProduto
    dtUltDataDia = dtData

    dUltQuantidadeDevolvidaDia = dQuantidadeDevolvida
    dUltQuantidadeFaturadaDia = dQuantidadeFaturada
    dUltValorDevolvidoDia = dValorDevolvido
    dUltValorFaturadoDia = dValorFaturado

    ObterRelCliProdDia = SUCESSO

    Exit Function

Erro_ObterRelCliProdDia:

    ObterRelCliProdDia = gErr

    Select Case gErr

        Case 90329

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153097)

    End Select

    Exit Function

End Function

Private Function ObterRelCliProdMes(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double
Dim dtDataIni As Date
Dim dtDataFim As Date

On Error GoTo Erro_ObterRelCliProdMes

    dtDataIni = CDate("01/" & Month(dtData) & "/" & Year(dtData))
    dtDataFim = DateAdd("m", 1, dtDataIni) - 1

    lErro = Obter_QtdeDevFat_ValorDevFat_CliProdPer(iFilialEmpresa, lCliente, sProduto, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90330

    iUltFilialEmpresaMes = iFilialEmpresa
    lUltClienteMes = lCliente
    sUltProdutoMes = sProduto
    dtUltDataMes = dtData

    dUltQuantidadeDevolvidaMes = dQuantidadeDevolvida
    dUltQuantidadeFaturadaMes = dQuantidadeFaturada
    dUltValorDevolvidoMes = dValorDevolvido
    dUltValorFaturadoMes = dValorFaturado

    ObterRelCliProdMes = SUCESSO

    Exit Function

Erro_ObterRelCliProdMes:

    ObterRelCliProdMes = gErr

    Select Case gErr

        Case 90330

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153098)

    End Select

    Exit Function

End Function

Private Function ObterRelCliProdAno(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtData As Date) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dQuantidadeDevolvida As Double
Dim dQuantidadeFaturada As Double
Dim dValorDevolvido As Double
Dim dValorFaturado As Double
Dim dtDataIni As Date
Dim dtDataFim As Date

On Error GoTo Erro_ObterRelCliProdAno

    dtDataIni = CDate("01/01/" & Year(dtData))
    dtDataFim = CDate("31/12/" & Year(dtData))

    lErro = Obter_QtdeDevFat_ValorDevFat_CliProdPer(iFilialEmpresa, lCliente, sProduto, dtDataIni, dtDataFim, dQuantidadeDevolvida, dQuantidadeFaturada, dValorDevolvido, dValorFaturado, lComando)
    If lErro <> SUCESSO Then gError 90331

    iUltFilialEmpresaAno = iFilialEmpresa
    lUltClienteAno = lCliente
    sUltProdutoAno = sProduto
    dtUltDataAno = dtData

    dUltQuantidadeDevolvidaAno = dQuantidadeDevolvida
    dUltQuantidadeFaturadaAno = dQuantidadeFaturada
    dUltValorDevolvidoAno = dValorDevolvido
    dUltValorFaturadoAno = dValorFaturado

    ObterRelCliProdAno = SUCESSO

    Exit Function

Erro_ObterRelCliProdAno:

    ObterRelCliProdAno = gErr

    Select Case gErr

        Case 90331

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153099)

    End Select

    Exit Function

End Function

Function Obter_QtdeDevFat_ValorDevFat_CliProdPer(iFilialEmpresa As Integer, lCliente As Long, sProduto As String, dtDataIni As Date, dtDataFim As Date, dQuantidadeDevolvida As Double, dQuantidadeFaturada As Double, dValorDevolvido As Double, dValorFaturado As Double, lComando As Long) As Long
'Obtem a Quantidade devolvida e faturada, Valor devolvido e faturado no Dia, Mes e Ano, de uma FilialEmpresa(passada po parametro),
'de um Cliente(passado por parametro), de um sProduto(passado por parametro), de uma data(passada por parametro)

Dim lErro As Long
Dim dFator As Double
Dim dQuantidade As Double
Dim dValorTotal As Double
Dim iFaturamento As Integer
Dim sUnidadeMed As String, sUMBase As String
Dim iClasseUM As Integer
Dim iCodTipoDocInfo As Integer

On Error GoTo Erro_Obter_QtdeDevFat_ValorDevFat_CliProdPer

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90332
    End If

    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sUMBase = String(STRING_UM_SIGLA, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        
        'Acumula a Quantidade e o Valor, devolvido e Faturado do mes da data passada por parametro, nao comparando a Filial Empresa.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
        "FROM TiposDocInfo, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM " & _
        "WHERE NFiscal.Cliente = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?  AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
        "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
        dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, lCliente, sProduto, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)

    Else
    
        'Acumula a Quantidade e o Valor, devolvido e Faturado do mes da data passada por parametro.
        lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal, TiposDocInfo.Faturamento, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, Produtos.ClasseUM, TiposDocInfo.Codigo " & _
        "FROM TiposDocInfo, ItensNFiscal, NFiscal, Produtos, UnidadesdeMedida, ClasseUM " & _
        "WHERE NFiscal.FilialEmpresa = ?  AND NFiscal.Cliente = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?  AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe " & _
        "GROUP BY ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase,Produtos.ClasseUM,TiposDocInfo.Faturamento, TiposDocInfo.Codigo ", _
        dQuantidade, dValorTotal, iFaturamento, sUnidadeMed, sUMBase, iClasseUM, iCodTipoDocInfo, iFilialEmpresa, lCliente, sProduto, dtDataIni, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90333

    Do While lErro = AD_SQL_SUCESSO

        'Faturada
        If iFaturamento = 1 Then
            dValorFaturado = dValorFaturado + dValorTotal

'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90334

                If iCodTipoDocInfo <> DOCINFO_NFISCP And iCodTipoDocInfo <> DOCINFO_NFISFCP Then
                    dQuantidadeFaturada = dQuantidadeFaturada + (dQuantidade * dFator)
                End If
'            End If

        Else
            'Devolvida
            dValorDevolvido = dValorDevolvido + dValorTotal

'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 90335

                dQuantidadeDevolvida = dQuantidadeDevolvida + (dQuantidade * dFator)
'            End If

        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90336

    Loop

    Obter_QtdeDevFat_ValorDevFat_CliProdPer = SUCESSO

    Exit Function

Erro_Obter_QtdeDevFat_ValorDevFat_CliProdPer:

    Obter_QtdeDevFat_ValorDevFat_CliProdPer = gErr

    Select Case gErr

        Case 90332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90333, 90336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 90334, 90335

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153100)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

