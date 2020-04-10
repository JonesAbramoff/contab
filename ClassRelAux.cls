VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ALTERAÇÕES INPAL
''Márcio 09/2000 - criação de três novas funções - incluido cliente e vendedor

''esta classe faz a interface com adrellib.dll recebendo as chamadas originadas por uso de EXT_FUNC no Forprint

Dim colObj As New Collection 'para guardar referencias para objetos que podem ser reutilizados entre chamadas de funcoes de apoio a excucao de relatorios (EXT_FUNC no Forprint)

Private gcolChv(1 To 64) As Collection
Private gValAcumChv(1 To 64) As Double

Private gValAcum As Double
Private gsUltChvAcum As String

Private gValSom As Double
Private gsUltChvSom As String

Private gaValAcum(1 To 64) As Double
Private gasUltChvAcum(1 To 64) As String
Private baZerarAcum(1 To 64) As Boolean

Private Function GuardaObj(sID As String, obj As Object) As Long
'guarda referencias para objetos que podem ser reutilizados entre chamadas de funcoes de apoio a excucao de relatorios
Dim lErro As Long
On Error GoTo Erro_GuardaObj

    Call colObj.Add(obj, sID)

    GuardaObj = SUCESSO

    Exit Function

Erro_GuardaObj:

    GuardaObj = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152785)

    End Select

    Exit Function

End Function

Private Function ObtemObj(sID As String) As Object
'obtem referencia para objeto armazenado na colecao global a classe

Dim lErro As Long
On Error GoTo Erro_ObtemObj

    Set ObtemObj = colObj.Item(sID)

    Exit Function

Erro_ObtemObj:

    Set ObtemObj = Nothing

    Select Case Err

        Case 5: 'o elemento nao estava na colecao

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152786)

    End Select

    Exit Function

End Function

Public Function Inicializa(ByVal lSistema As Long, ByVal lCodEmp As Long, ByVal sNomeEmpresa As String, ByVal iCodFilial As Integer, ByVal sNomeFilial As String, ByVal dtDataAtual As Date, ByVal lExecucaoInproc As Long) As Long
'funcao chamada por adrellib.dll (em VC) para inicializar ambiente de apoio a execucao de relatorios

Dim lErro As Long

On Error GoTo Erro_Inicializa

    If (lExecucaoInproc = 0) Then

        lErro = Sistema_Abrir_Rel(lSistema, lCodEmp, sNomeEmpresa, iCodFilial, sNomeFilial, dtDataAtual)
        If lErro <> SUCESSO Then Error 41559

    End If

    Inicializa = SUCESSO

    Exit Function

Erro_Inicializa:

    Inicializa = Err

    Select Case Err

        Case 41559

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152787)

    End Select

    Exit Function

End Function

Public Function Finaliza() As Long

    'libera os objetos da colecao
    Set colObj = Nothing

    Finaliza = 0 'SUCESSO

End Function

'********************************************************************************
'********************************************************************************
'********** FUNCOES QUE SAO CHAMADAS POR ADRELLIB *******************************
'********************************************************************************
'********************************************************************************

'o primeiro parametro é o retorno da funcao no relatorio (.tsk) e nao é byval, todos os outros sao.
'para guardar informacao entre uma chamada e outra utilize GuardaObj e ObtemObj acima

Public Function ObterUltEmpenhoProd(dtDataUltEmpenho As Date, ByVal sProduto As String) As Long
'para obter a data do ultimo empenho de um produto

Dim lErro As Long, objUltEmp As ClassUltEmpenho

On Error GoTo Erro_ObterUltEmpenhoProd

Set objUltEmp = ObtemObj("UltEmpenho")
    If objUltEmp Is Nothing Then

        Set objUltEmp = New ClassUltEmpenho
        lErro = GuardaObj("UltEmpenho", objUltEmp)
        If lErro <> SUCESSO Then Error 41558

    End If

    ObterUltEmpenhoProd = objUltEmp.ObterUltEmpenhoProd(sProduto, dtDataUltEmpenho)

    Exit Function

Erro_ObterUltEmpenhoProd:

    ObterUltEmpenhoProd = Err

    Select Case Err

        Case 41558

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152788)

    End Select

    Exit Function

End Function

Public Function ObterValorFatProdPer(dValorFaturado As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterValorFatProdPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then Error 43318

    End If

    ObterValorFatProdPer = objRelMesFat.ObterValorFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorFaturado)

    Exit Function

Erro_ObterValorFatProdPer:

    ObterValorFatProdPer = Err

    Select Case Err

        Case 43318

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152789)

    End Select

    Exit Function

End Function

Public Function ObterDescontoFatProdPer(dDesconto As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o valor de descontos no faturamento de uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterDescontoFatProdPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then Error 43322

    End If

    ObterDescontoFatProdPer = objRelMesFat.ObterDescontoFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dDesconto)

    Exit Function

Erro_ObterDescontoFatProdPer:

    ObterDescontoFatProdPer = Err

    Select Case Err

        Case 43322

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152790)

    End Select

    Exit Function

End Function

Public Function ObterQtdePedidaFatProdPer(dQuantidadePedida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para a qtde em pedidos de venda de uma filialempresa para um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterQtdePedidaFatProdPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then Error 43323

    End If

    ObterQtdePedidaFatProdPer = objRelMesFat.ObterQtdePedidaFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadePedida)

    Exit Function

Erro_ObterQtdePedidaFatProdPer:

    ObterQtdePedidaFatProdPer = Err

    Select Case Err

        Case 43323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152791)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFatProdPer(dQuantidadeFaturada As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde faturada por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterQtdeFatProdPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then Error 43323

    End If

    ObterQtdeFatProdPer = objRelMesFat.ObterQtdeFaturadaFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFatProdPer:

    ObterQtdeFatProdPer = Err

    Select Case Err

        Case 43323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152792)

    End Select

    Exit Function

End Function

Public Function ObterQtdeDevolvidaProdPer(dQuantidadeDevolvida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterQtdeDevolvidaProdPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then Error 43350

    End If

    ObterQtdeDevolvidaProdPer = objRelMesFat.ObterQtdeDevolvidaFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaProdPer:

    ObterQtdeDevolvidaProdPer = Err

    Select Case Err

        Case 43350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152793)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornCod(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterFornCod

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111870

    End If

    MapaCot_ObterFornCod = objRelMapaCot.MapaCot_ObterFornCod(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterFornCod:

    MapaCot_ObterFornCod = gErr

    Select Case gErr

        Case 111870

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152794)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornContato(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterFornContato

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111871

    End If

    MapaCot_ObterFornContato = objRelMapaCot.MapaCot_ObterFornContato(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterFornContato:

    MapaCot_ObterFornContato = gErr

    Select Case gErr

        Case 111871

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152795)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornNome(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterFornNome

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111872

    End If

    MapaCot_ObterFornNome = objRelMapaCot.MapaCot_ObterFornNome(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterFornNome:

    MapaCot_ObterFornNome = gErr

    Select Case gErr

        Case 111872

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152796)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornTel1(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterFornTel1

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111873

    End If

    MapaCot_ObterFornTel1 = objRelMapaCot.MapaCot_ObterFornTel1(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterFornTel1:

    MapaCot_ObterFornTel1 = gErr

    Select Case gErr

        Case 111873

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152797)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterFornTel2(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterFornTel2

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111874

    End If

    MapaCot_ObterFornTel2 = objRelMapaCot.MapaCot_ObterFornTel2(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterFornTel2:

    MapaCot_ObterFornTel2 = gErr

    Select Case gErr

        Case 111874

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152798)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCCod(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCCod

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111893

    End If

    MapaCot_ObterPCCod = objRelMapaCot.MapaCot_ObterPCCod(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCCod:

    MapaCot_ObterPCCod = gErr

    Select Case gErr

        Case 111893

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152799)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCData(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCData

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111894

    End If

    MapaCot_ObterPCData = objRelMapaCot.MapaCot_ObterPCData(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCData:

    MapaCot_ObterPCData = gErr

    Select Case gErr

        Case 111894

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152800)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCForn(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCForn

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111895

    End If

    MapaCot_ObterPCForn = objRelMapaCot.MapaCot_ObterPCForn(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCForn:

    MapaCot_ObterPCForn = gErr

    Select Case gErr

        Case 111895

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152801)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCQuant(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCQuant

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111896

    End If

    MapaCot_ObterPCQuant = objRelMapaCot.MapaCot_ObterPCQuant(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCQuant:

    MapaCot_ObterPCQuant = gErr

    Select Case gErr

        Case 111896

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152802)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCPreco(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCPreco

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111897

    End If

    MapaCot_ObterPCPreco = objRelMapaCot.MapaCot_ObterPCPreco(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCPreco:

    MapaCot_ObterPCPreco = gErr

    Select Case gErr

        Case 111897

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152803)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCDesconto(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCDesconto

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111898

    End If

    MapaCot_ObterPCDesconto = objRelMapaCot.MapaCot_ObterPCDesconto(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCDesconto:

    MapaCot_ObterPCDesconto = gErr

    Select Case gErr

        Case 111898

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152804)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCIPI(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCIPI

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111899

    End If

    MapaCot_ObterPCIPI = objRelMapaCot.MapaCot_ObterPCIPI(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCIPI:

    MapaCot_ObterPCIPI = gErr

    Select Case gErr

        Case 111899

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152805)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCCondPagto(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCCondPagto

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120000

    End If

    MapaCot_ObterPCCondPagto = objRelMapaCot.MapaCot_ObterPCCondPagto(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCCondPagto:

    MapaCot_ObterPCCondPagto = gErr

    Select Case gErr

        Case 120000

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152806)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCTaxaMoeda(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCTaxaMoeda

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120001

    End If

    MapaCot_ObterPCTaxaMoeda = objRelMapaCot.MapaCot_ObterPCTaxaMoeda(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCTaxaMoeda:

    MapaCot_ObterPCTaxaMoeda = gErr

    Select Case gErr

        Case 120001

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152807)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCDataEntrega(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCDataEntrega

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120002

    End If

    MapaCot_ObterPCDataEntrega = objRelMapaCot.MapaCot_ObterPCDataEntrega(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCDataEntrega:

    MapaCot_ObterPCDataEntrega = gErr

    Select Case gErr

        Case 120002

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152808)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCStatus(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCStatus

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120003

    End If

    MapaCot_ObterPCStatus = objRelMapaCot.MapaCot_ObterPCStatus(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCStatus:

    MapaCot_ObterPCStatus = gErr

    Select Case gErr

        Case 120003

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152809)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCCodigo(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCCodigo

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120004

    End If

    MapaCot_ObterRCCodigo = objRelMapaCot.MapaCot_ObterRCCodigo(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCCodigo:

    MapaCot_ObterRCCodigo = gErr

    Select Case gErr

        Case 120004

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152810)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCRequisitante(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCRequisitante

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120005

    End If

    MapaCot_ObterRCRequisitante = objRelMapaCot.MapaCot_ObterRCRequisitante(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCRequisitante:

    MapaCot_ObterRCRequisitante = gErr

    Select Case gErr

        Case 120005

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152811)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCObs(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCObs

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120006

    End If

    MapaCot_ObterRCObs = objRelMapaCot.MapaCot_ObterRCObs(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCObs:

    MapaCot_ObterRCObs = gErr

    Select Case gErr

        Case 120006

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152812)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCCCL(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCCCL

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120007

    End If

    MapaCot_ObterRCCCL = objRelMapaCot.MapaCot_ObterRCCCL(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCCCL:

    MapaCot_ObterRCCCL = gErr

    Select Case gErr

        Case 120007

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152813)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCQuant(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCQuant

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120008

    End If

    MapaCot_ObterRCQuant = objRelMapaCot.MapaCot_ObterRCQuant(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCQuant:

    MapaCot_ObterRCQuant = gErr

    Select Case gErr

        Case 120008

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152814)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCPrazo(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCPrazo

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120009
        
    End If

    MapaCot_ObterRCPrazo = objRelMapaCot.MapaCot_ObterRCPrazo(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCPrazo:

    MapaCot_ObterRCPrazo = gErr

    Select Case gErr

        Case 120009

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152815)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterRCUrgente(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterRCUrgente

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 120010

    End If

    MapaCot_ObterRCUrgente = objRelMapaCot.MapaCot_ObterRCUrgente(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterRCUrgente:

    MapaCot_ObterRCUrgente = gErr

    Select Case gErr

        Case 120010

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152816)

    End Select

    Exit Function

End Function


Public Function ObterQtdeVendEstProdPer(dQuantidadeVendida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde vendida por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterQtdeVendEstProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 43347

    End If

    ObterQtdeVendEstProdPer = objRelMesEst.ObterQtdeVendidaProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeVendida)

    Exit Function

Erro_ObterQtdeVendEstProdPer:

    ObterQtdeVendEstProdPer = Err

    Select Case Err

        Case 43347

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152817)

    End Select

    Exit Function

End Function

Public Function ObterQtdeEntradaProdPer(dQuantidadeEntrada As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde entrada em uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterQtdeEntradaProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 43355

    End If

    ObterQtdeEntradaProdPer = objRelMesEst.ObterQtdeEntradaProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeEntrada)

    Exit Function

Erro_ObterQtdeEntradaProdPer:

    ObterQtdeEntradaProdPer = Err

    Select Case Err

        Case 43355

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152818)

    End Select

    Exit Function

End Function

Public Function ObterQtdeSaidaProdPer(dQuantidadeSaida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde saida de uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterQtdeSaidaProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 43356

    End If

    ObterQtdeSaidaProdPer = objRelMesEst.ObterQtdeSaidaProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeSaida)

    Exit Function

Erro_ObterQtdeSaidaProdPer:

    ObterQtdeSaidaProdPer = Err

    Select Case Err

        Case 43356

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152819)

    End Select

    Exit Function

End Function

Public Function ObterValorEntradaProdPer(dValorEntrada As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o custo das entradas em uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterValorEntradaProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 43357

    End If

    ObterValorEntradaProdPer = objRelMesEst.ObterValorEntradaProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorEntrada)

    Exit Function

Erro_ObterValorEntradaProdPer:

    ObterValorEntradaProdPer = Err

    Select Case Err

        Case 43357

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152820)

    End Select

    Exit Function

End Function

Public Function ObterValorSaidaProdPer(dValorSaida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o custo das saidas uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterValorSaidaProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 43358

    End If

    ObterValorSaidaProdPer = objRelMesEst.ObterValorSaidaProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorSaida)

    Exit Function

Erro_ObterValorSaidaProdPer:

    ObterValorSaidaProdPer = Err

    Select Case Err

        Case 43358

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152821)

    End Select

    Exit Function

End Function

Public Function ObterTextoNotasPC(sTexto As String, ByVal lNumIntPC As Long, ByVal iSeq As Integer) As Long
'Para obter o texto da nota de sequencial "iSeq" do pedido de compra "lNumIntPC"
'retorna o texto em sTexto (OUTPUT)
'lNumIntPC eh o numero interno do pedido (INPUT)
'iSeq eh o sequencial da nota (INPUT)
'para + esclarecimentos, vide tabela "NotasPedCompras"

Dim lErro As Long, objRelPedCompras As ClassRelPedCompras, sTextoAux As String

On Error GoTo Erro_ObterTextoNotasPC

    Set objRelPedCompras = ObtemObj("RelPedCompras")
    
    If objRelPedCompras Is Nothing Then

        Set objRelPedCompras = New ClassRelPedCompras
        lErro = GuardaObj("RelPedCompras", objRelPedCompras)
        If lErro <> SUCESSO Then gError 111846

    End If

    ObterTextoNotasPC = objRelPedCompras.ObterTextoNotasPC(sTextoAux, lNumIntPC, iSeq)
    Call Copia_String(sTexto, sTextoAux)
    
    Exit Function

Erro_ObterTextoNotasPC:

    ObterTextoNotasPC = gErr

    Select Case gErr

        Case 111846

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152822)

    End Select

    Exit Function

End Function

Public Function ObterCustoTotalEstProdPer(dCustoTotal As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o custo do material vendido por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterCustoTotalEstProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 43348

    End If

    ObterCustoTotalEstProdPer = objRelMesEst.ObterCustoTotalProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dCustoTotal)

    Exit Function

Erro_ObterCustoTotalEstProdPer:

    ObterCustoTotalEstProdPer = Err

    Select Case Err

        Case 43348

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152823)

    End Select

    Exit Function

End Function

Public Function ObterQtdeMediaProdPer(dQuantidade As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde media em estoque de uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterQtdeMediaProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 41950

    End If

    ObterQtdeMediaProdPer = objRelMesEst.ObterEstMedProdPerQtde(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidade)

    Exit Function

Erro_ObterQtdeMediaProdPer:

    ObterQtdeMediaProdPer = Err

    Select Case Err

        Case 41950

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152824)

    End Select

    Exit Function

End Function

Public Function ObterValorMedioProdPer(dValor As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o valor medio em estoque de uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterValorMedioProdPer

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 41951

    End If

    ObterValorMedioProdPer = objRelMesEst.ObterEstMedProdPerVal(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValor)

    Exit Function

Erro_ObterValorMedioProdPer:

    ObterValorMedioProdPer = Err

    Select Case Err

        Case 41951

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152825)

    End Select

    Exit Function

End Function

Public Function ObterConsumoMedioProd(dQuantidade As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde consumida media nos ultimos 12 meses em uma filialempresa de um produto

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterConsumoMedioProd

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 41962

    End If

    ObterConsumoMedioProd = objRelMesEst.ObterQtdeConsProd(sProduto, iFilialEmpresa, dQuantidade)

    Exit Function

Erro_ObterConsumoMedioProd:

    ObterConsumoMedioProd = Err

    Select Case Err

        Case 41962

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152826)

    End Select

    Exit Function

End Function

Public Function ObterVendaMediaQtdeProd(dQuantidade As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde media vendida nos ultimos 12 meses em uma filialempresa de um produto

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterVendaMediaQtdeProd

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 41962

    End If

    ObterVendaMediaQtdeProd = objRelMesEst.ObterQtdeVendProd(sProduto, iFilialEmpresa, dQuantidade)

    Exit Function

Erro_ObterVendaMediaQtdeProd:

    ObterVendaMediaQtdeProd = Err

    Select Case Err

        Case 41962

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152827)

    End Select

    Exit Function

End Function

Public Function ObterCusto(dCusto As Double, ByVal iFilialEmpresa As Integer, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstAberto As Integer, ByVal iMesEstAberto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal sProduto As String) As Long
'Para obter o custo do produtos passado como parametro

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterCusto

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 55158

    End If

    ObterCusto = objRelMesEst.ObterRelCusto(dCusto, iFilialEmpresa, iApropriacaoCusto, iAnoEstAberto, iMesEstAberto, iAnoEstApurado, iMesEstApurado, sProduto)

    Exit Function

Erro_ObterCusto:

    ObterCusto = Err

    Select Case Err

        Case 55158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152828)

    End Select

    Exit Function

End Function

Public Function ObterCustoFilial(dCusto As Double, ByVal iFilialEmpresa As Integer, ByVal iApropriacaoCusto As Integer, ByVal sProduto As String) As Long
'Para obter o custo do produtos passado como parametro

Dim lErro As Long, objRelMesEst As ClassRelMesEst

On Error GoTo Erro_ObterCustoFilial

    Set objRelMesEst = ObtemObj("RelMesEst")
    If objRelMesEst Is Nothing Then

        Set objRelMesEst = New ClassRelMesEst
        lErro = GuardaObj("RelMesEst", objRelMesEst)
        If lErro <> SUCESSO Then Error 55158

    End If

    ObterCustoFilial = objRelMesEst.ObterRelCustoFilial(dCusto, iFilialEmpresa, iApropriacaoCusto, sProduto)

    Exit Function

Erro_ObterCustoFilial:

    ObterCustoFilial = Err

    Select Case Err

        Case 55158

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152829)

    End Select

    Exit Function

End Function

Public Function ObterQuantInicial(dQuantInicial As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'obtem  a qtde inicial de um produto/almoxarifado antes da data inicial
'dQuantInicial é a quantidade inicial do produto sProduto no Almoxarifado iAlmoxarifado antes da data inicial dtDataInicial, que é retornada
'Caso dQuantInicial = -1 significa que a quantidade não está disponível.
'sProduto - é o codigo do produto em questao
'iAlmoxarifado é o codigo do almoxarifado em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstAlm As ClassRelMesEstAlm

On Error GoTo Erro_ObterQuantInicial

    Set objRelMesEstAlm = ObtemObj("RelMesEstAlm")
    If objRelMesEstAlm Is Nothing Then

        Set objRelMesEstAlm = New ClassRelMesEstAlm
        lErro = GuardaObj("RelMesEstAlm", objRelMesEstAlm)
        If lErro <> SUCESSO Then Error 55078

    End If

    ObterQuantInicial = objRelMesEstAlm.ObterQuantInicial1(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dQuantInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)

    Exit Function

Erro_ObterQuantInicial:

    ObterQuantInicial = Err

    Select Case Err

        Case 55078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152830)

    End Select

    Exit Function

End Function

Public Function ObterValorInicial(dValorInicial As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'obtem o saldo(valor) inicial de um produto/almoxarifado antes da data inicial
'dValorInicial é o saldo (valor) inicial do produto sProduto no Almoxarifado iAlmoxarifado antes da data inicial dtDataInicial, que é retornado
'Caso dValorInicial = -1 significa que o valor não está disponível.
'sProduto - é o codigo do produto em questao
'iAlmoxarifado é o codigo do almoxarifado em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstAlm As ClassRelMesEstAlm

On Error GoTo Erro_ObterValorInicial

    Set objRelMesEstAlm = ObtemObj("RelMesEstAlm")
    If objRelMesEstAlm Is Nothing Then

        Set objRelMesEstAlm = New ClassRelMesEstAlm
        lErro = GuardaObj("RelMesEstAlm", objRelMesEstAlm)
        If lErro <> SUCESSO Then Error 55080

    End If

    ObterValorInicial = objRelMesEstAlm.ObterValorInicial1(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dValorInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)

    Exit Function

Erro_ObterValorInicial:

    ObterValorInicial = Err

    Select Case Err

        Case 55080

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152831)

    End Select

    Exit Function

End Function

Public Function Formata_CGC_CPF(sTextoFormatado As String, ByVal sTextoOriginal As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_CGC_CPF

    Select Case Len(Trim(sTextoOriginal))

        Case STRING_CPF 'CPF

            sTextoFormatado = Format(sTextoOriginal, "000\.000\.000-00")

        Case STRING_CGC  'CGC

            sTextoFormatado = Format(sTextoOriginal, "00\.000\.000\/0000-00")

        Case Else

            sTextoFormatado = sTextoOriginal

    End Select

    Formata_CGC_CPF = SUCESSO

    Exit Function

Erro_Formata_CGC_CPF:

    Formata_CGC_CPF = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152832)

    End Select

    Exit Function

End Function

Public Function ObterDataParcRec(dtData As Date, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter a data de vencimento da parcela a receber

Dim lErro As Long, objRelTitRec As ClassRelTitRec

On Error GoTo Erro_ObterDataParcRec

    Set objRelTitRec = ObtemObj("RelTitRec")
    If objRelTitRec Is Nothing Then

        Set objRelTitRec = New ClassRelTitRec
        lErro = GuardaObj("RelTitRec", objRelTitRec)
        If lErro <> SUCESSO Then Error 58221

    End If

    ObterDataParcRec = objRelTitRec.ObterDataParcRec(dtData, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterDataParcRec:

    ObterDataParcRec = Err

    Select Case Err

        Case 58221

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152833)

    End Select

    Exit Function

End Function

Public Function ObterValorParcRec(dValor As Double, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long, objRelTitRec As ClassRelTitRec

On Error GoTo Erro_ObterValorParcRec

    Set objRelTitRec = ObtemObj("RelTitRec")
    If objRelTitRec Is Nothing Then

        Set objRelTitRec = New ClassRelTitRec
        lErro = GuardaObj("RelTitRec", objRelTitRec)
        If lErro <> SUCESSO Then Error 58222

    End If

    ObterValorParcRec = objRelTitRec.ObterValorParcRec(dValor, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterValorParcRec:

    ObterValorParcRec = Err

    Select Case Err

        Case 58222

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152834)

    End Select

    Exit Function

End Function

Public Function ObterDataParcPag(dtData As Date, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter a data de vencimento da parcela a Pagar

Dim lErro As Long, objRelTitPag As ClassRelTitPag

On Error GoTo Erro_ObterDataParcPag

    Set objRelTitPag = ObtemObj("RelTitPag")
    If objRelTitPag Is Nothing Then

        Set objRelTitPag = New ClassRelTitPag
        lErro = GuardaObj("RelTitPag", objRelTitPag)
        If lErro <> SUCESSO Then Error 58235

    End If

    ObterDataParcPag = objRelTitPag.ObterDataParcPag(dtData, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterDataParcPag:

    ObterDataParcPag = Err

    Select Case Err

        Case 58235

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152835)

    End Select

    Exit Function

End Function

Public Function ObterValorParcPag(dValor As Double, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter o valor da parcela a Pagar

Dim lErro As Long, objRelTitPag As ClassRelTitPag

On Error GoTo Erro_ObterValorParcPag

    Set objRelTitPag = ObtemObj("RelTitPag")
    If objRelTitPag Is Nothing Then

        Set objRelTitPag = New ClassRelTitPag
        lErro = GuardaObj("RelTitPag", objRelTitPag)
        If lErro <> SUCESSO Then Error 58236

    End If

    ObterValorParcPag = objRelTitPag.ObterValorParcPag(dValor, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterValorParcPag:

    ObterValorParcPag = Err

    Select Case Err

        Case 58236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152836)

    End Select

    Exit Function

End Function

'??? criacao interrompida em 08/09/99 - Jones
''Public Function ObterSaldoAplicCCI(dSaldo As Double, ByVal dtData As Date, ByVal iCCI As Integer) As Long
''' retorna o saldo das aplicacoes feitas a partir da conta corrente identificada por iCCI na data informada
'''É o valor das Aplicacoes feitas até a data final do periodo menos resgates destas aplicacoes feitos até a data final.
''Dim lErro As Long, objRelTes As ClassRelTes
''
''On Error GoTo Erro_ObterSaldoAplicCCI
''
''    Set objRelTes = ObtemObj("RelTes")
''    If objRelTes Is Nothing Then
''
''        Set objRelTes = New ClassRelTes
''        lErro = GuardaObj("RelTes", objRelTes)
''        If lErro <> SUCESSO Then Error 56814
''
''    End If
''
''    ObterSaldoAplicCCI = objRelTes.ObterSaldoAplicCCI(dSaldo, dtData, iCCI)
''
''    Exit Function
''
''Erro_ObterSaldoAplicCCI:
''
''    ObterSaldoAplicCCI = Err
''
''    Select Case Err
''
''        Case 56814
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152837)
''
''    End Select
''
''    Exit Function
''
''End Function

Public Function ObterSituacaoTrib(sSituacaoTrib As String, ByVal sProduto As String, ByVal iTipoTribICMS As Integer) As Long
'Para obter o Código da Situação Tributária

Dim lErro As Long, objRelNF As ClassRelNF
Dim sAux As String

On Error GoTo Erro_ObterSituacaoTrib

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 58571

    End If

    ObterSituacaoTrib = objRelNF.ObterSituacaoTrib(sAux, sProduto, iTipoTribICMS)
    Call Copia_String(sSituacaoTrib, sAux)

    Exit Function

Erro_ObterSituacaoTrib:

    ObterSituacaoTrib = Err

    Select Case Err

        Case 58571 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152838)

    End Select

    Exit Function

End Function

Public Function ObterSaldoAdiantForn(dSaldo As Double, ByVal lFornecedor As Long) As Long
'Retorna o Saldo de Adiantamentos do Fornecedor

Dim lErro As Long, objRelTitPag As ClassRelTitPag

On Error GoTo Erro_ObterSaldoAdiantForn

    Set objRelTitPag = ObtemObj("RelTitPag")
    If objRelTitPag Is Nothing Then

        Set objRelTitPag = New ClassRelTitPag
        lErro = GuardaObj("RelTitPag", objRelTitPag)
        If lErro <> SUCESSO Then Error 59093

    End If

    ObterSaldoAdiantForn = objRelTitPag.Fornecedor_Obter_SaldoAdiantamentos(dSaldo, lFornecedor)

    Exit Function

Erro_ObterSaldoAdiantForn:

    ObterSaldoAdiantForn = Err

    Select Case Err

        Case 59093

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152839)

    End Select

    Exit Function

End Function

Public Function ObterSaldoCredForn(dSaldo As Double, ByVal lFornecedor As Long) As Long
'Retorna o Saldo de Créditos para com o Fornecedor

Dim lErro As Long, objRelTitPag As ClassRelTitPag

On Error GoTo Erro_ObterSaldoCredForn

    Set objRelTitPag = ObtemObj("RelTitPag")
    If objRelTitPag Is Nothing Then

        Set objRelTitPag = New ClassRelTitPag
        lErro = GuardaObj("RelTitPag", objRelTitPag)
        If lErro <> SUCESSO Then Error 59094

    End If

    ObterSaldoCredForn = objRelTitPag.Fornecedor_Obter_SaldoCreditos(dSaldo, lFornecedor)

    Exit Function

Erro_ObterSaldoCredForn:

    ObterSaldoCredForn = Err

    Select Case Err

        Case 59094

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152840)

    End Select

    Exit Function

End Function

Public Function ObterComissDoc(sDoc As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna um texto para identificar uma comissao

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sAux As String

On Error GoTo Erro_ObterComissDoc

    '??? MsgBox (" Doc: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then Error 59128

    End If

    ObterComissDoc = objRelComissao.Comissao_Obter_Doc(iTipoComiss, lNumIntComiss, sAux, dtDataGeracao)

    Call Copia_String(sDoc, sAux)

    '??? MsgBox (" sDoc = sDoc >> iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Exit Function

Erro_ObterComissDoc:

    ObterComissDoc = Err

    Select Case Err

        Case 59128

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152841)

    End Select

    Exit Function

End Function

Public Function ObterComissNomeRedCli(sNomeRedCli As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna o nome reduzido de um cliente associado à uma comissao

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sAux As String

On Error GoTo Erro_ObterComissNomeRedCli

    '??? MsgBox (" NomeRed: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then Error 59129

    End If

    ObterComissNomeRedCli = objRelComissao.Comissao_Obter_NomeRedCli(iTipoComiss, lNumIntComiss, sAux, dtDataGeracao)

    Call Copia_String(sNomeRedCli, sAux)

    '??? MsgBox (" FimNomered: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Exit Function

Erro_ObterComissNomeRedCli:

    ObterComissNomeRedCli = Err

    Select Case Err

        Case 59129

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152842)

    End Select

    Exit Function

End Function

Public Function ObterComissDocOrd(sDocOrd As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna um texto para ordenar uma comissao

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sAux As String

On Error GoTo Erro_ObterComissDocOrd

    '??? MsgBox (" DocOrd: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then Error 59130

    End If

    ObterComissDocOrd = objRelComissao.Comissao_Obter_DocOrd(iTipoComiss, lNumIntComiss, sAux, dtDataGeracao)

    Call Copia_String(sDocOrd, sAux)

    '??? MsgBox (" FimDocOrd: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Exit Function

Erro_ObterComissDocOrd:

    ObterComissDocOrd = Err

    Select Case Err

        Case 59130

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152843)

    End Select

    Exit Function

End Function

Public Function ObterComissDataDoc(dtDataDoc As Date, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna um texto para identificar uma comissao

Dim lErro As Long, objRelComissao As ClassRelComissao

On Error GoTo Erro_ObterComissDataDoc

    '??? MsgBox (" dataDoc: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then Error 59131

    End If

    ObterComissDataDoc = objRelComissao.Comissao_Obter_DataDoc(iTipoComiss, lNumIntComiss, dtDataDoc, dtDataGeracao)

    '??? MsgBox (" fim Data Doc: iTipo = " & CStr(iTipoComiss) & " NumInt = " & CStr(lNumIntComiss))
    Exit Function

Erro_ObterComissDataDoc:

    ObterComissDataDoc = Err

    Select Case Err

        Case 59131

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152844)

    End Select

    Exit Function

End Function

Public Function ObterComissNumPVCli(sNumPVCli As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna o numero do pedido associado à uma comissao se houver

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sAux As String

On Error GoTo Erro_ObterComissNumPVCli

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then gError 133095

    End If

    ObterComissNumPVCli = objRelComissao.Comissao_Obter_NumPVCli(iTipoComiss, lNumIntComiss, sAux, dtDataGeracao)

    Call Copia_String(sNumPVCli, sAux)


    Exit Function

Erro_ObterComissNumPVCli:

    ObterComissNumPVCli = gErr

    Select Case gErr

        Case 133095

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152845)

    End Select

    Exit Function

End Function

Public Function ObterSaldoAdiantFornTotal(dSaldo As Double, ByVal lFornInicial As Long, ByVal lFornFinal As Long) As Long
'Retorna o Total dos Saldos dos Adiantamentos dos Fornecedores

Dim lErro As Long, objRelTitPag As ClassRelTitPag

On Error GoTo Erro_ObterSaldoAdiantFornTotal

    Set objRelTitPag = ObtemObj("RelTitPag")
    If objRelTitPag Is Nothing Then

        Set objRelTitPag = New ClassRelTitPag
        lErro = GuardaObj("RelTitPag", objRelTitPag)
        If lErro <> SUCESSO Then Error 58630

    End If

    ObterSaldoAdiantFornTotal = objRelTitPag.Fornecedor_Obter_SaldoAdiantamentosTotal(lFornInicial, lFornFinal, dSaldo)

    Exit Function

Erro_ObterSaldoAdiantFornTotal:

    ObterSaldoAdiantFornTotal = Err

    Select Case Err

        Case 58630

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152846)

    End Select

    Exit Function

End Function

Public Function ObterSaldoCredFornTotal(dSaldo As Double, ByVal lFornInicial As Long, ByVal lFornFinal As Long) As Long
'Retorna o Total dos Saldos de Créditos para os Fornecedores

Dim lErro As Long, objRelTitPag As ClassRelTitPag

On Error GoTo Erro_ObterSaldoCredFornTotal

    Set objRelTitPag = ObtemObj("RelTitPag")
    If objRelTitPag Is Nothing Then

        Set objRelTitPag = New ClassRelTitPag
        lErro = GuardaObj("RelTitPag", objRelTitPag)
        If lErro <> SUCESSO Then Error 58631

    End If

    ObterSaldoCredFornTotal = objRelTitPag.Fornecedor_Obter_SaldoCreditosTotal(lFornInicial, lFornFinal, dSaldo)

    Exit Function

Erro_ObterSaldoCredFornTotal:

    ObterSaldoCredFornTotal = Err

    Select Case Err

        Case 58631

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152847)

    End Select

    Exit Function

End Function

Public Function Obter_PrimeiraCompra_Cliente(dtPrimeiraCompra As Date, ByVal lCliente As Long, ByVal iFilialCliente As Integer) As Long
'Para obter a primeira compra do cliente

Dim lErro As Long, objRelCliente As ClassRelCliente

On Error GoTo Erro_Obter_PrimeiraCompra_Cliente

    Set objRelCliente = ObtemObj("RelCliente")

    If objRelCliente Is Nothing Then

        Set objRelCliente = New ClassRelCliente

        lErro = GuardaObj("RelCliente", objRelCliente)
        If lErro <> SUCESSO Then Error 60455

    End If

    Obter_PrimeiraCompra_Cliente = objRelCliente.Obter_PrimeiraCompra_Cliente(dtPrimeiraCompra, lCliente, iFilialCliente)

    Exit Function

Erro_Obter_PrimeiraCompra_Cliente:

    Obter_PrimeiraCompra_Cliente = Err

    Select Case Err

        Case 60455 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152848)

    End Select

    Exit Function

End Function

Public Function Obter_UltimaCompra_Cliente(dtUltimaCompra As Date, ByVal lCliente As Long, ByVal iFilialCliente As Integer) As Long
'Para obter a Ultima compra do cliente

Dim lErro As Long, objRelCliente As ClassRelCliente

On Error GoTo Erro_Obter_UltimaCompra_Cliente

    Set objRelCliente = ObtemObj("RelCliente")
    If objRelCliente Is Nothing Then

        Set objRelCliente = New ClassRelCliente

        lErro = GuardaObj("RelCliente", objRelCliente)
        If lErro <> SUCESSO Then Error 60456

    End If

    Obter_UltimaCompra_Cliente = objRelCliente.Obter_UltimaCompra_Cliente(dtUltimaCompra, lCliente, iFilialCliente)

    Exit Function

Erro_Obter_UltimaCompra_Cliente:

    Obter_UltimaCompra_Cliente = Err

    Select Case Err

        Case 60456 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152849)

    End Select

    Exit Function

End Function


Public Function Obter_NumeroCompras_Cliente(dNumeroCompras As Double, ByVal lCliente As Long, ByVal iFilialCliente As Integer) As Long
'Para obter o Número de Compras do Cliente

Dim lErro As Long, objRelCliente As ClassRelCliente, lNumeroCompras As Long

On Error GoTo Erro_Obter_NumeroCompras_Cliente

    Set objRelCliente = ObtemObj("RelCliente")
    If objRelCliente Is Nothing Then

        Set objRelCliente = New ClassRelCliente

        lErro = GuardaObj("RelCliente", objRelCliente)
        If lErro <> SUCESSO Then Error 60457

    End If

    Obter_NumeroCompras_Cliente = objRelCliente.Obter_NumeroCompras_Cliente(lNumeroCompras, lCliente, iFilialCliente)
    dNumeroCompras = lNumeroCompras

    Exit Function

Erro_Obter_NumeroCompras_Cliente:

    Obter_NumeroCompras_Cliente = Err

    Select Case Err

        Case 60457 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152850)

    End Select

    Exit Function

End Function

Public Function ObterSldIniCCI(dSldIni As Double, ByVal iCodCCI As Integer, ByVal dtData As Date) As Long
'Para obter o saldo inicial de uma conta numa determinada data

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldIniCCI

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then Error 60756

    End If

    ObterSldIniCCI = objRelTes.ObterSldIni(iCodCCI, dtData, dSldIni)

    Exit Function

Erro_ObterSldIniCCI:

    ObterSldIniCCI = Err

    Select Case Err

        Case 60756

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152851)

    End Select

    Exit Function

End Function

Public Function ObterSldAplicCCI(dSldAplic As Double, ByVal iCodCCI As Integer, ByVal dtData As Date) As Long
'Para obter o saldo das aplicações associadas a uma conta numa determinada data

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldAplicCCI

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then Error 60757

    End If

    ObterSldAplicCCI = objRelTes.ObterSldAplic(iCodCCI, dtData, dSldAplic)

    Exit Function

Erro_ObterSldAplicCCI:

    ObterSldAplicCCI = Err

    Select Case Err

        Case 60757

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152852)

    End Select

    Exit Function

End Function

Public Function Obter_Classificacao_Fiscal(sClass As String, ByVal sLetra As String) As Long

Dim lErro As Long
Dim objClassiFiscal As ClassRelClassiFiscal
Dim sAux As String

On Error GoTo Erro_Obter_Classificacao_Fiscal

    Set objClassiFiscal = ObtemObj("ClassiFiscal")

    If objClassiFiscal Is Nothing Then

        Set objClassiFiscal = New ClassRelClassiFiscal

        lErro = GuardaObj("ClassiFiscal", objClassiFiscal)
        If lErro <> SUCESSO Then Error 61277

    End If

    Call objClassiFiscal.Obter_Classificacao_Fiscal(sAux, sLetra)

    Call Copia_String(sClass, sAux)
    
    Obter_Classificacao_Fiscal = SUCESSO

    Exit Function

Erro_Obter_Classificacao_Fiscal:

    Obter_Classificacao_Fiscal = Err

    Select Case Err

        Case 61277

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152853)

    End Select

    Exit Function

End Function

Public Function Obter_Letra(sLetra As String, ByVal sClass As String, ByVal sLetraVariavel As String, ByVal lNumIntDoc As Long) As Long
'para os itens da nota fiscal

Dim lErro As Long
Dim objClassiFiscal As ClassRelClassiFiscal
Dim sAux As String

On Error GoTo Erro_Obter_Letra

    Set objClassiFiscal = ObtemObj("ClassiFiscal")

    If objClassiFiscal Is Nothing Then

        Set objClassiFiscal = New ClassRelClassiFiscal

        lErro = GuardaObj("ClassiFiscal", objClassiFiscal)
        If lErro <> SUCESSO Then Error 61278

    End If

    Call objClassiFiscal.Obter_Letra(sAux, sClass, sLetraVariavel, lNumIntDoc)
    
    Call Copia_String(sLetra, sAux)
    
    Obter_Letra = SUCESSO

    Exit Function

Erro_Obter_Letra:

    Obter_Letra = Err

    Select Case Err

        Case 61278

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152854)

    End Select

    Exit Function

End Function

Public Function Obter_Letras_NF(dNumIntDocNF1 As Double, ByVal dNumIntDocNF As Double) As Long
'para os itens da nota fiscal

Dim lErro As Long
Dim objClassiFiscal As ClassRelClassiFiscal
Dim lNumIntDoc As Long

On Error GoTo Erro_Obter_Letras_NF

    Set objClassiFiscal = ObtemObj("ClassiFiscal")
    
    lNumIntDoc = dNumIntDocNF

    If objClassiFiscal Is Nothing Then

        Set objClassiFiscal = New ClassRelClassiFiscal

        lErro = GuardaObj("ClassiFiscal", objClassiFiscal)
        If lErro <> SUCESSO Then gError 61278

    End If

    Obter_Letras_NF = objClassiFiscal.Obter_Letras_NF(lNumIntDoc)

    Exit Function

Erro_Obter_Letras_NF:

    Obter_Letras_NF = gErr

    Select Case gErr

        Case 61278

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152854)

    End Select

    Exit Function

End Function

Public Function ObterSldDevCR(dSldDev As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo em valor de devolucoes/creditos em aberto no contas a receber

Dim lErro As Long
Dim objRelCR As ClassRelCR

On Error GoTo Erro_ObterSldDevCR

    Set objRelCR = ObtemObj("RelCR")
    If objRelCR Is Nothing Then

        Set objRelCR = New ClassRelCR
        lErro = GuardaObj("RelCR", objRelCR)
        If lErro <> SUCESSO Then Error 59363

    End If

    ObterSldDevCR = objRelCR.ObterSldDevCR(dSldDev, iFilialEmpresa)

    Exit Function

Erro_ObterSldDevCR:

    ObterSldDevCR = Err

    Select Case Err

        Case 59363

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152855)

    End Select

    Exit Function

End Function

Public Function ObterQtdDevCR(dQtdDev As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde de devolucoes/creditos em aberto no contas a receber

Dim lErro As Long
Dim objRelCR As ClassRelCR

On Error GoTo Erro_ObterQtdDevCR

    Set objRelCR = ObtemObj("RelCR")
    If objRelCR Is Nothing Then

        Set objRelCR = New ClassRelCR
        lErro = GuardaObj("RelCR", objRelCR)
        If lErro <> SUCESSO Then Error 59364

    End If

    ObterQtdDevCR = objRelCR.ObterQtdDevCR(dQtdDev, iFilialEmpresa)

    Exit Function

Erro_ObterQtdDevCR:

    ObterQtdDevCR = Err

    Select Case Err

        Case 59364

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152856)

    End Select

    Exit Function

End Function

Public Function ObterSldAdiantCR(dSldAdiant As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo em valor de adiantamentos em aberto no contas a receber

Dim lErro As Long
Dim objRelCR As ClassRelCR

On Error GoTo Erro_ObterSldAdiantCR

    Set objRelCR = ObtemObj("RelCR")
    If objRelCR Is Nothing Then

        Set objRelCR = New ClassRelCR
        lErro = GuardaObj("RelCR", objRelCR)
        If lErro <> SUCESSO Then Error 59365

    End If

    ObterSldAdiantCR = objRelCR.ObterSldAdiantCR(dSldAdiant, iFilialEmpresa)

    Exit Function

Erro_ObterSldAdiantCR:

    ObterSldAdiantCR = Err

    Select Case Err

        Case 59365

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152857)

    End Select

    Exit Function

End Function

Public Function ObterQtdAdiantCR(dQtdAdiant As Double, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde de adiantamentos em aberto no contas a receber

Dim lErro As Long
Dim objRelCR As ClassRelCR

On Error GoTo Erro_ObterQtdAdiantCR

    Set objRelCR = ObtemObj("RelCR")
    If objRelCR Is Nothing Then

        Set objRelCR = New ClassRelCR
        lErro = GuardaObj("RelCR", objRelCR)
        If lErro <> SUCESSO Then Error 59366

    End If

    ObterQtdAdiantCR = objRelCR.ObterQtdAdiantCR(dQtdAdiant, iFilialEmpresa)

    Exit Function

Erro_ObterQtdAdiantCR:

    ObterQtdAdiantCR = Err

    Select Case Err

        Case 59366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152858)

    End Select

    Exit Function

End Function

Public Function NFICMS_TemDebitoOuCredito(dCredito As Double, ByVal lNumIntNF As Long) As Long
'retorna em iCredito: 0:nem credito nem debito, 1:tem credito ou debito

Dim lErro As Long, objRelNF As ClassRelNF, iCredito As Integer

On Error GoTo Erro_NFICMS_TemDebitoOuCredito

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59441

    End If

    NFICMS_TemDebitoOuCredito = objRelNF.TemDebitoOuCredito(lNumIntNF, iCredito)
    dCredito = iCredito

    Exit Function

Erro_NFICMS_TemDebitoOuCredito:

    NFICMS_TemDebitoOuCredito = Err

    Select Case Err

        Case 59441 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152859)

    End Select

    Exit Function

End Function

Public Function NFICMS_ObterICMSBase(dValorBase As Double, ByVal lNumIntNF As Long) As Long
'retorna em dValorBase a base de calculo de icms que gerou debito ou credito

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NFICMS_ObterICMSBase

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59442

    End If

    NFICMS_ObterICMSBase = objRelNF.ObterICMSBase(lNumIntNF, dValorBase)

    Exit Function

Erro_NFICMS_ObterICMSBase:

    NFICMS_ObterICMSBase = Err

    Select Case Err

        Case 59442 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152860)

    End Select

    Exit Function

End Function

Public Function NFICMS_ObterICMSValor(dValor As Double, ByVal lNumIntNF As Long) As Long
'retorna em dValor o valor de icms que gerou debito ou credito

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NFICMS_ObterICMSValor

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59443

    End If

    NFICMS_ObterICMSValor = objRelNF.ObterICMSValor(lNumIntNF, dValor)

    Exit Function

Erro_NFICMS_ObterICMSValor:

    NFICMS_ObterICMSValor = Err

    Select Case Err

        Case 59443 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152861)

    End Select

    Exit Function

End Function

Public Function NFICMS_ObterICMSSubstBase(dValorBaseSubst As Double, ByVal lNumIntNF As Long) As Long
'retorna em dValorBaseSubst a base de calculo de icms de substituicao que gerou debito ou credito

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NFICMS_ObterICMSSubstBase

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59444

    End If

    NFICMS_ObterICMSSubstBase = objRelNF.ObterICMSSubstBase(lNumIntNF, dValorBaseSubst)

    Exit Function

Erro_NFICMS_ObterICMSSubstBase:

    NFICMS_ObterICMSSubstBase = Err

    Select Case Err

        Case 59444 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152862)

    End Select

    Exit Function

End Function

Public Function NFICMS_ObterICMSSubstValor(dValorSubst As Double, ByVal lNumIntNF As Long) As Long
'retorna em dValorSubst o valor de icms de substituicao que gerou debito ou credito

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NFICMS_ObterICMSSubstValor

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59445

    End If

    NFICMS_ObterICMSSubstValor = objRelNF.ObterICMSSubstValor(lNumIntNF, dValorSubst)

    Exit Function

Erro_NFICMS_ObterICMSSubstValor:

    NFICMS_ObterICMSSubstValor = Err

    Select Case Err

        Case 59445 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152863)

    End Select

    Exit Function

End Function

Public Function ObterSldIniCCI_Total(dSldIni As Double, ByVal dtData As Date) As Long
'Para obter o saldo inicial de todas

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldIniCCI_Total

    Set objRelTes = ObtemObj("RelTes")

    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 69075

    End If

    ObterSldIniCCI_Total = objRelTes.ObterSldIni_Total(dtData, dSldIni)

    Exit Function

Erro_ObterSldIniCCI_Total:

    ObterSldIniCCI_Total = gErr

    Select Case gErr

        Case 69075

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152864)

    End Select

    Exit Function

End Function

Public Function ObterSldAplicCCI_Total(dSldAplic As Double, ByVal dtData As Date) As Long
'Para obter o saldo das aplicações Total para Todas as Contas

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldAplicCCI_Total

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 69076

    End If

    ObterSldAplicCCI_Total = objRelTes.ObterSldAplic_Total(dtData, dSldAplic)

    Exit Function

Erro_ObterSldAplicCCI_Total:

    ObterSldAplicCCI_Total = gErr

    Select Case gErr

        Case 69076

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152865)

    End Select

    Exit Function

End Function

Public Function ObterSldEntrada(dSldEntrada As Double, ByVal iCodConta As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Para obter o saldo Entrada da Conta passada

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldEntrada

    Set objRelTes = ObtemObj("RelTes")

    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 69085

    End If

    ObterSldEntrada = objRelTes.ObterSldEntrada(dtDataInicial, dtDataFinal, iCodConta, dSldEntrada)

    Exit Function

Erro_ObterSldEntrada:

    ObterSldEntrada = gErr

    Select Case gErr

        Case 69085

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152866)

    End Select

    Exit Function

End Function


Public Function ObterSldSaida(dSldSaida As Double, ByVal iCodConta As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Para obter o saldo Saida da Conta passada

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldSaida

    Set objRelTes = ObtemObj("RelTes")

    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 69086

    End If

    ObterSldSaida = objRelTes.ObterSldSaida(dtDataInicial, dtDataFinal, iCodConta, dSldSaida)

    Exit Function

Erro_ObterSldSaida:

    ObterSldSaida = gErr

    Select Case gErr

        Case 69086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152867)

    End Select

    Exit Function

End Function

Public Function ObterPaginasNF(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal lMaximoDeLinhas As Long) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim lNumPagina As Long

On Error GoTo Erro_ObterPaginasNF

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 69091

    End If

    ObterPaginasNF = objRelNF.ObterPaginasNF(lNumIntDoc, lMaximoDeLinhas, lNumPagina)

    dNumPagina = lNumPagina

    Exit Function

Erro_ObterPaginasNF:

    ObterPaginasNF = gErr

    Select Case gErr

        Case 69091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152868)

    End Select

    Exit Function

End Function

Public Function ObterPaginasNFGuedes(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal iMaxLinhasPagina As Integer, ByVal iMaxCaracLinha As Integer, ByVal iLinhasDesconto As Integer, ByVal iLinhasSuframa As Integer) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim iNumPagina As Integer

On Error GoTo Erro_ObterPaginasNFGuedes

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 69091

    End If

    ObterPaginasNFGuedes = objRelNF.ObterPaginasNFGuedes(lNumIntDoc, iMaxLinhasPagina, iNumPagina, iMaxCaracLinha, iLinhasDesconto, iLinhasSuframa)

    dNumPagina = iNumPagina

    Exit Function

Erro_ObterPaginasNFGuedes:

    ObterPaginasNFGuedes = gErr

    Select Case gErr

        Case 69091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152869)

    End Select

    Exit Function

End Function

Public Function ApuracaoICMSRes_Saidas(dValor As Double, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'prestações com débito de imposto

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_Saidas

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_Saidas = objRelResApur.ApuracaoICMSRes_Saidas(dValor, dtDataInicio, dtDataFinal, iFilialEmpresa)

    Exit Function

Erro_ApuracaoICMSRes_Saidas:

    ApuracaoICMSRes_Saidas = gErr

    Select Case gErr

        Case 78027

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152870)

    End Select

    Exit Function

End Function

Public Function ApuracaoICMSRes_Entradas(dValor As Double, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'prestações com Crédito de imposto

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_Entradas

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78028

    End If

    ApuracaoICMSRes_Entradas = objRelResApur.ApuracaoICMSRes_Entradas(dValor, dtDataInicio, dtDataFinal, iFilialEmpresa)

    Exit Function

Erro_ApuracaoICMSRes_Entradas:

    ApuracaoICMSRes_Entradas = gErr

    Select Case gErr

        Case 78028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152871)

    End Select

    Exit Function

End Function

Public Function ApuracaoRes_Saidas(dValor As Double, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'prestações com débito de imposto (IPI)

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoRes_Saidas

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78063

    End If

    ApuracaoRes_Saidas = objRelResApur.ApuracaoRes_Saidas(dValor, dtDataInicio, dtDataFinal, iFilialEmpresa)

    Exit Function

Erro_ApuracaoRes_Saidas:

    ApuracaoRes_Saidas = gErr

    Select Case gErr

        Case 78063

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152872)

    End Select

    Exit Function

End Function

Public Function ApuracaoRes_Entradas(dValor As Double, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'prestações com Crédito de imposto

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoRes_Entradas

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78064

    End If

    ApuracaoRes_Entradas = objRelResApur.ApuracaoRes_Entradas(dValor, dtDataInicio, dtDataFinal, iFilialEmpresa)

    Exit Function

Erro_ApuracaoRes_Entradas:

    ApuracaoRes_Entradas = gErr

    Select Case gErr

        Case 78064

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152873)

    End Select

    Exit Function

End Function

Public Function ApuracaoRes_SaldoDev(dValor As Double, ByVal iLixo As Integer) As Long
'Devolve o Saldo devedor

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoRes_SaldoDev

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78029

    End If

    ApuracaoRes_SaldoDev = objRelResApur.ApuracaoRes_SaldoDev(dValor)

    Exit Function

Erro_ApuracaoRes_SaldoDev:

    ApuracaoRes_SaldoDev = gErr

    Select Case gErr

        Case 78029

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152874)

    End Select

    Exit Function

End Function

Public Function ApuracaoRes_SaldoCred(dValor As Double, ByVal iLixo As Integer) As Long
'Devolve o Saldo Credor

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoRes_SaldoCred

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78030

    End If

    ApuracaoRes_SaldoCred = objRelResApur.ApuracaoRes_SaldoCred(dValor)

    Exit Function

Erro_ApuracaoRes_SaldoCred:

    ApuracaoRes_SaldoCred = gErr

    Select Case gErr

        Case 78030

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152875)

    End Select

    Exit Function

End Function

Public Function SubTotalDebito(dValorPassado As Double, ByVal dValorRecebido As Double) As Long
'Acumula o Sub total e o devolve em dValorRecebido

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_SubTotalDebito

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78031

    End If

    SubTotalDebito = objRelResApur.SubTotalDebito(dValorRecebido, dValorPassado)

    Exit Function

Erro_SubTotalDebito:

    SubTotalDebito = gErr

    Select Case gErr

        Case 78031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152876)

    End Select

    Exit Function

End Function

Public Function SubTotalCredito(dValorPassado As Double, ByVal dValorRecebido As Double) As Long
'Acumula o Sub total e o devolve em dValorRecebido

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_SubTotalCredito

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78032

    End If

    SubTotalCredito = objRelResApur.SubTotalCredito(dValorRecebido, dValorPassado)

    Exit Function

Erro_SubTotalCredito:

    SubTotalCredito = gErr

    Select Case gErr

        Case 78032

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152877)

    End Select

    Exit Function

End Function

Public Function ApuracaoICMSRes_SalCredAnt(dValor As Double, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long
'Obtem o saldo credor do periodo anterior

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_SalCredAnt

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78033

    End If

    ApuracaoICMSRes_SalCredAnt = objRelResApur.ApuracaoICMSRes_SalCredAnt(dValor, dtDataDe, dtDataAte, iFilialEmpresa)

    Exit Function

Erro_ApuracaoICMSRes_SalCredAnt:

    ApuracaoICMSRes_SalCredAnt = gErr

    Select Case gErr

        Case 78033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152878)

    End Select

    Exit Function

End Function

Public Function ApuracaoRes_SalCredAnt(dValor As Double, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long
'Obtem o saldo credor do periodo anterior

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoRes_SalCredAnt

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 88024

    End If

    ApuracaoRes_SalCredAnt = objRelResApur.ApuracaoRes_SalCredAnt(dValor, dtDataDe, dtDataAte, iFilialEmpresa)

    Exit Function

Erro_ApuracaoRes_SalCredAnt:

    ApuracaoRes_SalCredAnt = gErr

    Select Case gErr

        Case 88024

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152879)

    End Select

    Exit Function

End Function

Public Function ApuracaoICMSRes_ImpRec(dValor As Double, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long
'Imposto a recoher (ainda vai ser feito)

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_ImpRec

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78034

    End If

    ApuracaoICMSRes_ImpRec = objRelResApur.ApuracaoICMSRes_ImpRec(dValor, dtDataDe, dtDataAte, iFilialEmpresa)

    Exit Function

Erro_ApuracaoICMSRes_ImpRec:

    ApuracaoICMSRes_ImpRec = gErr

    Select Case gErr

        Case 78034

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152880)

    End Select

    Exit Function

End Function

Public Function ApuracaoRes_ImpRecolhe(dValor As Double, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iFilialEmpresa As Integer) As Long
'Imposto a recoher (ainda vai ser feito)

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoRes_ImpRecolhe

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 88023

    End If

    ApuracaoRes_ImpRecolhe = objRelResApur.ApuracaoRes_ImpRecolhe(dValor, dtDataDe, dtDataAte, iFilialEmpresa)

    Exit Function

Erro_ApuracaoRes_ImpRecolhe:

    ApuracaoRes_ImpRecolhe = gErr

    Select Case gErr

        Case 88023

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152881)

    End Select

    Exit Function

End Function

Public Function Declan_EstoqueInicial(dValor As Double, ByVal dtDataInicial As Date, ByVal iFilialEmpresa As Integer) As Long
'Retorna o Valor do Estoque na data passada

Dim lErro As Long
Dim objRelDeclan As ClassRelDeclan

On Error GoTo Erro_Declan_EstoqueInicial

    Set objRelDeclan = ObtemObj("objRelDeclan")

    If objRelDeclan Is Nothing Then

        Set objRelDeclan = New ClassRelDeclan
        lErro = GuardaObj("objRelDeclan", objRelDeclan)
        If lErro <> SUCESSO Then gError 78076

    End If

    Declan_EstoqueInicial = objRelDeclan.Declan_EstoqueInicial(dValor, dtDataInicial, iFilialEmpresa)

    Exit Function

Erro_Declan_EstoqueInicial:

    Declan_EstoqueInicial = gErr

    Select Case gErr

        Case 78076

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152882)

    End Select

    Exit Function

End Function

Public Function Declan_EstoqueFinal(dValor As Double, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Retorna o Valor do Estoque na data passada

Dim lErro As Long
Dim objRelDeclan As ClassRelDeclan

On Error GoTo Erro_Declan_EstoqueFinal

    Set objRelDeclan = ObtemObj("objRelDeclan")

    If objRelDeclan Is Nothing Then

        Set objRelDeclan = New ClassRelDeclan
        lErro = GuardaObj("objRelDeclan", objRelDeclan)
        If lErro <> SUCESSO Then gError 78077

    End If

    Declan_EstoqueFinal = objRelDeclan.Declan_EstoqueFinal(dValor, dtDataFinal, iFilialEmpresa)

    Exit Function

Erro_Declan_EstoqueFinal:

    Declan_EstoqueFinal = gErr

    Select Case gErr

        Case 78077

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152883)

    End Select

    Exit Function

End Function

Public Function ObterSldIniCcl(dValor As Double, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sCcl As String) As Long
'Obtem saldo inicial CCl

Dim lErro As Long
Dim objRelMvPerCcl As ClassRelMvPerCcl

On Error GoTo Erro_ObterSldIniCcl

    Set objRelMvPerCcl = ObtemObj("objRelMvPerCcl")

    If objRelMvPerCcl Is Nothing Then

        Set objRelMvPerCcl = New ClassRelMvPerCcl
        lErro = GuardaObj("objRelMvPerCcl", objRelMvPerCcl)
        If lErro <> SUCESSO Then gError 78084

    End If

    ObterSldIniCcl = objRelMvPerCcl.ObterSldIniCcl(dValor, iFilialEmpresa, iExercicio, iPeriodo, sCcl)

    Exit Function

Erro_ObterSldIniCcl:

    ObterSldIniCcl = gErr

    Select Case gErr

        Case 78084

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152884)

    End Select

    Exit Function

End Function

Public Function ObterValorProduto(dValor As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal iAlmoxarifado As Integer) As Long
'Obtem saldo inicial CCl

Dim lErro As Long
Dim objRelInventario As ClassRelInventario

On Error GoTo Erro_ObterValorProduto

    Set objRelInventario = ObtemObj("objRelInventario")

    If objRelInventario Is Nothing Then

        Set objRelInventario = New ClassRelInventario
        lErro = GuardaObj("objRelInventario", objRelInventario)
        If lErro <> SUCESSO Then gError 78120

    End If

    ObterValorProduto = objRelInventario.ObterValorProduto(dValor, sProduto, iFilialEmpresa, iAlmoxarifado)

    Exit Function

Erro_ObterValorProduto:

    ObterValorProduto = gErr

    Select Case gErr

        Case 78120

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152885)

    End Select

    Exit Function

End Function

'***** Inicio - Márcio 09/2000 - criação de três novas funções - incluido cliente e vendedor *******

Public Function ObterQtdeDevolvidaProdVendCliPer(dQuantidadeDevolvida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iVendedor As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterQtdeDevolvidaProdVendCliPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then gError 500215

    End If

    ObterQtdeDevolvidaProdVendCliPer = objRelMesFat.ObterQtdeDevolvidaFatProdVendCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaProdVendCliPer:

    ObterQtdeDevolvidaProdVendCliPer = gErr

    Select Case gErr

        Case 500215

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152886)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFatProdVendCliPer(dQuantidadeFaturada As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iVendedor As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde faturada por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterQtdeFatProdVendCliPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then gError 500223

    End If

    ObterQtdeFatProdVendCliPer = objRelMesFat.ObterQtdeFaturadaFatProdVendCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFatProdVendCliPer:

    ObterQtdeFatProdVendCliPer = gErr

    Select Case gErr

        Case 500223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152887)

    End Select

    Exit Function

End Function

Public Function ObterValorFatProdVendCliPer(dValorFaturado As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iVendedor As Integer, ByVal lCliente As Long) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterValorFatProdVendCliPer

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then gError 500224

    End If

    ObterValorFatProdVendCliPer = objRelMesFat.ObterValorFatProdVendCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iVendedor, lCliente, dValorFaturado)

    Exit Function

Erro_ObterValorFatProdVendCliPer:

    ObterValorFatProdVendCliPer = gErr

    Select Case gErr

        Case 500224

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152888)

    End Select

    Exit Function

End Function

'***** Fim - Márcio 09/2000 - criação de três novas funções - incluido cliente e vendedor *******


Public Function Cliente_Obtem_SaldoAnterior(dSaldo As Double, ByVal lCliente As Long, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Retorna o Saldo Contabil a receber do Cliente

Dim lErro As Long, objRelMvPerCli As ClassRelMvPerCli

On Error GoTo Erro_Cliente_Obtem_SaldoAnterior

    Set objRelMvPerCli = ObtemObj("RelMvPerCli")
    If objRelMvPerCli Is Nothing Then

        Set objRelMvPerCli = New ClassRelMvPerCli
        lErro = GuardaObj("RelMvPerCli", objRelMvPerCli)
        If lErro <> SUCESSO Then gError 76441

    End If

    Cliente_Obtem_SaldoAnterior = objRelMvPerCli.Cliente_Obtem_SaldoAnterior(dSaldo, lCliente, dtData, iFilialEmpresa)

    Exit Function

Erro_Cliente_Obtem_SaldoAnterior:

    Cliente_Obtem_SaldoAnterior = gErr

    Select Case gErr

        Case 76441

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152889)

    End Select

    Exit Function

End Function


Public Function Fornecedor_Obtem_SaldoAnterior(dSaldo As Double, ByVal lFornecedor As Long, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Retorna o Saldo Contabil a receber do Fornecedor

Dim lErro As Long, objRelMvPerForn As ClassRelMvPerForn

On Error GoTo Erro_Fornecedor_Obtem_SaldoAnterior

    Set objRelMvPerForn = ObtemObj("RelMvPerForn")
    If objRelMvPerForn Is Nothing Then

        Set objRelMvPerForn = New ClassRelMvPerForn
        lErro = GuardaObj("RelMvPerForn", objRelMvPerForn)
        If lErro <> SUCESSO Then gError 76451

    End If

    Fornecedor_Obtem_SaldoAnterior = objRelMvPerForn.Fornecedor_Obtem_SaldoAnterior(dSaldo, lFornecedor, dtData, iFilialEmpresa)

    Exit Function

Erro_Fornecedor_Obtem_SaldoAnterior:

    Fornecedor_Obtem_SaldoAnterior = gErr

    Select Case gErr

        Case 76451

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152890)

    End Select

    Exit Function

End Function

Public Function CGC_CPF_Critica(dRetorno As Double, ByVal sCGCCPF As String) As Long
'retorna zero se vazio, 1 se OK, 2 se invalido

Dim lErro As Long

On Error GoTo Erro_CGC_CPF_Critica

    dRetorno = 2

    If Len(Trim(sCGCCPF)) = 0 Then

        dRetorno = 0

    Else

        lErro = Cgc_Critica_Ext(right$("00000000000000" & sCGCCPF, 14))
        If lErro = SUCESSO Then

            dRetorno = 1

        Else

            lErro = Cpf_Critica_Ext(right$("00000000000" & sCGCCPF, 11))
            If lErro = SUCESSO Then dRetorno = 1

        End If

    End If

    CGC_CPF_Critica = SUCESSO

    Exit Function

Erro_CGC_CPF_Critica:

    CGC_CPF_Critica = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152891)

    End Select

    Exit Function

End Function

Function Cgc_Critica_Ext(ByVal sCgc As String) As Long

Dim lErro As Long
Dim dCgc As Double
Dim iDig(1 To STRING_CGC) As Integer
Dim iDigVer(1 To 3) As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iSum As Integer
Dim iParcela As Integer
Dim dCgcMaximo As Double

On Error GoTo Erro_Cgc_Critica_Ext

    sCgc = Trim(sCgc)

    If Len(sCgc) > 0 Then

        If Not IsNumeric(sCgc) Then Error 6641

        dCgc = CDbl(sCgc)

        If dCgc <= 0 Then Error 6642

        dCgcMaximo = CDbl(String(STRING_CGC, "9"))

        If dCgc > dCgcMaximo Then Error 6897

        If dCgc <> Int(dCgc) Then Error 6643

        sCgc = CStr(dCgc)

        If Len(sCgc) < STRING_CGC Then
            sCgc = String(STRING_CGC - Len(sCgc), "0") & sCgc
        End If

        For iIndice = 1 To STRING_CGC
            iDig(iIndice) = CInt(Mid(sCgc, iIndice, 1))
        Next

        For iIndice2 = 1 To 2

            iSum = 0

            For iIndice = 1 To 3 + iIndice2
                iSum = iSum + (5 + iIndice2 - iIndice) * iDig(iIndice)
            Next

            For iIndice = 4 + iIndice2 To STRING_CGC - 2
                iSum = iSum + (13 + iIndice2 - iIndice) * iDig(iIndice)
            Next

            iSum = IIf(iIndice2 = 1, iSum, iSum + 2 * iDigVer(1))

            iDigVer(iIndice2) = ((iSum * 10) Mod 11) Mod 10

        Next

        iSum = 0

        For iIndice = 1 To 7

            iParcela = iDig(iIndice) * ((iIndice Mod 2) + 1)
            iSum = iSum + (iParcela Mod 10) + Int(iParcela / 10)

        Next

'       este o 3º DV foi eliminado em 94/95
'       iDigVer(3) = IIf(iSum / 10 = Int(iSum / 10), iSum / 10, Int(iSum / 10) + 1) * 10 - iSum

'       If iDigVer(1) <> iDig(STRING_CGC - 1) Or iDigVer(2) <> iDig(STRING_CGC) Or iDigVer(3) <> iDig(8) Then Error 6644

        If iDigVer(1) <> iDig(STRING_CGC - 1) Or iDigVer(2) <> iDig(STRING_CGC) Then Error 6644

    End If

    Cgc_Critica_Ext = SUCESSO

    Exit Function

Erro_Cgc_Critica_Ext:

    Cgc_Critica_Ext = Err

    Select Case Err

        Case 6641

        Case 6642

        Case 6643

        Case 6644

       Case 6897

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152892)

    End Select

    Exit Function

End Function

Function Cpf_Critica_Ext(ByVal sCpf As String) As Long

Dim lErro As Long
Dim dCpf As Double
Dim iDig(1 To STRING_CPF) As Integer
Dim iDigVer(1 To 2) As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iSum As Integer
Dim dCpfMaximo As Double

On Error GoTo Erro_Cpf_Critica_Ext

    sCpf = Trim(sCpf)

    If Len(sCpf) > 0 Then

        If Not IsNumeric(sCpf) Then Error 6624

        dCpf = CDbl(sCpf)

        If dCpf <= 0 Then Error 6626

        dCpfMaximo = CDbl(String(STRING_CPF, "9"))

        If dCpf > dCpfMaximo Then Error 6896

        If dCpf <> Int(dCpf) Then Error 6630

        sCpf = CStr(dCpf)

        If Len(sCpf) < STRING_CPF Then
            sCpf = String(STRING_CPF - Len(sCpf), "0") & sCpf
        End If

        For iIndice = 1 To STRING_CPF
            iDig(iIndice) = CInt(Mid(sCpf, iIndice, 1))
        Next

        For iIndice2 = 1 To 2

            iSum = 0

            For iIndice = 1 To STRING_CPF - 2
                iSum = iSum + (10 + iIndice2 - iIndice) * iDig(iIndice)
            Next

            iSum = IIf(iIndice2 = 1, iSum, iSum + 2 * iDigVer(1))

            iDigVer(iIndice2) = ((iSum * 10) Mod 11) Mod 10

        Next


        If iDigVer(1) <> iDig(STRING_CPF - 1) Or iDigVer(2) <> iDig(STRING_CPF) Then Error 6633

    End If

    Cpf_Critica_Ext = SUCESSO

    Exit Function

Erro_Cpf_Critica_Ext:

    Cpf_Critica_Ext = Err

    Select Case Err

        Case 6624

        Case 6626

        Case 6630

        Case 6633

        Case 6896

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152893)

    End Select

    Exit Function

End Function

Public Function MvPerConta_Obtem_Media(dMedia As Double, ByVal iPeriodoPedido As Integer, ByVal iExercicioPedido As Integer, ByVal iFilialEmpresa As Integer, ByVal sConta As String) As Long
'Obtem a média dos 12 periodos anteriores ao periodo pedido.

Dim lErro As Long, objRelMvPerCta As ClassRelMvPerCta

On Error GoTo Erro_MvPerConta_Obtem_Media

    Set objRelMvPerCta = ObtemObj("RelMvPerCta")
    If objRelMvPerCta Is Nothing Then

        Set objRelMvPerCta = New ClassRelMvPerCta
        lErro = GuardaObj("RelMvPerCta", objRelMvPerCta)
        If lErro <> SUCESSO Then gError 71419

    End If

    MvPerConta_Obtem_Media = objRelMvPerCta.MvPerConta_Obtem_Media(dMedia, iPeriodoPedido, iExercicioPedido, iFilialEmpresa, sConta)

    Exit Function

Erro_MvPerConta_Obtem_Media:

    MvPerConta_Obtem_Media = gErr

    Select Case gErr

        Case 71419

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152894)

    End Select

    Exit Function

End Function

Public Function MvPerConta_Obtem_Total(dTotal As Double, iPeriodoPedido As Integer, iExercicioPedido As Integer, iFilialEmpresa As Integer, sConta As String) As Long
'Obtem o Total dos 12 periodos anteriores ao periodo pedido.

Dim lErro As Long, objRelMvPerCta As ClassRelMvPerCta

On Error GoTo Erro_MvPerConta_Obtem_Total

    Set objRelMvPerCta = ObtemObj("RelMvPerCta")
    If objRelMvPerCta Is Nothing Then

        Set objRelMvPerCta = New ClassRelMvPerCta
        lErro = GuardaObj("RelMvPerCta", objRelMvPerCta)
        If lErro <> SUCESSO Then gError 88025

    End If

    MvPerConta_Obtem_Total = objRelMvPerCta.MvPerConta_Obtem_Total(dTotal, iPeriodoPedido, iExercicioPedido, iFilialEmpresa, sConta)

    Exit Function

Erro_MvPerConta_Obtem_Total:

    MvPerConta_Obtem_Total = gErr

    Select Case gErr

        Case 88025

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152895)

    End Select

    Exit Function

End Function

Public Function ObterSaldoProdDtAlmox(dQuantidade As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo em estoque da propria empresa de um produto em um almoxarifado em uma data

Dim lErro As Long, objRelInventario As ClassRelInventario

On Error GoTo Erro_ObterSaldoProdDtAlmox

    Set objRelInventario = ObtemObj("RelInventario")
    If objRelInventario Is Nothing Then

        Set objRelInventario = New ClassRelInventario
        lErro = GuardaObj("RelInventario", objRelInventario)
        If lErro <> SUCESSO Then gError 81615

    End If

    ObterSaldoProdDtAlmox = objRelInventario.ObterSaldoProdDtAlmox(dQuantidade, sProduto, iAlmoxarifado, dtData, iFilialEmpresa)

    Exit Function

Erro_ObterSaldoProdDtAlmox:

    ObterSaldoProdDtAlmox = gErr

    Select Case gErr

        Case 81615

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152896)

    End Select

    Exit Function

End Function

Public Function ObterSaldoProdDtAlmoxEmTerc(dQuantidade As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque em 3os alocada a um produto-almoxarifado em uma data

Dim lErro As Long, objRelInventario As ClassRelInventario

On Error GoTo Erro_ObterSaldoProdDtAlmoxEmTerc

    Set objRelInventario = ObtemObj("RelInventario")
    If objRelInventario Is Nothing Then

        Set objRelInventario = New ClassRelInventario
        lErro = GuardaObj("RelInventario", objRelInventario)
        If lErro <> SUCESSO Then gError 81616

    End If

    ObterSaldoProdDtAlmoxEmTerc = objRelInventario.ObterSaldoProdDtAlmoxEmTerc(dQuantidade, sProduto, iAlmoxarifado, dtData, iFilialEmpresa)

    Exit Function

Erro_ObterSaldoProdDtAlmoxEmTerc:

    ObterSaldoProdDtAlmoxEmTerc = gErr

    Select Case gErr

        Case 81616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152897)

    End Select

    Exit Function

End Function

Public Function ObterSaldoProdDtAlmoxDeTerc(dQuantidade As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque de 3os em um produto-almoxarifado em uma data

Dim lErro As Long, objRelInventario As ClassRelInventario

On Error GoTo Erro_ObterSaldoProdDtAlmoxDeTerc

    Set objRelInventario = ObtemObj("RelInventario")
    If objRelInventario Is Nothing Then

        Set objRelInventario = New ClassRelInventario
        lErro = GuardaObj("RelInventario", objRelInventario)
        If lErro <> SUCESSO Then gError 81617

    End If

    ObterSaldoProdDtAlmoxDeTerc = objRelInventario.ObterSaldoProdDtAlmoxDeTerc(dQuantidade, sProduto, iAlmoxarifado, dtData, iFilialEmpresa)

    Exit Function

Erro_ObterSaldoProdDtAlmoxDeTerc:

    ObterSaldoProdDtAlmoxDeTerc = gErr

    Select Case gErr

        Case 81617

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152898)

    End Select

    Exit Function

End Function

Public Function Comissao_Promotores_Periodo(dValorComissao As Double, dtDataInicial As Date, dtDataFinal As Date, iTipoVend As Integer) As Long
'Retorna o valor total de comissões a receber de um vendedor dentro de uma determinada faixa de tempo

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Comissao_Promotores_Periodo

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87567

    'Faz leitura nas tabelas Comissoes, Vendedores, TiposDeVendedor
    lErro = Comando_Executar(lComando, "SELECT SUM(Comissoes.Valor) FROM Comissoes, Vendedores, TiposDeVendedor WHERE Comissoes.CodVendedor = Vendedores.Codigo AND Vendedores.Tipo = TiposDeVendedor.Codigo AND Comissoes.DataGeracao <= ? AND Comissoes.DataGeracao >= ? AND Vendedores.Tipo = ?", _
    dValorComissao, dtDataFinal, dtDataInicial, iTipoVend)
    If lErro <> AD_SQL_SUCESSO Then gError 87568

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 87569

    'Fecha comando
    Call Comando_Fechar(lComando)

    Comissao_Promotores_Periodo = SUCESSO

    Exit Function

Erro_Comissao_Promotores_Periodo:

    Comissao_Promotores_Periodo = gErr

    Select Case gErr

        Case 87567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 87568, 87569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, "SELECT Comissoes.Valor FROM Comissoes INNER JOIN Vendedores ON Comissoes.CodVendedor = Vendedores.Codigo INNER JOIN TiposDeVendedor ON Vendedores.Tipo = TiposDeVendedor.Codigo WHERE Comissoes.DataGeracao < ? AND Comissoes.DataGeracao > ? AND Vendedores.Tipo = ?")

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152899)

    End Select

    'Fecha comando ---> saída por erro
    Call Comando_Fechar(lComando)

    Exit Function

End Function

' *** INÍCIO DO CÓDIGO A SER REVISADO ***

Public Function Obter_Quant_PC_Forn_Produto(dQtdPedCompra As Double, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório a quantidade em pedidos de compra em aberto para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Quant_PC_Forn_Produto

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 79993

    End If

    Obter_Quant_PC_Forn_Produto = objRelFornProduto.ObterRel_Quant_PC_Forn_Produto(dQtdPedCompra, iFilialEmpresa, sProduto, lFornecedor, iFilialForn)

    Exit Function

Erro_Obter_Quant_PC_Forn_Produto:

    Obter_Quant_PC_Forn_Produto = gErr

    Select Case gErr

        Case 79993

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152900)
    End Select

End Function

Public Function Obter_Interv_Ressup_Medio(dTempoRessupMed As Double, iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório o intervalo médio de ressuprimento para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Interv_Ressup_Medio

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 79995

    End If

    Obter_Interv_Ressup_Medio = objRelFornProduto.ObterRel_Interv_Ressup_Medio(dTempoRessupMed, iFilialEmpresa, lFornecedor, iFilialForn, sProduto)

    Exit Function

Erro_Obter_Interv_Ressup_Medio:

    Obter_Interv_Ressup_Medio = gErr

    Select Case gErr

        Case 79995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152901)

    End Select

    Exit Function

End Function

Public Function Obter_Dt_Ult_Compra_Forn_Produto(dtUltCompra As Date, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a data da última compra para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Dt_Ult_Compra_Forn_Produto

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 79997

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Dt_Ult_Compra_Forn_Produto = objRelFornProduto.ObterRel_Info_Ult_Compra(objFornecedorProdutoFF)

    dtUltCompra = objFornecedorProdutoFF.dtDataUltimaCompra

    Exit Function

Erro_Obter_Dt_Ult_Compra_Forn_Produto:

    Obter_Dt_Ult_Compra_Forn_Produto = gErr

    Select Case gErr

        Case 79997

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152902)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_Pedida_Ult_PC_Fechado(dQtdPedida As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a quantidade pedida no último pedido de compra feito para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Qtd_Pedida_Ult_PC_Fechado

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90500

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Qtd_Pedida_Ult_PC_Fechado = objRelFornProduto.ObterRel_Info_Ult_PC_Fechado(objFornecedorProdutoFF)

    'Guarda o valor obtido na variável que será retornada ao FORPRINT
    dQtdPedida = objFornecedorProdutoFF.dQuantPedida

    Exit Function

Erro_Obter_Qtd_Pedida_Ult_PC_Fechado:

    Obter_Qtd_Pedida_Ult_PC_Fechado = gErr

    Select Case gErr

        Case 90500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152903)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_Recebida_Ult_PC_Fechado(dQtdRecebida As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a quantidade recebida no último pedido de compra feito para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Qtd_Recebida_Ult_PC_Fechado

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90501

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Qtd_Recebida_Ult_PC_Fechado = objRelFornProduto.ObterRel_Info_Ult_PC_Fechado(objFornecedorProdutoFF)

    'Guarda o valor obtido na variável que será retornada ao FORPRINT
    dQtdRecebida = objFornecedorProdutoFF.dQuantRecebida

    Exit Function

Erro_Obter_Qtd_Recebida_Ult_PC_Fechado:

    Obter_Qtd_Recebida_Ult_PC_Fechado = gErr

    Select Case gErr

        Case 90501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152904)

    End Select

    Exit Function

End Function

Public Function Obter_Dt_Ult_PC_Fechado(dtUltPCFechado As Date, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a data do último pedido de compra feito para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Dt_Ult_PC_Fechado

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90503

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Dt_Ult_PC_Fechado = objRelFornProduto.ObterRel_Info_Ult_PC_Fechado(objFornecedorProdutoFF)

    'Guarda o valor obtido na variável que será retornada ao FORPRINT
    dtUltPCFechado = objFornecedorProdutoFF.dtDataPedido

    Exit Function

Erro_Obter_Dt_Ult_PC_Fechado:

    Obter_Dt_Ult_PC_Fechado = gErr

    Select Case gErr

        Case 90503

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152905)

    End Select

    Exit Function

End Function

Public Function Obter_Dt_Receb_Ult_PC_Fechado(dtRecebUltPCFechado As Date, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a data de recebimento do último pedido de compra feito para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Dt_Receb_Ult_PC_Fechado

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90502

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Dt_Receb_Ult_PC_Fechado = objRelFornProduto.ObterRel_Info_Ult_PC_Fechado(objFornecedorProdutoFF)

    'Guarda o valor obtido na variável que será retornada ao FORPRINT
    dtRecebUltPCFechado = objFornecedorProdutoFF.dtDataReceb

    Exit Function

Erro_Obter_Dt_Receb_Ult_PC_Fechado:

    Obter_Dt_Receb_Ult_PC_Fechado = gErr

    Select Case gErr

        Case 90502

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152906)

    End Select

    Exit Function

End Function

Public Function Obter_TotalProduto_Ult_Compra(dTotProduto As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório o valor total da última compra para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_TotalProduto_Ult_Compra

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90504

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_TotalProduto_Ult_Compra = objRelFornProduto.ObterRel_Info_Ult_Compra(objFornecedorProdutoFF)

    dTotProduto = objFornecedorProdutoFF.dPrecoTotal

    Exit Function

Erro_Obter_TotalProduto_Ult_Compra:

    Obter_TotalProduto_Ult_Compra = gErr

    Select Case gErr

        Case 90504

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152907)

    End Select

    Exit Function

End Function

Public Function Obter_Dt_Ult_Cotacao(dtUltCotacao As Date, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a data da última cotação para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Dt_Ult_Cotacao

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90519

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Dt_Ult_Cotacao = objRelFornProduto.ObterRel_Info_Ult_Cotacao(objFornecedorProdutoFF)

    dtUltCotacao = objFornecedorProdutoFF.dtDataUltimaCotacao

    Exit Function

Erro_Obter_Dt_Ult_Cotacao:

    Obter_Dt_Ult_Cotacao = gErr

    Select Case gErr

        Case 90519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152908)

    End Select

    Exit Function

End Function

Public Function Obter_CondPagto_DescRed_Ult_Cotacao(sCondPagto As String, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a descrição da condição de pagto última cotação para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto
Dim sAux As String

On Error GoTo Erro_Obter_CondPagto_DescRed_Ult_Cotacao

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90520

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_CondPagto_DescRed_Ult_Cotacao = objRelFornProduto.ObterRel_Info_Ult_Cotacao(objFornecedorProdutoFF)
    
    sAux = objFornecedorProdutoFF.sCondPagto
    
    Call Copia_String(sCondPagto, sAux)

    Exit Function

Erro_Obter_CondPagto_DescRed_Ult_Cotacao:

    Obter_CondPagto_DescRed_Ult_Cotacao = gErr

    Select Case gErr

        Case 90520

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152909)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_Ult_Cotacao(dQtdUltCotacao As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório a quantidade da última cotação para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_Qtd_Ult_Cotacao

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90521

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_Qtd_Ult_Cotacao = objRelFornProduto.ObterRel_Info_Ult_Cotacao(objFornecedorProdutoFF)

    dQtdUltCotacao = objFornecedorProdutoFF.dQuantUltimaCotacao

    Exit Function

Erro_Obter_Qtd_Ult_Cotacao:

    Obter_Qtd_Ult_Cotacao = gErr

    Select Case gErr

        Case 90521

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152910)

    End Select

    Exit Function

End Function

Public Function Obter_ValorUnit_Ult_Cotacao(dValorUnitUltCotacao As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório o valor unitário da última cotação para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_ValorUnit_Ult_Cotacao

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90522

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_ValorUnit_Ult_Cotacao = objRelFornProduto.ObterRel_Info_Ult_Cotacao(objFornecedorProdutoFF)

    dValorUnitUltCotacao = objFornecedorProdutoFF.dUltimaCotacao

    Exit Function

Erro_Obter_ValorUnit_Ult_Cotacao:

    Obter_ValorUnit_Ult_Cotacao = gErr

    Select Case gErr

        Case 90522

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152911)

    End Select

    Exit Function

End Function

Public Function Obter_TipoFrete_Ult_Cotacao(dTipoFrete As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Devolve para o relatório o valor unitário da última cotação para o produto / fornecedor / filialforn / filialempresa

Dim lErro As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim objRelFornProduto As ClassRelFornProduto

On Error GoTo Erro_Obter_TipoFrete_Ult_Cotacao

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFornProduto = ObtemObj("RelFornProduto")

    'Se o obj estiver vazio
    If objRelFornProduto Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFornProduto = New ClassRelFornProduto

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFornProduto", objRelFornProduto)
        If lErro <> SUCESSO Then gError 90523

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = lFornecedor
    objFornecedorProdutoFF.iFilialForn = iFilialForn
    objFornecedorProdutoFF.sProduto = sProduto

    Obter_TipoFrete_Ult_Cotacao = objRelFornProduto.ObterRel_Info_Ult_Cotacao(objFornecedorProdutoFF)

    dTipoFrete = objFornecedorProdutoFF.iTipoFreteUltimaCotacao

    Exit Function

Erro_Obter_TipoFrete_Ult_Cotacao:

    Obter_TipoFrete_Ult_Cotacao = gErr

    Select Case gErr

        Case 90523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152912)

    End Select

    Exit Function

End Function

' *** FIM DO CÓDIGO A SER REVISADO ***


Public Function FilFornFilEmp_Obter_NumCompras(dNumeroCompras As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório o numero de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_NumCompras

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84617

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_NumCompras = objRelFilFornFilEmp.Obter_NumCompras(objFilFornFilEmp)

    dNumeroCompras = objFilFornFilEmp.lNumeroCompras

    Exit Function

Erro_FilFornFilEmp_Obter_NumCompras:

    FilFornFilEmp_Obter_NumCompras = gErr

    Select Case gErr

        Case 84617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152913)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_ValorAcumulado(dValorAcumulado As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório o Valor Acumulado de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_ValorAcumulado

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84618

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_ValorAcumulado = objRelFilFornFilEmp.Obter_ValorAcumulado(objFilFornFilEmp)

    dValorAcumulado = objFilFornFilEmp.dValorAcumuladoCompras

    Exit Function

Erro_FilFornFilEmp_Obter_ValorAcumulado:

    FilFornFilEmp_Obter_ValorAcumulado = gErr

    Select Case gErr

        Case 84618

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152914)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_MediaCompra(dMediaCompra As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório a Media de compras do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_MediaCompra

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84619

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_MediaCompra = objRelFilFornFilEmp.Obter_MediaCompra(objFilFornFilEmp)

    dMediaCompra = objFilFornFilEmp.dMediaCompra

    Exit Function

Erro_FilFornFilEmp_Obter_MediaCompra:

    FilFornFilEmp_Obter_MediaCompra = gErr

    Select Case gErr

        Case 84619

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152915)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_NumDevolucoes(dNumDevolucoes As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório o numero de Devolucoes do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_NumDevolucoes

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84620

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_NumDevolucoes = objRelFilFornFilEmp.Obter_NumDevolucoes(objFilFornFilEmp)

    dNumDevolucoes = objFilFornFilEmp.lNumTotalDevolucoes

    Exit Function

Erro_FilFornFilEmp_Obter_NumDevolucoes:

    FilFornFilEmp_Obter_NumDevolucoes = gErr

    Select Case gErr

        Case 84620

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152916)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_UltimaCompra(dtUltimaCompra As Date, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório a Data da Última compra do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_UltimaCompra

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84621

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_UltimaCompra = objRelFilFornFilEmp.Obter_UltimaCompra(objFilFornFilEmp)

    dtUltimaCompra = objFilFornFilEmp.dtDataUltimaCompra
    
    Exit Function

Erro_FilFornFilEmp_Obter_UltimaCompra:

    FilFornFilEmp_Obter_UltimaCompra = gErr

    Select Case gErr

        Case 84621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152917)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_MediaAtraso(dMediaAtraso As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório a média de atrasos do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_MediaAtraso

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84622

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_MediaAtraso = objRelFilFornFilEmp.Obter_MediaAtraso(objFilFornFilEmp)

    If objFilFornFilEmp.lItensPedidosRecebidos <> 0 Then
        dMediaAtraso = objFilFornFilEmp.lAtrasoAcumulado / objFilFornFilEmp.lItensPedidosRecebidos
    Else
        dMediaAtraso = 0
    End If
    
    Exit Function

Erro_FilFornFilEmp_Obter_MediaAtraso:

    FilFornFilEmp_Obter_MediaAtraso = gErr

    Select Case gErr

        Case 84622

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152918)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_MaiorAtraso(dMaiorAtraso As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório o maior atraso do / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_MaiorAtraso

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84623

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_MaiorAtraso = objRelFilFornFilEmp.Obter_MaiorAtraso(objFilFornFilEmp)

    dMaiorAtraso = objFilFornFilEmp.lMaiorAtraso

    Exit Function

Erro_FilFornFilEmp_Obter_MaiorAtraso:

    FilFornFilEmp_Obter_MaiorAtraso = gErr

    Select Case gErr

        Case 84623

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152919)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_PedCompraAberto(dPedidosEmAberto As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório os Pedidos de Compras em Aberto / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_PedCompraAberto

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84624

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_PedCompraAberto = objRelFilFornFilEmp.Obter_PedCompraAberto(objFilFornFilEmp)

    dPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto
    
    Exit Function

Erro_FilFornFilEmp_Obter_PedCompraAberto:

    FilFornFilEmp_Obter_PedCompraAberto = gErr

    Select Case gErr

        Case 84624

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152920)

    End Select

    Exit Function

End Function

Public Function FilFornFilEmp_Obter_SaldoTitulos(dSaldoTitulos As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Devolve para o relatório os Saldos em Títulos / fornecedor / filialforn na filialempresa

Dim lErro As Long
Dim objRelFilFornFilEmp As ClassRelFilFornFilEmp
Dim objFilFornFilEmp As New ClassFilFornFilEmp

On Error GoTo Erro_FilFornFilEmp_Obter_SaldoTitulos

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelFilFornFilEmp = ObtemObj("RelFilFornFilEmp")

    'Se o obj estiver vazio
    If objRelFilFornFilEmp Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelFilFornFilEmp = New ClassRelFilFornFilEmp

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelFilFornFilEmp", objRelFilFornFilEmp)
        If lErro <> SUCESSO Then gError 84625

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objFilFornFilEmp.iFilialEmpresa = iFilialEmpresa
    objFilFornFilEmp.lCodFornecedor = lFornecedor
    objFilFornFilEmp.iCodFilial = iFilialForn

    FilFornFilEmp_Obter_SaldoTitulos = objRelFilFornFilEmp.Obter_SaldoTitulos(objFilFornFilEmp, dSaldoTitulos)
    
    Exit Function

Erro_FilFornFilEmp_Obter_SaldoTitulos:

    FilFornFilEmp_Obter_SaldoTitulos = gErr

    Select Case gErr

        Case 84625

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152921)

    End Select

    Exit Function

End Function

Public Function Comissao_TipoVendedor(dTotalComissoes As Double, dtDataIni As Date, dtDataFinal As Date, iTipoVendedor As Integer)

Dim lErro As Long
Dim objRelComisTipoVend As ClassRelComisTipoVend

On Error GoTo Erro_Comissao_TipoVendedor

    Set objRelComisTipoVend = ObtemObj("objRelComisTipoVend")

    If objRelComisTipoVend Is Nothing Then

        Set objRelComisTipoVend = New ClassRelComisTipoVend
        lErro = GuardaObj("objRelComisTipoVend", objRelComisTipoVend)
        If lErro <> SUCESSO Then gError 88032

    End If

    Comissao_TipoVendedor = objRelComisTipoVend.Le_Comissoes_PorGeracoes(dTotalComissoes, dtDataIni, dtDataFinal, iTipoVendedor)

    Exit Function

Erro_Comissao_TipoVendedor:

    Comissao_TipoVendedor = gErr

    Select Case gErr

        Case 88032

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152922)

    End Select

    Exit Function

End Function

Public Function Comissao_TipoVendedor_Vend(dTotalComissoes As Double, dtDataIni As Date, dtDataFinal As Date, iTipoVendedor As Integer, iCodVendedor As Integer) As Long

Dim lErro As Long
Dim objRelComisTipoVend As ClassRelComisTipoVend

On Error GoTo Erro_Comissao_TipoVendedor_Vend

    Set objRelComisTipoVend = ObtemObj("objRelComisTipoVend")

    If objRelComisTipoVend Is Nothing Then

        Set objRelComisTipoVend = New ClassRelComisTipoVend
        lErro = GuardaObj("objRelComisTipoVend", objRelComisTipoVend)
        If lErro <> SUCESSO Then gError 133009

    End If

    Comissao_TipoVendedor_Vend = objRelComisTipoVend.Le_Comissoes_PorGeracoes_Vendedor(dTotalComissoes, dtDataIni, dtDataFinal, iTipoVendedor, iCodVendedor)

    Exit Function

Erro_Comissao_TipoVendedor_Vend:

    Comissao_TipoVendedor_Vend = gErr

    Select Case gErr

        Case 133009

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152923)

    End Select

    Exit Function

End Function


'******************************************************************************************
'******************************************************************************************
'******************************************************************************************

'***** Maristela - filial cliente em um periodo

Public Function ObterQtdeDevolvidaProdFilCliPer(dQuantidadeDevolvida As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelProdFatFilCliPer As ClassRelProdFatFilCliPer

On Error GoTo Erro_ObterQtdeDevolvidaProdFilCliPer

    Set objRelProdFatFilCliPer = ObtemObj("RelProdFatFilCliPer")
    If objRelProdFatFilCliPer Is Nothing Then

        Set objRelProdFatFilCliPer = New ClassRelProdFatFilCliPer
        lErro = GuardaObj("RelProdFatFilCliPer", objRelProdFatFilCliPer)
        If lErro <> SUCESSO Then gError 90211

    End If

    ObterQtdeDevolvidaProdFilCliPer = objRelProdFatFilCliPer.ObterQtdeDevolvidaFatProdFilCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iFilialCli, lCliente, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaProdFilCliPer:

    ObterQtdeDevolvidaProdFilCliPer = gErr

    Select Case gErr

        Case 90211

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152924)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFatProdFilCliPer(dQuantidadeFaturada As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde faturada por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelProdFatFilCliPer As ClassRelProdFatFilCliPer

On Error GoTo Erro_ObterQtdeFatProdFilCliPer

    Set objRelProdFatFilCliPer = ObtemObj("RelProdFatFilCliPer")
    If objRelProdFatFilCliPer Is Nothing Then

        Set objRelProdFatFilCliPer = New ClassRelProdFatFilCliPer
        lErro = GuardaObj("RelProdFatFilCliPer", objRelProdFatFilCliPer)
        If lErro <> SUCESSO Then gError 90212

    End If

    ObterQtdeFatProdFilCliPer = objRelProdFatFilCliPer.ObterQtdeFaturadaFatProdFilCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iFilialCli, lCliente, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFatProdFilCliPer:

    ObterQtdeFatProdFilCliPer = gErr

    Select Case gErr

        Case 90212

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152925)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidaProdFilCliPer(dValorDevolvido As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter o valor devolvido por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelProdFatFilCliPer As ClassRelProdFatFilCliPer

On Error GoTo Erro_ObterValorDevolvidaProdFilCliPer

    Set objRelProdFatFilCliPer = ObtemObj("RelProdFatFilCliPer")
    If objRelProdFatFilCliPer Is Nothing Then

        Set objRelProdFatFilCliPer = New ClassRelProdFatFilCliPer
        lErro = GuardaObj("RelProdFatFilCliPer", objRelProdFatFilCliPer)
        If lErro <> SUCESSO Then gError 90683

    End If

    ObterValorDevolvidaProdFilCliPer = objRelProdFatFilCliPer.ObterValorDevolvidoFatProdFilCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iFilialCli, lCliente, dValorDevolvido)

    Exit Function

Erro_ObterValorDevolvidaProdFilCliPer:

    ObterValorDevolvidaProdFilCliPer = gErr

    Select Case gErr

        Case 90683

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152926)

    End Select

    Exit Function

End Function

Public Function ObterValorFatProdFilCliPer(dValorFaturado As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelProdFatFilCliPer As ClassRelProdFatFilCliPer

On Error GoTo Erro_ObterValorFatProdFilCliPer

    Set objRelProdFatFilCliPer = ObtemObj("RelProdFatFilCliPer")
    If objRelProdFatFilCliPer Is Nothing Then

        Set objRelProdFatFilCliPer = New ClassRelProdFatFilCliPer
        lErro = GuardaObj("RelProdFatFilCliPer", objRelProdFatFilCliPer)
        If lErro <> SUCESSO Then gError 90684

    End If

    ObterValorFatProdFilCliPer = objRelProdFatFilCliPer.ObterValorFaturadoFatProdFilCliPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, iFilialCli, lCliente, dValorFaturado)

    Exit Function

Erro_ObterValorFatProdFilCliPer:

    ObterValorFatProdFilCliPer = gErr

    Select Case gErr

        Case 90684

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152927)

    End Select

    Exit Function

End Function

'***** Maristela - filialEmpresa e data

Public Function ObterQtdeFaturadaFilEmpPer(dQuantidadeFaturada As Double, ByVal iFilialEmpresa As Integer, ByVal dtData As Date) As Long
'Para obter a quantidade faturada de uma filialempresa em um periodo

Dim lErro As Long, objRelFilEmpPer As ClassRelFilEmpPer

On Error GoTo Erro_ObterQtdeFaturadaFilEmpPer

    Set objRelFilEmpPer = ObtemObj("RelFilEmpPer")
    If objRelFilEmpPer Is Nothing Then

        Set objRelFilEmpPer = New ClassRelFilEmpPer
        lErro = GuardaObj("RelFilEmpPer", objRelFilEmpPer)
        If lErro <> SUCESSO Then gError 90270

    End If

    ObterQtdeFaturadaFilEmpPer = objRelFilEmpPer.ObterQtdeFaturadaFilEmpPer(iFilialEmpresa, dtData, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFaturadaFilEmpPer:

    ObterQtdeFaturadaFilEmpPer = gErr

    Select Case gErr

        Case 90270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152928)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoFilEmpPer(dValorFaturado As Double, ByVal iFilialEmpresa As Integer, ByVal dtData As Date) As Long
'Para obter o valor faturado de uma filialempresa em um periodo

Dim lErro As Long, objRelFilEmpPer As ClassRelFilEmpPer

On Error GoTo Erro_ObterValorFaturadoFilEmpPer

    Set objRelFilEmpPer = ObtemObj("RelFilEmpPer")
    If objRelFilEmpPer Is Nothing Then

        Set objRelFilEmpPer = New ClassRelFilEmpPer
        lErro = GuardaObj("RelFilEmpPer", objRelFilEmpPer)
        If lErro <> SUCESSO Then gError 90271

    End If

    ObterValorFaturadoFilEmpPer = objRelFilEmpPer.ObterValorFaturadoFilEmpPer(iFilialEmpresa, dtData, dValorFaturado)

    Exit Function

Erro_ObterValorFaturadoFilEmpPer:

    ObterValorFaturadoFilEmpPer = gErr

    Select Case gErr

        Case 90271

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152929)

    End Select

    Exit Function

End Function

'***** Maristela - Regiao e Mes

Public Function ObterQtdeDevolvidaRegiaoPer(dQuantidadeDevolvida As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal iMes As Integer) As Long
'Para obter a qtde devolvida por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterQtdeDevolvidaRegiaoPer

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90272

    End If

    ObterQtdeDevolvidaRegiaoPer = objRelRegiaoPer.ObterQtdeDevolvidaRegiaoPer(iFilialEmpresa, iCodRegiao, iMes, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaRegiaoPer:

    ObterQtdeDevolvidaRegiaoPer = gErr

    Select Case gErr

        Case 90272

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152930)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaRegiaoPer(dQuantidadeFaturada As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal iMes As Integer) As Long
'Para obter a qtde faturada por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterQtdeFaturadaRegiaoPer

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90273

    End If

    ObterQtdeFaturadaRegiaoPer = objRelRegiaoPer.ObterQtdeFaturadaRegiaoPer(iFilialEmpresa, iCodRegiao, iMes, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFaturadaRegiaoPer:

    ObterQtdeFaturadaRegiaoPer = gErr

    Select Case gErr

        Case 90273

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152931)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoRegiaoPer(dValorDevolvido As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal iMes As Integer) As Long
'Para obter a valor devolvido por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterValorDevolvidoRegiaoPer

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90274

    End If

    ObterValorDevolvidoRegiaoPer = objRelRegiaoPer.ObterValorDevolvidoRegiaoPer(iFilialEmpresa, iCodRegiao, iMes, dValorDevolvido)

    Exit Function

Erro_ObterValorDevolvidoRegiaoPer:

    ObterValorDevolvidoRegiaoPer = gErr

    Select Case gErr

        Case 90274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152932)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoRegiaoPer(dValorFaturado As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal iMes As Integer) As Long
'Para obter o valor faturado por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterValorFaturadoRegiaoPer

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90275

    End If

    ObterValorFaturadoRegiaoPer = objRelRegiaoPer.ObterValorFaturadoRegiaoPer(iFilialEmpresa, iCodRegiao, iMes, dValorFaturado)

    Exit Function

Erro_ObterValorFaturadoRegiaoPer:

    ObterValorFaturadoRegiaoPer = gErr

    Select Case gErr

        Case 90275

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152933)

    End Select

    Exit Function

End Function

'***** Maristela - Regiao e Data

Public Function ObterQtdeDevolvidaRegiaoAteData(dQuantidadeDevolvida As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal dtData As Date) As Long
'Para obter a qtde devolvida por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterQtdeDevolvidaRegiaoAteData

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90276

    End If

    ObterQtdeDevolvidaRegiaoAteData = objRelRegiaoPer.ObterQtdeDevolvidaRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaRegiaoAteData:

    ObterQtdeDevolvidaRegiaoAteData = gErr

    Select Case gErr

        Case 90276

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152934)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaRegiaoAteData(dQuantidadeFaturada As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal dtData As Date) As Long
'Para obter a qtde faturada por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterQtdeFaturadaRegiaoAteData

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90277

    End If

    ObterQtdeFaturadaRegiaoAteData = objRelRegiaoPer.ObterQtdeFaturadaRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFaturadaRegiaoAteData:

    ObterQtdeFaturadaRegiaoAteData = gErr

    Select Case gErr

        Case 90277

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152935)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoRegiaoAteData(dValorDevolvido As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal dtData As Date) As Long
'Para obter a valor devolvido por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterValorDevolvidoRegiaoAteData

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90278

    End If

    ObterValorDevolvidoRegiaoAteData = objRelRegiaoPer.ObterValorDevolvidoRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData, dValorDevolvido)

    Exit Function

Erro_ObterValorDevolvidoRegiaoAteData:

    ObterValorDevolvidoRegiaoAteData = gErr

    Select Case gErr

        Case 90278

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152936)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoRegiaoAteData(dValorFaturado As Double, ByVal iFilialEmpresa As Integer, ByVal iCodRegiao As Integer, ByVal dtData As Date) As Long
'Para obter o valor faturado por uma Região no mês

Dim lErro As Long, objRelRegiaoPer As ClassRelRegiaoPer

On Error GoTo Erro_ObterValorFaturadoRegiaoAteData

    Set objRelRegiaoPer = ObtemObj("RelRegiaoPer")
    If objRelRegiaoPer Is Nothing Then

        Set objRelRegiaoPer = New ClassRelRegiaoPer
        lErro = GuardaObj("RelRegiaoPer", objRelRegiaoPer)
        If lErro <> SUCESSO Then gError 90279

    End If

    ObterValorFaturadoRegiaoAteData = objRelRegiaoPer.ObterValorFaturadoRegiaoAteData(iFilialEmpresa, iCodRegiao, dtData, dValorFaturado)

    Exit Function

Erro_ObterValorFaturadoRegiaoAteData:

    ObterValorFaturadoRegiaoAteData = gErr

    Select Case gErr

        Case 90279

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152937)

    End Select

    Exit Function

End Function

'***** Maristela - Filial Cliente por mês

Public Function ObterQtdeDevFilCliMes(dQuantidadeDevolvida As Double, ByVal iMes As Integer, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde devolvida de uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterQtdeDevFilCliMes

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 90290

    End If

    ObterQtdeDevFilCliMes = objRelFatFilCliPer.ObterQtdeDevFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevFilCliMes:

    ObterQtdeDevFilCliMes = gErr

    Select Case gErr

        Case 90290

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152938)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obterqtdedevfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterQtdeDevFilCliMes2(dQuantidadeDevolvida As Double, ByVal dtDataFim As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde devolvida de uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterQtdeDevFilCliMes2

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 111710

    End If

    ObterQtdeDevFilCliMes2 = objRelFatFilCliPer.ObterQtdeDevFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevFilCliMes2:

    ObterQtdeDevFilCliMes2 = gErr

    Select Case gErr

        Case 111710

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152939)

    End Select

    Exit Function

End Function


Public Function ObterQtdeFatFilCliMes(dQuantidadeFaturada As Double, ByVal iMes As Integer, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde faturada por uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterQtdeFatFilCliMes

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 90291

    End If

    ObterQtdeFatFilCliMes = objRelFatFilCliPer.ObterQtdeFatFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFatFilCliMes:

    ObterQtdeFatFilCliMes = gErr

    Select Case gErr

        Case 90291

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152940)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obterqtdefatfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterQtdeFatFilCliMes2(dQuantidadeFaturada As Double, ByVal dtDataFim As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter a qtde faturada por uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterQtdeFatFilCliMes2

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 111711

    End If

    ObterQtdeFatFilCliMes2 = objRelFatFilCliPer.ObterQtdeFatFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFatFilCliMes2:

    ObterQtdeFatFilCliMes2 = gErr

    Select Case gErr

        Case 111711

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152941)

    End Select

    Exit Function

End Function

Public Function ObterValorDevFilCliMes(dValorDevolvido As Double, ByVal iMes As Integer, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter o valor devolvido por uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterValorDevFilCliMes

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 90292

    End If

    ObterValorDevFilCliMes = objRelFatFilCliPer.ObterValorDevFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes, dValorDevolvido)

    Exit Function

Erro_ObterValorDevFilCliMes:

    ObterValorDevFilCliMes = gErr

    Select Case gErr

        Case 90292

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152942)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obtervalordevfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterValorDevFilCliMes2(dValorDevolvido As Double, ByVal dtDataFim As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter o valor devolvido por uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterValorDevFilCliMes2

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 111712

    End If

    ObterValorDevFilCliMes2 = objRelFatFilCliPer.ObterValorDevFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim, dValorDevolvido)

    Exit Function

Erro_ObterValorDevFilCliMes2:

    ObterValorDevFilCliMes2 = gErr

    Select Case gErr

        Case 111712

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152943)

    End Select

    Exit Function

End Function

Public Function ObterValorFatFilCliMes(dValorFaturado As Double, ByVal iMes As Integer, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter o valor faturado por uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterValorFatFilCliMes

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 90293

    End If

    ObterValorFatFilCliMes = objRelFatFilCliPer.ObterValorFatFilCliMes(iFilialEmpresa, iFilialCli, lCliente, iMes, dValorFaturado)

    Exit Function

Erro_ObterValorFatFilCliMes:

    ObterValorFatFilCliMes = gErr

    Select Case gErr

        Case 90293

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152944)

    End Select

    Exit Function

End Function

'copiada por tulio271202
'a versao anterior (obtervalorfatfilclimes) soh recebia o mes...
'essa recebe a datafinal logo de cara ao inves de receber soh o mes
Public Function ObterValorFatFilCliMes2(dValorFaturado As Double, ByVal dtDataFim As Date, ByVal iFilialEmpresa As Integer, ByVal iFilialCli As Integer, ByVal lCliente As Long) As Long
'Para obter o valor faturado por uma filial cliente no mês

Dim lErro As Long, objRelFatFilCliPer As ClassRelFatFilCliPer

On Error GoTo Erro_ObterValorFatFilCliMes2

    Set objRelFatFilCliPer = ObtemObj("RelFatFilCliPer")
    If objRelFatFilCliPer Is Nothing Then

        Set objRelFatFilCliPer = New ClassRelFatFilCliPer
        lErro = GuardaObj("RelFatFilCliPer", objRelFatFilCliPer)
        If lErro <> SUCESSO Then gError 111713

    End If

    ObterValorFatFilCliMes2 = objRelFatFilCliPer.ObterValorFatFilCliMes2(iFilialEmpresa, iFilialCli, lCliente, dtDataFim, dValorFaturado)

    Exit Function

Erro_ObterValorFatFilCliMes2:

    ObterValorFatFilCliMes2 = gErr

    Select Case gErr

        Case 111713

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152945)

    End Select

    Exit Function

End Function

'**********Dia
Public Function ObterQtdeDevolvidaCliProdDia(dQuantidadeDevolvida As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto no dia

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterQtdeDevolvidaCliProdDia

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90305

    End If

    ObterQtdeDevolvidaCliProdDia = objRelFatCliPer.ObterQtdeDevolvidaCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaCliProdDia:

    ObterQtdeDevolvidaCliProdDia = gErr

    Select Case gErr

        Case 90305

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152946)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaCliProdDia(dQuantidadeFaturada As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter a qtde faturada por uma filialempresa de um produto no dia

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterQtdeFaturadaCliProdDia

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90306

    End If

    ObterQtdeFaturadaCliProdDia = objRelFatCliPer.ObterQtdeFaturadaCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFaturadaCliProdDia:

    ObterQtdeFaturadaCliProdDia = gErr

    Select Case gErr

        Case 90306

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152947)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoCliProdDia(dValorDevolvido As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter o valor devolvido por uma filialempresa de um produto no dia

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterValorDevolvidoCliProdDia

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90307

    End If

    ObterValorDevolvidoCliProdDia = objRelFatCliPer.ObterValorDevolvidoCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData, dValorDevolvido)

    Exit Function

Erro_ObterValorDevolvidoCliProdDia:

    ObterValorDevolvidoCliProdDia = gErr

    Select Case gErr

        Case 90307

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152948)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoCliProdDia(dValorFaturado As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter o valor faturado por uma filialempresa de um produto no dia

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterValorFaturadoCliProdDia

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90308

    End If

    ObterValorFaturadoCliProdDia = objRelFatCliPer.ObterValorFaturadoCliProdDia(iFilialEmpresa, lCliente, sProduto, dtData, dValorFaturado)

    Exit Function

Erro_ObterValorFaturadoCliProdDia:

    ObterValorFaturadoCliProdDia = gErr

    Select Case gErr

        Case 90308

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152949)

    End Select

    Exit Function

End Function

'*********** Mes
Public Function ObterQtdeDevolvidaCliProdMes(dQuantidadeDevolvida As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto no mes

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterQtdeDevolvidaCliProdMes

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90309

    End If

    ObterQtdeDevolvidaCliProdMes = objRelFatCliPer.ObterQtdeDevolvidaCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaCliProdMes:

    ObterQtdeDevolvidaCliProdMes = gErr

    Select Case gErr

        Case 90309

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152950)

    End Select

    Exit Function

End Function

Public Function ObterQtdeFaturadaCliProdMes(dQuantidadeFaturada As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter a qtde faturada por uma filialempresa de um produto no mes

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterQtdeFaturadaCliProdMes

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90310

    End If

    ObterQtdeFaturadaCliProdMes = objRelFatCliPer.ObterQtdeFaturadaCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFaturadaCliProdMes:

    ObterQtdeFaturadaCliProdMes = gErr

    Select Case gErr

        Case 90310

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152951)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoCliProdMes(dValorDevolvido As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter o valor devolvido por uma filialempresa de um produto no mes

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterValorDevolvidoCliProdMes

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90311

    End If

    ObterValorDevolvidoCliProdMes = objRelFatCliPer.ObterValorDevolvidoCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData, dValorDevolvido)

    Exit Function

Erro_ObterValorDevolvidoCliProdMes:

    ObterValorDevolvidoCliProdMes = gErr

    Select Case gErr

        Case 90311

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152952)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoCliProdMes(dValorFaturado As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter o valor faturado por uma filialempresa de um produto no mes

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterValorFaturadoCliProdMes

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90312

    End If

    ObterValorFaturadoCliProdMes = objRelFatCliPer.ObterValorFaturadoCliProdMes(iFilialEmpresa, lCliente, sProduto, dtData, dValorFaturado)

    Exit Function

Erro_ObterValorFaturadoCliProdMes:

    ObterValorFaturadoCliProdMes = gErr

    Select Case gErr

        Case 90312

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152953)

    End Select

    Exit Function

End Function

'**********Ano
Public Function ObterQtdeDevolvidaCliProdAno(dQuantidadeDevolvida As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter a qtde devolvida por uma filialempresa de um produto no ano

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterQtdeDevolvidaCliProdAno

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90313

    End If

    ObterQtdeDevolvidaCliProdAno = objRelFatCliPer.ObterQtdeDevolvidaCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData, dQuantidadeDevolvida)

    Exit Function

Erro_ObterQtdeDevolvidaCliProdAno:

    ObterQtdeDevolvidaCliProdAno = gErr

    Select Case gErr

        Case 90313

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152954)

    End Select

    Exit Function

End Function

'William(Inicio)

Public Function ObterDataVenctoReal(dtDataVenctoReal As Date, ByVal dtData As Date) As Long
'Obter data de vencto real desconsiderando sabados, domingos e feriados

Dim lErro As Long, objRelTes As ClassRelTes

On Error GoTo Erro_ObterDataVenctoReal

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 93346

    End If

    ObterDataVenctoReal = objRelTes.ObterDataVenctoReal(dtDataVenctoReal, dtData)

    Exit Function

Erro_ObterDataVenctoReal:

    ObterDataVenctoReal = gErr

    Select Case gErr

        Case 93346

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152955)

    End Select

    Exit Function

End Function
'(Fim)


Public Function DblParaTime(sHora As String, ByVal dValor As Double) As Long

Dim lErro As Long, objRelFat As ClassRelFat

On Error GoTo Erro_DblParaTime

    Set objRelFat = ObtemObj("RelFat")
    If objRelFat Is Nothing Then

        Set objRelFat = New ClassRelFat
        lErro = GuardaObj("RelFat", objRelFat)
        If lErro <> SUCESSO Then gError 93369

    End If

    DblParaTime = objRelFat.DblParaTime(sHora, dValor)

    Exit Function

Erro_DblParaTime:

    DblParaTime = gErr

    Select Case gErr

        Case 93369

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152956)

    End Select

    Exit Function

End Function


Public Function ObterQtdeFaturadaCliProdAno(dQuantidadeFaturada As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter a qtde faturada por uma filialempresa de um produto no ano

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterQtdeFaturadaCliProdAno

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90314

    End If

    ObterQtdeFaturadaCliProdAno = objRelFatCliPer.ObterQtdeFaturadaCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData, dQuantidadeFaturada)

    Exit Function

Erro_ObterQtdeFaturadaCliProdAno:

    ObterQtdeFaturadaCliProdAno = gErr

    Select Case gErr

        Case 90314

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152957)

    End Select

    Exit Function

End Function

Public Function ObterValorDevolvidoCliProdAno(dValorDevolvido As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter o valor devolvido por uma filialempresa de um produto no ano

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterValorDevolvidoCliProdAno

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90315

    End If

    ObterValorDevolvidoCliProdAno = objRelFatCliPer.ObterValorDevolvidoCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData, dValorDevolvido)

    Exit Function

Erro_ObterValorDevolvidoCliProdAno:

    ObterValorDevolvidoCliProdAno = gErr

    Select Case gErr

        Case 90315

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152958)

    End Select

    Exit Function

End Function

Public Function ObterValorFaturadoCliProdAno(dValorFaturado As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal sProduto As String, ByVal dtData As Date) As Long
'Para obter o valor faturado por uma filialempresa de um produto no ano

Dim lErro As Long, objRelFatCliPer As ClassRelFatCliPer

On Error GoTo Erro_ObterValorFaturadoCliProdAno

    Set objRelFatCliPer = ObtemObj("RelFatCliPer")
    If objRelFatCliPer Is Nothing Then

        Set objRelFatCliPer = New ClassRelFatCliPer
        lErro = GuardaObj("RelFatCliPer", objRelFatCliPer)
        If lErro <> SUCESSO Then gError 90316

    End If

    ObterValorFaturadoCliProdAno = objRelFatCliPer.ObterValorFaturadoCliProdAno(iFilialEmpresa, lCliente, sProduto, dtData, dValorFaturado)

    Exit Function

Erro_ObterValorFaturadoCliProdAno:

    ObterValorFaturadoCliProdAno = gErr

    Select Case gErr

        Case 90316

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152959)

    End Select

    Exit Function

End Function

'***** 06/06/01 Livro modelo 3 - Maristela (Testar)

Public Function ItemNF_ObterIPIBase(dIPIBaseCalculo As Double, ByVal lNumIntDoc As Long) As Long
'Para obter a Base de Calculo

Dim lErro As Long, objRelItemNFInfo As ClassRelItemNFInfo

On Error GoTo Erro_ItemNF_ObterIPIBase

    Set objRelItemNFInfo = ObtemObj("RelItemNFInfo")
    If objRelItemNFInfo Is Nothing Then

        Set objRelItemNFInfo = New ClassRelItemNFInfo
        lErro = GuardaObj("RelItemNFInfo", objRelItemNFInfo)
        If lErro <> SUCESSO Then gError 90342

    End If

    ItemNF_ObterIPIBase = objRelItemNFInfo.ItemNF_ObterIPIBase(lNumIntDoc, dIPIBaseCalculo)

    Exit Function

Erro_ItemNF_ObterIPIBase:

    ItemNF_ObterIPIBase = gErr

    Select Case gErr

        Case 90342

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152960)

    End Select

    Exit Function

End Function

Public Function ItemNF_ObterIPICredito(dIPICredito As Double, ByVal lNumIntDoc As Long) As Long
'Para obter Credito

Dim lErro As Long, objRelItemNFInfo As ClassRelItemNFInfo

On Error GoTo Erro_ItemNF_ObterIPICredito

    Set objRelItemNFInfo = ObtemObj("RelItemNFInfo")
    If objRelItemNFInfo Is Nothing Then

        Set objRelItemNFInfo = New ClassRelItemNFInfo
        lErro = GuardaObj("RelItemNFInfo", objRelItemNFInfo)
        If lErro <> SUCESSO Then gError 90343

    End If

    ItemNF_ObterIPICredito = objRelItemNFInfo.ItemNF_ObterIPICredito(lNumIntDoc, dIPICredito)

    Exit Function

Erro_ItemNF_ObterIPICredito:

    ItemNF_ObterIPICredito = gErr

    Select Case gErr

        Case 90343

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152961)

    End Select

    Exit Function

End Function

'***** 11/06/01 Livro modelo 3 - Maristela (Testar)
Public Function ObterSaldoProdDtFil(dQuantidade As Double, ByVal sProduto As String, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo em estoque da propria empresa de um produto em uma data

Dim lErro As Long, objRelInventarioFil As ClassRelInventarioFil

On Error GoTo Erro_ObterSaldoProdDtFil

    Set objRelInventarioFil = ObtemObj("RelInventarioFil")
    If objRelInventarioFil Is Nothing Then

        Set objRelInventarioFil = New ClassRelInventarioFil
        lErro = GuardaObj("RelInventarioFil", objRelInventarioFil)
        If lErro <> SUCESSO Then gError 90368

    End If

    ObterSaldoProdDtFil = objRelInventarioFil.ObterSaldoProdDtFil(dQuantidade, sProduto, dtData, iFilialEmpresa)

    Exit Function

Erro_ObterSaldoProdDtFil:

    ObterSaldoProdDtFil = gErr

    Select Case gErr

        Case 90368

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152962)

    End Select

    Exit Function

End Function

Public Function ObterSaldoProdDtFilDeTerc(dQuantidade As Double, ByVal sProduto As String, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque de 3os em estoque de um produto em uma data

Dim lErro As Long, objRelInventarioFil As ClassRelInventarioFil

On Error GoTo Erro_ObterSaldoProdDtFilDeTerc

    Set objRelInventarioFil = ObtemObj("RelInventarioFil")
    If objRelInventarioFil Is Nothing Then

        Set objRelInventarioFil = New ClassRelInventarioFil
        lErro = GuardaObj("RelInventarioFil", objRelInventarioFil)
        If lErro <> SUCESSO Then gError 90369

    End If

    ObterSaldoProdDtFilDeTerc = objRelInventarioFil.ObterSaldoProdDtFilDeTerc(dQuantidade, sProduto, dtData, iFilialEmpresa)

    Exit Function

Erro_ObterSaldoProdDtFilDeTerc:

    ObterSaldoProdDtFilDeTerc = gErr

    Select Case gErr

        Case 90369

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152963)

    End Select

    Exit Function

End Function

Public Function ObterSaldoProdDtFilEmTerc(dQuantidade As Double, ByVal sProduto As String, ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque em 3os de um produto em uma data

Dim lErro As Long, objRelInventarioFil As ClassRelInventarioFil

On Error GoTo Erro_ObterSaldoProdDtFilEmTerc

    Set objRelInventarioFil = ObtemObj("RelInventarioFil")
    If objRelInventarioFil Is Nothing Then

        Set objRelInventarioFil = New ClassRelInventarioFil
        lErro = GuardaObj("RelInventarioFil", objRelInventarioFil)
        If lErro <> SUCESSO Then gError 90370

    End If

    ObterSaldoProdDtFilEmTerc = objRelInventarioFil.ObterSaldoProdDtFilEmTerc(dQuantidade, sProduto, dtData, iFilialEmpresa)

    Exit Function

Erro_ObterSaldoProdDtFilEmTerc:

    ObterSaldoProdDtFilEmTerc = gErr

    Select Case gErr

        Case 90370

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152964)

    End Select

    Exit Function

End Function

'JONES - Completar as funções abaixo para RCPE Modelo 3
'Maristela(inicio)
Public Function ObterSerieNF(sSerieNF As String, ByVal lNumIntDoc As Long) As Long
'Para obter a Série da Nota Fiscal

Dim lErro As Long, objRelItemNFiscal As ClassRelItemNFiscal
Dim sAux As String

On Error GoTo Erro_ObterSerieNF

    Set objRelItemNFiscal = ObtemObj("RelItemNFiscal")
    If objRelItemNFiscal Is Nothing Then

        Set objRelItemNFiscal = New ClassRelItemNFiscal
        lErro = GuardaObj("RelItemNFiscal", objRelItemNFiscal)
        If lErro <> SUCESSO Then gError 90458

    End If
      
    ObterSerieNF = objRelItemNFiscal.ObterSerieNF(lNumIntDoc, sAux)

    Call Copia_String(sSerieNF, sAux)

    Exit Function

Erro_ObterSerieNF:

    ObterSerieNF = gErr

    Select Case gErr

        Case 90458

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152965)

    End Select

    Exit Function

End Function

Public Function ObterNumeroNF(dNumeroNF As Double, ByVal lNumIntDoc As Long) As Long
'Para obter o Número da Nota Fiscal

Dim lErro As Long, objRelItemNFiscal As ClassRelItemNFiscal, lNumeroNF As Long

On Error GoTo Erro_ObterNumeroNF

    Set objRelItemNFiscal = ObtemObj("RelItemNFiscal")
    If objRelItemNFiscal Is Nothing Then

        Set objRelItemNFiscal = New ClassRelItemNFiscal
        lErro = GuardaObj("RelItemNFiscal", objRelItemNFiscal)
        If lErro <> SUCESSO Then gError 90459

    End If

    ObterNumeroNF = objRelItemNFiscal.ObterNumeroNF(lNumIntDoc, lNumeroNF)
    dNumeroNF = lNumeroNF

    Exit Function

Erro_ObterNumeroNF:

    ObterNumeroNF = gErr

    Select Case gErr

        Case 90459

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152966)

    End Select

    Exit Function

End Function


Public Function ObterDataNF(dtDataNF As Date, ByVal lNumIntDoc As Long) As Long
'Para obter a Data de Emissão da Nota Fiscal

Dim lErro As Long, objRelItemNFiscal As ClassRelItemNFiscal

On Error GoTo Erro_ObterDataNF

    Set objRelItemNFiscal = ObtemObj("RelItemNFiscal")
    If objRelItemNFiscal Is Nothing Then

        Set objRelItemNFiscal = New ClassRelItemNFiscal
        lErro = GuardaObj("RelItemNFiscal", objRelItemNFiscal)
        If lErro <> SUCESSO Then gError 90460

    End If

    ObterDataNF = objRelItemNFiscal.ObterDataNF(lNumIntDoc, dtDataNF)

    Exit Function

Erro_ObterDataNF:

    ObterDataNF = gErr

    Select Case gErr

        Case 90460

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152967)

    End Select

    Exit Function

End Function

Public Function ObterNaturezaNF(sNaturezaNF As String, ByVal lNumIntDoc As Long) As Long
'Para obter a Natureza de Operação da Nota Fiscal

Dim lErro As Long, objRelItemNFiscal As ClassRelItemNFiscal
Dim sAux As String

On Error GoTo Erro_ObterNaturezaNF

    Set objRelItemNFiscal = ObtemObj("RelItemNFiscal")
    If objRelItemNFiscal Is Nothing Then

        Set objRelItemNFiscal = New ClassRelItemNFiscal
        lErro = GuardaObj("RelItemNFiscal", objRelItemNFiscal)
        If lErro <> SUCESSO Then gError 90461

    End If

    ObterNaturezaNF = objRelItemNFiscal.ObterNaturezaNF(lNumIntDoc, sAux)

    Call Copia_String(sNaturezaNF, sAux)

    Exit Function

Erro_ObterNaturezaNF:

    ObterNaturezaNF = gErr

    Select Case gErr

        Case 90461

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152968)

    End Select

    Exit Function

End Function
'Maristela(inicio)
'Término - JONES - Completar as funções abaixo para RCPE Modelo 3

'Maristela(Inicio)
Function EsvaziaProduto_INPAL(sStringOut As String, ByVal sStringIn As String) As Long
'Retira zero`s entre entre letra e número.Ex: A000010 = A10

Dim iPOS As Integer
Dim iCont  As Integer
Dim sCaracter As String
Dim sString As String
Dim sStringOutDepois As String
Dim sStringOutAntes As String
Dim bPassou As Boolean
Dim iPosNum As Integer

    iPOS = Len(sStringIn)
    sString = sStringIn
    
    sCaracter = left(sString, 1)
    Do While Not IsNumeric(sCaracter)
        sStringOutAntes = sStringOutAntes & sCaracter
        iCont = iCont + 1
        sString = right(sStringIn, iPOS - iCont)
        sCaracter = left(sString, 1)
    Loop
    
    iPosNum = iPOS - Len(sStringOutAntes)
    
    For iCont = 1 To iPosNum
        sCaracter = left(sString, 1)
        sString = right(sString, iPosNum - iCont)
        If sCaracter <> "0" Or bPassou = True Then
            sStringOutDepois = sStringOutDepois & sCaracter
            bPassou = True
        End If
    Next
    
    sStringOut = UCase(sStringOutAntes) & sStringOutDepois
    
    EsvaziaProduto_INPAL = SUCESSO

End Function
'Maristela(Fim)

Function Lcto_GeraHistoricoAux(sHistorico As String, ByVal lNumIntDoc As Long, ByVal iTransacao As Integer) As Long
'retorna em sHistorico um historico automatico a partir do numintdoc e transacao do lcto contabil informados

Dim lErro As Long, objRelLctoHistAux As ClassRelLctoHistAux
Dim sAux As String

On Error GoTo Erro_Lcto_GeraHistoricoAux

    Set objRelLctoHistAux = ObtemObj("RelLctoHistAux")
    If objRelLctoHistAux Is Nothing Then

        Set objRelLctoHistAux = New ClassRelLctoHistAux
        lErro = GuardaObj("RelLctoHistAux", objRelLctoHistAux)
        If lErro <> SUCESSO Then gError 93631

    End If

    Lcto_GeraHistoricoAux = objRelLctoHistAux.Lcto_GeraHistoricoAux(sAux, lNumIntDoc, iTransacao)

    Call Copia_String(sHistorico, sAux)

    Exit Function

Erro_Lcto_GeraHistoricoAux:

    Lcto_GeraHistoricoAux = gErr

    Select Case gErr

        Case 93631

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152969)

    End Select

    Exit Function

End Function

Function ArredondaNum(dNumeroArredondado As Double, ByVal dNumero As Double, ByVal iDecimais As Integer) As Long
'retorna um numero arredondado p/iDecimais

Dim lErro As Long

On Error GoTo Erro_ArredondaNum

    'estou somando 0.00000000001 p/evitar problemas de "Banker's rounding" que fazem com que round(111.245,2) seja 111.24

    If dNumero >= 0 Then
        dNumeroArredondado = Round(dNumero + 0.00000000001, iDecimais)
    Else
        dNumeroArredondado = Round(dNumero - 0.00000000001, iDecimais)
    End If
    
    ArredondaNum = SUCESSO
     
    Exit Function
    
Erro_ArredondaNum:

    ArredondaNum = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152970)
     
    End Select
     
    Exit Function

End Function

Public Function ObterQuantInicialTercEsc(dQuantInicial As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'obtem  a qtde inicial de 3os de um produto/almoxarifado em um escaninho antes da data inicial
'dQuantInicial é a quantidade inicial do produto sProduto no Almoxarifado iAlmoxarifado antes da data inicial dtDataInicial, que é retornada
'Caso dQuantInicial = -1 significa que a quantidade não está disponível.
'sProduto - é o codigo do produto em questao
'iAlmoxarifado é o codigo do almoxarifado em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstAlm3 As ClassRelMesEstAlm3

On Error GoTo Erro_ObterQuantInicialTercEsc

    Set objRelMesEstAlm3 = ObtemObj("RelMesEstAlm3")
    If objRelMesEstAlm3 Is Nothing Then

        Set objRelMesEstAlm3 = New ClassRelMesEstAlm3
        lErro = GuardaObj("RelMesEstAlm3", objRelMesEstAlm3)
        If lErro <> SUCESSO Then Error 55078

    End If

    ObterQuantInicialTercEsc = objRelMesEstAlm3.ObterQuantInicialTercEsc(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dQuantInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)

    Exit Function

Erro_ObterQuantInicialTercEsc:

    ObterQuantInicialTercEsc = Err

    Select Case Err

        Case 55078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152971)

    End Select

    Exit Function

End Function

Public Function ObterValorInicialTercEsc(dValorInicial As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'obtem o saldo(valor) inicial de ou em 3os de um produto/almoxarifado num escaninho antes da data inicial
'dValorInicial é o saldo (valor) inicial do produto sProduto no Almoxarifado iAlmoxarifado antes da data inicial dtDataInicial, que é retornado
'Caso dValorInicial = -1 significa que o valor não está disponível.
'sProduto - é o codigo do produto em questao
'iAlmoxarifado é o codigo do almoxarifado em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstAlm3 As ClassRelMesEstAlm3

On Error GoTo Erro_ObterValorInicialTercEsc

    Set objRelMesEstAlm3 = ObtemObj("RelMesEstAlm3")
    If objRelMesEstAlm3 Is Nothing Then

        Set objRelMesEstAlm3 = New ClassRelMesEstAlm3
        lErro = GuardaObj("RelMesEstAlm3", objRelMesEstAlm3)
        If lErro <> SUCESSO Then gError 93679

    End If

    ObterValorInicialTercEsc = objRelMesEstAlm3.ObterValorInicialTercEsc(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dValorInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)

    Exit Function

Erro_ObterValorInicialTercEsc:

    ObterValorInicialTercEsc = gErr

    Select Case gErr

        Case 93679

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152972)

    End Select

    Exit Function

End Function

Public Function ObterViaTranspDesc(sViaTransp As String, iViaTransp As Integer) As Long

Dim lErro As Long
Dim objLayoutDuplicata As ClassLayoutDuplicata
Dim sAux As String

On Error GoTo Erro_ObterViaTranspDesc

    Set objLayoutDuplicata = ObtemObj("LayoutDuplicata")
    If objLayoutDuplicata Is Nothing Then

        Set objLayoutDuplicata = New ClassLayoutDuplicata
        lErro = GuardaObj("LayoutDuplicata", objLayoutDuplicata)
        If lErro <> SUCESSO Then gError 94571

    End If

    ObterViaTranspDesc = objLayoutDuplicata.Obtem_ViaTransp_Descricao_Rel(sAux, iViaTransp)

    Call Copia_String(sViaTransp, sAux)

    Exit Function

Erro_ObterViaTranspDesc:

    ObterViaTranspDesc = gErr

    Select Case gErr

        Case 94571

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152973)

    End Select

    Exit Function

End Function

Public Function AcumulaValor(dValAcum As Double, ByVal sChave As String, ByVal dVal As Double, ByVal iRound As Integer) As Long

    If gsUltChvAcum <> sChave And sChave <> "" Then
    
        If iRound >= 0 Then
        
            gValAcum = Round(gValAcum + dVal, iRound)
            
        Else
        
            gValAcum = gValAcum + dVal
            
        End If
            
        gsUltChvAcum = sChave
        
    End If
    
    dValAcum = gValAcum
    
    AcumulaValor = SUCESSO
    
End Function

Public Function AcumulaValorChv(dValAcum As Double, ByVal sChave As String, ByVal iIndice As Integer, ByVal dVal As Double, ByVal iRound As Integer) As Long

Dim vValor As Variant
Dim bAchou As Boolean

    If iIndice <= 64 Then

        If gcolChv(iIndice) Is Nothing Then Set gcolChv(iIndice) = New Collection
    
        If Len(Trim(sChave)) > 0 Then
            bAchou = False
            For Each vValor In gcolChv(iIndice)
                If sChave = vValor Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                If iRound >= 0 Then
                    gValAcumChv(iIndice) = Round(gValAcumChv(iIndice) + dVal, iRound)
                Else
                    gValAcumChv(iIndice) = gValAcumChv(iIndice) + dVal
                End If
                gcolChv(iIndice).Add sChave
            End If
        End If
        
        dValAcum = gValAcumChv(iIndice)
        
    End If
    
    AcumulaValorChv = SUCESSO
    
End Function

Public Function SomaValorChv(dValAcum As Double, ByVal sChave As String, ByVal dVal As Double, ByVal iRound As Integer) As Long

Dim vValor As Variant
Dim bAchou As Boolean

    If gsUltChvSom <> sChave Then
        gValSom = 0
        gsUltChvSom = sChave
    End If
    
    If iRound >= 0 Then
        gValSom = Round(gValSom + dVal, iRound)
    Else
        gValSom = gValSom + dVal
    End If
    
    dValAcum = gValSom
    
    SomaValorChv = SUCESSO
    
End Function

'07/11/01 - Marcelo inicio
Public Function ObterQuantInicialFil(dQuantInicial As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'obtem  a qtde inicial de um produto/filial antes da data inicial
'dQuantInicial é a quantidade inicial do produto sProduto na Filial iFilialEmpresa antes da data inicial dtDataInicial, que é retornada
'Caso dQuantInicial = -1 significa que a quantidade não está disponível.
'sProduto - é o codigo do produto em questao
'iFilialEmpresa é o codigo da filial em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstFil As ClassRelMesEstFil

On Error GoTo Erro_ObterQuantInicialFil

    Set objRelMesEstFil = ObtemObj("RelMesEstFil")
    If objRelMesEstFil Is Nothing Then

        Set objRelMesEstFil = New ClassRelMesEstFil
        lErro = GuardaObj("RelMesEstFil", objRelMesEstFil)
        If lErro <> SUCESSO Then gError 93731

    End If

    ObterQuantInicialFil = objRelMesEstFil.ObterQuantInicial1(sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, dQuantInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)

    Exit Function

Erro_ObterQuantInicialFil:

    ObterQuantInicialFil = gErr

    Select Case gErr

        Case 93731

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152974)

    End Select

    Exit Function

End Function
Public Function ObterValorInicialFil(dValorInicial As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'obtem o saldo(valor) inicial de um produto/filial antes da data inicial
'dValorInicial é o saldo (valor) inicial do produto sProduto na Filial iFilialEmpresa antes da data inicial dtDataInicial, que é retornado
'Caso dValorInicial = -1 significa que o valor não está disponível.
'sProduto - é o codigo do produto em questao
'iFilialEmpresa é o codigo da filial em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstFil As ClassRelMesEstFil

On Error GoTo Erro_ObterValorInicialFil

    Set objRelMesEstFil = ObtemObj("RelMesEstFil")
    If objRelMesEstFil Is Nothing Then

        Set objRelMesEstFil = New ClassRelMesEstFil
        lErro = GuardaObj("RelMesEstFil", objRelMesEstFil)
        If lErro <> SUCESSO Then gError 93732

    End If

    ObterValorInicialFil = objRelMesEstFil.ObterValorInicial1(sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, dValorInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)

    Exit Function

Erro_ObterValorInicialFil:

    ObterValorInicialFil = gErr

    Select Case gErr

        Case 93732

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152975)

    End Select

    Exit Function

End Function
'13/11/01 - Marcelo inicio
Public Function ObterQuantInicialFilTercEsc(dQuantInicial As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'obtem  a qtde inicial de ou em 3os de um produto/filial num escaninho antes da data inicial
'dQuantInicial é a quantidade inicial do produto sProduto na Filial iFilialEmpresa antes da data inicial dtDataInicial, que é retornada
'Caso dQuantInicial = -1 significa que a quantidade não está disponível.
'sProduto - é o codigo do produto em questao
'iFilialEmpresa é o codigo da filial em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstFil3 As ClassRelMesEstFil3

On Error GoTo Erro_ObterQuantInicialFilTercEsc

    Set objRelMesEstFil3 = ObtemObj("RelMesEstFil3")
    If objRelMesEstFil3 Is Nothing Then

        Set objRelMesEstFil3 = New ClassRelMesEstFil3
        lErro = GuardaObj("RelMesEstFil3", objRelMesEstFil3)
        If lErro <> SUCESSO Then gError 93742

    End If

    ObterQuantInicialFilTercEsc = objRelMesEstFil3.ObterQuantInicial1(sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, dQuantInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)

    Exit Function

Erro_ObterQuantInicialFilTercEsc:

    ObterQuantInicialFilTercEsc = gErr

    Select Case gErr

        Case 93742

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152976)

    End Select

    Exit Function

End Function

Public Function ObterValorInicialFilTercEsc(dValorInicial As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'obtem o saldo(valor) inicial de ou em 3os de um produto/filial num escaninho antes da data inicial
'dValorInicial é o saldo (valor) inicial do produto sProduto na Filial iFilialEmpresa antes da data inicial dtDataInicial, que é retornado
'Caso dValorInicial = -1 significa que o valor não está disponível.
'sProduto - é o codigo do produto em questao
'iFilialEmpresa é o codigo da filial em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula
'dtDataFinal é a data final do periodo que está sendo consultado ou data nula
'iApropriacaoCusto é o tipo de apropriacao associado ao produto
'iAnoEstApurado é o ultimo ano em que houve apuração de custo de produção
'iMesEstApurado é o ultimo mes em que houve apuração de custo de produção

Dim lErro As Long, objRelMesEstFil3 As ClassRelMesEstFil3

On Error GoTo Erro_ObterValorInicialFilTercEsc

    Set objRelMesEstFil3 = ObtemObj("RelMesEstFil3")
    If objRelMesEstFil3 Is Nothing Then

        Set objRelMesEstFil3 = New ClassRelMesEstFil3
        lErro = GuardaObj("RelMesEstFil3", objRelMesEstFil3)
        If lErro <> SUCESSO Then gError 93741

    End If

    ObterValorInicialFilTercEsc = objRelMesEstFil3.ObterValorInicial1(sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, dValorInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)

    Exit Function

Erro_ObterValorInicialFilTercEsc:

    ObterValorInicialFilTercEsc = gErr

    Select Case gErr

        Case 93741

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152977)

    End Select

    Exit Function

End Function
' 13/11/01 Marcelo fim

'27/11/01 Marcelo inicio

Public Function ObterDifProducao(dDIF As Double, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dCustoMatPrim As Double) As Long
'obtem o valor de despesas indiretas correspondentes à uma producao entrada
'iFilialEmpresa é o codigo da filial em questão.
'dtData é a data da produção entrada
'dCustoMatPrim é o custo da materia prima da producao

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_ObterDifProducao

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then gError 93752

    End If

    ObterDifProducao = objRelCustoProd.ObterDifProducao(iFilialEmpresa, dtData, dCustoMatPrim, dDIF)

    Exit Function

Erro_ObterDifProducao:

    ObterDifProducao = gErr

    Select Case gErr

        Case 93752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152978)

    End Select

    Exit Function

End Function

Public Function ObterDdfProducao(dDDF As Double, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal lHorasMaquina As Long) As Long
'obtem o valor de despesas diretas correspondentes à uma producao entrada
'iFilialEmpresa é o codigo da filial em questão.
'dtData é a data da produção entrada
'lHorasMaquina é a quantidade em minutos gasto para produzir os produtos

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_ObterDdfProducao

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then gError 93755

    End If

    ObterDdfProducao = objRelCustoProd.ObterDdfProducao(iFilialEmpresa, dtData, lHorasMaquina, dDDF)

    Exit Function

Erro_ObterDdfProducao:

    ObterDdfProducao = gErr

    Select Case gErr

        Case 93755

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152979)

    End Select

    Exit Function

End Function


' 27/11/01 Marcelo fim

Public Function RelCtaSaldo_Exclui(dNumIntRel2 As Double, ByVal dNumIntRel As Double) As Long
'Limpa os registros temporarios montados (a partir de NumIntRel) para execucao do Relatorio

Dim lComando As Long, lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long, lNumIntRel As Long

On Error GoTo Erro_RelCtaSaldo_Exclui
    
    lNumIntRel = dNumIntRel
    
    'Abre o Comando
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 95380
    
    lComando2 = Comando_AbrirExt(GL_lConexaoRel)
    If lComando2 = 0 Then gError 95380
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirExt(GL_lConexaoRel)
    If lTransacao = 0 Then gError 95381
    
    'Deleta todos os registros temporarios referentes ao relatorio com NumIntRel = lNumIntRel
    lErro = Comando_Executar(lComando, "DELETE FROM RelCtaSaldo WHERE NumIntRel = ?", lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 95383
    
    'Deleta todos os registros temporarios referentes ao relatorio com NumIntRel = lNumIntRel
    lErro = Comando_Executar(lComando2, "DELETE FROM RelLcto WHERE NumIntRel = ?", lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 95383
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 95382
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    'Retorna SUCESSO
    RelCtaSaldo_Exclui = SUCESSO
    
    Exit Function

Erro_RelCtaSaldo_Exclui:

    'Retorna o Erro
    RelCtaSaldo_Exclui = gErr
    
    'Mostra msg de Erro
    Select Case gErr
    
        Case 95380
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 95381
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 95382
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 95383
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELCTASALDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152980)
     
    End Select
    
    'Volta com a transacao
    If lTransacao <> 0 Then Call Transacao_RollbackExt(lTransacao)
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Function RelArqTemp_Exclui(dNumIntRel2 As Double, ByVal dNumIntRel As Double, ByVal sNomeTabela As String) As Long
'Limpa os registros temporarios montados (a partir de NumIntRel) para execucao do Relatorio

Dim lComando As Long
Dim lTransacao As Long
Dim lErro As Long, lNumIntRel As Long

On Error GoTo Erro_RelArqTemp_Exclui
    
    lNumIntRel = dNumIntRel
    
    'Abre o Comando
    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 95380
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirExt(GL_lConexaoRel)
    If lTransacao = 0 Then gError 95381
    
    'Deleta todos os registros temporarios referentes ao relatorio com NumIntRel = lNumIntRel
    lErro = Comando_Executar(lComando, "DELETE FROM " & sNomeTabela & " WHERE NumIntRel = ?", lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 95383
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 95382
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    
    'Retorna SUCESSO
    RelArqTemp_Exclui = SUCESSO
    
    Exit Function

Erro_RelArqTemp_Exclui:

    'Retorna o Erro
    RelArqTemp_Exclui = gErr
    
    'Mostra msg de Erro
    Select Case gErr
    
        Case 95380
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 95381
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 95382
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 95383
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELARQTEMP", gErr, sNomeTabela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152981)
     
    End Select
    
    'Volta com a transacao
    If lTransacao <> 0 Then Call Transacao_RollbackExt(lTransacao)
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'??? ERRO_EXCLUSAO_RELARQTEMP passar a tabela como parametro

Public Function ObterPaginasNFInpal(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal iMaxLinhasPagina As Integer, ByVal iMaxCaracLinha As Integer) As Long
'Para obter o Número de Páginas para a impressao de uma nota fiscal identificada por lNumIntDoc.
'Leva em conta o maximo de linhas por pagina e o numero maximo de caracteres p/impressao de descricao de item da nf, p/linha

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim iNumPaginas As Integer

On Error GoTo Erro_ObterPaginasNFInpal

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 69091

    End If

    ObterPaginasNFInpal = objRelNF.ObterPaginasNFInpal(lNumIntDoc, iNumPaginas, iMaxLinhasPagina, iMaxCaracLinha)

    dNumPagina = iNumPaginas

    Exit Function

Erro_ObterPaginasNFInpal:

    ObterPaginasNFInpal = gErr

    Select Case gErr

        Case 69091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152982)

    End Select

    Exit Function

End Function

Public Function ObterPagNF(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal lNumPagAtualEmissao As Long) As Long
'obtem o numero da pagina corrente da nota fiscal (1a pagina da nota, 2a,...)
'lNumPagAtualEmissao é o numero da pagina na emissao de notas sendo executada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim iNumPag As Integer

On Error GoTo Erro_ObterPagNF

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 81810

    End If

    ObterPagNF = objRelNF.ObterPagNF(lNumIntDoc, iNumPag, lNumPagAtualEmissao)

    dNumPagina = iNumPag

    Exit Function

Erro_ObterPagNF:

    ObterPagNF = gErr

    Select Case gErr

        Case 81810

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152983)

    End Select

    Exit Function

End Function

Public Function Fatura_ObterNFValor(dValorNF As Double, ByVal lNumIntDoc As Long, ByVal iNumNF As Integer) As Long
'retorna o valor da enesima nota fiscal cobrada pela fatura identificada por lNumIntDoc da tabela de titulosrec (n = iNumNF)
'a posicao das notas fiscais dentro da fatura é dada pela ordem crescente de numeracao

Dim lErro As Long
Dim objRelFatura As ClassRelFatura

On Error GoTo Erro_Fatura_ObterNFValor

    Set objRelFatura = ObtemObj("RelFatura")

    If objRelFatura Is Nothing Then

        Set objRelFatura = New ClassRelFatura
        lErro = GuardaObj("RelFatura", objRelFatura)
        If lErro <> SUCESSO Then gError 81823

    End If

    Fatura_ObterNFValor = objRelFatura.Fatura_ObterNFValor(dValorNF, lNumIntDoc, iNumNF)
 
    Exit Function
    
Erro_Fatura_ObterNFValor:

    Fatura_ObterNFValor = gErr
     
    Select Case gErr
          
        Case 81823
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152984)
     
    End Select
     
    Exit Function

End Function

Public Function Fatura_ObterNFNumero(dNumeroNF As Double, ByVal lNumIntDoc As Long, ByVal iNumNF As Integer) As Long
'retorna o valor da enesima nota fiscal cobrada pela fatura identificada por lNumIntDoc da tabela de titulosrec (n = iNumNF)
'a posicao das notas fiscais dentro da fatura é dada pela ordem crescente de numeracao

Dim lErro As Long, lNumNotaFiscal As Long
Dim objRelFatura As ClassRelFatura

On Error GoTo Erro_Fatura_ObterNFNumero

    Set objRelFatura = ObtemObj("RelFatura")

    If objRelFatura Is Nothing Then

        Set objRelFatura = New ClassRelFatura
        lErro = GuardaObj("RelFatura", objRelFatura)
        If lErro <> SUCESSO Then gError 81824

    End If

    Fatura_ObterNFNumero = objRelFatura.Fatura_ObterNFNumero(lNumNotaFiscal, lNumIntDoc, iNumNF)
 
    dNumeroNF = lNumNotaFiscal
    
    Exit Function
    
Erro_Fatura_ObterNFNumero:

    Fatura_ObterNFNumero = gErr
     
    Select Case gErr
          
        Case 81824
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152985)
     
    End Select
     
    Exit Function

End Function

Public Function ObterProduto_IndicacaoGastoInformado(dProdutoGastoInformado As Double, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAno As Integer, ByVal iMes As Integer) As Long
'Para obter a indicação se se trata de um produto cujo gasto foi informado ou não. iProdutoGastoInformato=1 ==> Gasto Informado.

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd, iProdutoGastoInformado As Integer

On Error GoTo Erro_ObterProduto_IndicacaoGastoInformado

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then gError 92923

    End If

    ObterProduto_IndicacaoGastoInformado = objRelCustoProd.ObterProduto_IndicacaoGastoInformado(iProdutoGastoInformado, iFilialEmpresa, sProduto, iAno, iMes)
    dProdutoGastoInformado = iProdutoGastoInformado

    Exit Function

Erro_ObterProduto_IndicacaoGastoInformado:

    ObterProduto_IndicacaoGastoInformado = gErr

    Select Case gErr

        Case 92923

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152986)

    End Select

    Exit Function

End Function

Public Function ObterProduto_ValorGasto(dValor As Double, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAno As Integer, ByVal iMes As Integer, ByVal dQuantidade As Double, ByVal sSiglaUM As String) As Long
'Para obter o valor gasto com o produto a partir do total informado.

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_ObterProduto_ValorGasto

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then gError 92924

    End If

    ObterProduto_ValorGasto = objRelCustoProd.ObterProduto_ValorGasto(dValor, iFilialEmpresa, sProduto, iAno, iMes, dQuantidade, sSiglaUM)

    Exit Function

Erro_ObterProduto_ValorGasto:

    ObterProduto_ValorGasto = gErr

    Select Case gErr

        Case 92924

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152987)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao(sPorto As String, ByVal lCodigo As Long) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objRelGR As Object, sAux As String

On Error GoTo Erro_ObterCodOperacao

    Set objRelGR = ObtemObj("RelGR")

    If objRelGR Is Nothing Then

        Set objRelGR = CreateObject("ADRELVBGR.ClassRelGR")
        lErro = GuardaObj("RelGR", objRelGR)
        If lErro <> SUCESSO Then gError 99297

    End If

    ObterCodOperacao = objRelGR.ObterCodOperacao(lCodigo, sAux)
    Call Copia_String(sPorto, sAux)
    
    Exit Function

Erro_ObterCodOperacao:

    ObterCodOperacao = gErr

    Select Case gErr

        Case 99297

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152988)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao1(sAlfandega As String, ByVal lCodigo As Long) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objRelGR As Object, sAux As String

On Error GoTo Erro_ObterCodOperacao1

    Set objRelGR = ObtemObj("RelGR")

    If objRelGR Is Nothing Then

        Set objRelGR = CreateObject("ADRELVBGR.ClassRelGR")
        lErro = GuardaObj("RelGR", objRelGR)
        If lErro <> SUCESSO Then gError 99300

    End If

    ObterCodOperacao1 = objRelGR.ObterCodOperacao1(lCodigo, sAux)
    Call Copia_String(sAlfandega, sAux)

    Exit Function

Erro_ObterCodOperacao1:

    ObterCodOperacao1 = gErr

    Select Case gErr

        Case 99300

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152989)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao2(sMotorista As String, ByVal lCodigo As Long) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objRelGR As Object, sAux As String

On Error GoTo Erro_ObterCodOperacao2

    Set objRelGR = ObtemObj("RelGR")

    If objRelGR Is Nothing Then

        Set objRelGR = CreateObject("ADRELVBGR.ClassRelGR")
        lErro = GuardaObj("RelGR", objRelGR)
        If lErro <> SUCESSO Then gError 99301

    End If

    ObterCodOperacao2 = objRelGR.ObterCodOperacao2(lCodigo, sAux)
    Call Copia_String(sMotorista, sAux)

    Exit Function

Erro_ObterCodOperacao2:

    ObterCodOperacao2 = gErr

    Select Case gErr

        Case 99301

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152990)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao3(sPlacaCaminhao As String, ByVal lCodigo As Long) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objRelGR As Object, sAux As String

On Error GoTo Erro_ObterCodOperacao3

    Set objRelGR = ObtemObj("RelGR")

    If objRelGR Is Nothing Then

        Set objRelGR = CreateObject("ADRELVBGR.ClassRelGR")
        lErro = GuardaObj("RelGR", objRelGR)
        If lErro <> SUCESSO Then gError 99302

    End If

    ObterCodOperacao3 = objRelGR.ObterCodOperacao3(lCodigo, sAux)
    Call Copia_String(sPlacaCaminhao, sAux)

    Exit Function

Erro_ObterCodOperacao3:

    ObterCodOperacao3 = gErr

    Select Case gErr

        Case 99302

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152991)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao4(sPlacaCarreta As String, ByVal lCodigo As Long) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objRelGR As Object, sAux As String

On Error GoTo Erro_ObterCodOperacao4

    Set objRelGR = ObtemObj("RelGR")

    If objRelGR Is Nothing Then

        Set objRelGR = CreateObject("ADRELVBGR.ClassRelGR")
        lErro = GuardaObj("RelGR", objRelGR)
        If lErro <> SUCESSO Then gError 99303

    End If

    ObterCodOperacao4 = objRelGR.ObterCodOperacao4(lCodigo, sAux)
    Call Copia_String(sPlacaCarreta, sAux)

    Exit Function

Erro_ObterCodOperacao4:

    ObterCodOperacao4 = gErr

    Select Case gErr

        Case 99303

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152992)

    End Select

    Exit Function

End Function

Sub Copia_String(sDestino As String, ByVal sOrigem As String)
Dim iIndice As Integer

    sDestino = ""
    
    For iIndice = 1 To Len(sOrigem)

        sDestino = sDestino & Mid(sOrigem, iIndice, 1)
        
    Next
    
End Sub

Public Function AcumulaValorInd(dValAcum As Double, ByVal sChave As String, ByVal dVal As Double, ByVal iRound As Integer, ByVal iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_AcumulaValorInd

    If gasUltChvAcum(iIndice) <> sChave And sChave <> "" Then
    
        If baZerarAcum(iIndice) Then
        
            gaValAcum(iIndice) = 0
            baZerarAcum(iIndice) = False
        
        End If
        
        If iRound >= 0 Then
        
            gaValAcum(iIndice) = Round(gaValAcum(iIndice) + dVal, iRound)
            
        Else
        
            gaValAcum(iIndice) = gaValAcum(iIndice) + dVal
            
        End If
            
        gasUltChvAcum(iIndice) = sChave
        
    End If
    
    dValAcum = gaValAcum(iIndice)
    
    If sChave = "" Then
        
        baZerarAcum(iIndice) = True
        
    End If
    
    AcumulaValorInd = SUCESSO
    
    Exit Function
     
Erro_AcumulaValorInd:

    AcumulaValorInd = gErr
    
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152993)
     
    End Select
     
    Exit Function

End Function

Public Function Calcula_Participacao_Componente_Kit(sParc As String, ByVal sProduto As String, ByVal sVersao As String, ByVal sComponente As String) As Long

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_Calcula_Participacao_Componente_Kit

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then Error 43323

    End If

    '??? interface ruim, deveria retornar lerro
    Call Copia_String(sParc, objRelCustoProd.Calcula_Participacao_Componente_Kit(sProduto, sVersao, sComponente))

    Calcula_Participacao_Componente_Kit = SUCESSO

    Exit Function

Erro_Calcula_Participacao_Componente_Kit:

    Calcula_Participacao_Componente_Kit = Err

    Select Case Err

        Case 43323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152994)

    End Select
    
End Function

Public Function Calcula_Rendimento_ProdutoVersao(sRend As String, ByVal sProduto As String, ByVal sVersao As String) As Long

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_Calcula_Rendimento_ProdutoVersao

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then Error 43323

    End If
    
    '??? interface ruim, deveria retornar lerro
    Call Copia_String(sRend, objRelCustoProd.Calcula_Rendimento_ProdutoVersao(sProduto, sVersao))
    
    Calcula_Rendimento_ProdutoVersao = SUCESSO

    Exit Function

Erro_Calcula_Rendimento_ProdutoVersao:

    Calcula_Rendimento_ProdutoVersao = Err

    Select Case Err

        Case 43323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152995)

    End Select
    
End Function

Public Function Obter_Adicao_MateriaPrima(sRetorno As String, ByVal sProduto As String, ByVal sCodOP As String) As Long

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_Obter_Adicao_MateriaPrima

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then Error 43323

    End If

    Call Copia_String(sRetorno, objRelCustoProd.Obter_AdicaoMP(sProduto, sCodOP))
    
    Obter_Adicao_MateriaPrima = SUCESSO

    Exit Function

Erro_Obter_Adicao_MateriaPrima:

    Obter_Adicao_MateriaPrima = Err

    Select Case Err

        Case 43323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152996)

    End Select
    
End Function

Public Function Obter_Adicao_QuantReal(dRetorno As Double, ByVal sProduto As String, ByVal sCodOPBaixada As String) As Long

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_Obter_Adicao_QuantReal

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then Error 43323

    End If

    dRetorno = objRelCustoProd.Obter_QuantidadeReal(sProduto, sCodOPBaixada)
    
    Obter_Adicao_QuantReal = SUCESSO

    Exit Function

Erro_Obter_Adicao_QuantReal:

    Obter_Adicao_QuantReal = Err

    Select Case Err

        Case 43323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152997)

    End Select
    
End Function

Public Function Excluir_PrevVendaPrevConsumo(ByVal lNumIntDoc As Long) As Long

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_Excluir_PrevVendaPrevConsumo

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then Error 43323

    End If

    Excluir_PrevVendaPrevConsumo = objRelCustoProd.Excluir_PrevVendaPrevConsumo_Aux(lNumIntDoc)
    
    Exit Function

Erro_Excluir_PrevVendaPrevConsumo:

    Excluir_PrevVendaPrevConsumo = Err

    Select Case Err

        Case 43323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152998)

    End Select
    
End Function

Public Function ObterNumLoteItensNFiscal(sNumLoteRetorno As String, ByVal lNumIntItemNF As Long) As Long
'Função para obter o numero do Lote atravez do Numero Interno da Nota Fiscal

Dim lErro As Long
Dim objNumLoteINF As ClassRelNF
Dim sAux As String

On Error GoTo Erro_ObterNumLoteItensNFiscal

    Set objNumLoteINF = ObtemObj("RelNF")
    If objNumLoteINF Is Nothing Then

        Set objNumLoteINF = New ClassRelNF
        lErro = GuardaObj("RelNF", objNumLoteINF)
        If lErro <> SUCESSO Then gError 111300

    End If
    
    ObterNumLoteItensNFiscal = objNumLoteINF.ObterNumLoteItensNFiscal(sAux, lNumIntItemNF)

    Call Copia_String(sNumLoteRetorno, sAux)

    Exit Function

Erro_ObterNumLoteItensNFiscal:

    ObterNumLoteItensNFiscal = gErr

    Select Case gErr

        Case 111300

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152999)

    End Select

End Function

Public Function ObterSituacaoTrib1(sSituacaoTrib As String, ByVal sProduto As String, ByVal iTipoTribICMS As Integer, ByVal iTipoTribIPI As Integer) As Long
'Para obter o Código da Situação Tributária (ICMS e IPI)

Dim lErro As Long, objRelNF As ClassRelNF
Dim sAux As String

On Error GoTo Erro_ObterSituacaoTrib1

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 108764

    End If

    ObterSituacaoTrib1 = objRelNF.ObterSituacaoTrib1(sAux, sProduto, iTipoTribICMS, iTipoTribIPI)

    Call Copia_String(sSituacaoTrib, sAux)

    Exit Function

Erro_ObterSituacaoTrib1:

    ObterSituacaoTrib1 = gErr

    Select Case gErr

        Case 108764 'Tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153000)

    End Select

End Function

'Função de Relatório do Loja

Function Obtem_MvCaixaCentral_SldInicialPrimeiro(dSaldoInicial As Double, dtDataIn As Date, iCodCaixa As Integer, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim ObjRelMovCxCentral As ClassRelMovCxCentral

On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicialPrimeiro

    Set ObjRelMovCxCentral = ObtemObj("RelMovCxCentral")
    If ObjRelMovCxCentral Is Nothing Then

        Set ObjRelMovCxCentral = New ClassRelMovCxCentral
        lErro = GuardaObj("RelMovCxCentral", ObjRelMovCxCentral)
        If lErro <> SUCESSO Then gError 113166

    End If
    
    Obtem_MvCaixaCentral_SldInicialPrimeiro = ObjRelMovCxCentral.Obtem_MvCaixaCentral_SldInicial(dtDataIn, dSaldoInicial, iCodCaixa, iFilialEmpresa)

    Exit Function

Erro_Obtem_MvCaixaCentral_SldInicialPrimeiro:

    Obtem_MvCaixaCentral_SldInicialPrimeiro = gErr

    Select Case gErr

        Case 113166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153001)

    End Select

End Function

Function Obtem_MvCaixaCentral_SldIni_TMP(dSaldoInicial As Double, dtDataIn As Date, iCodCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long

Dim lErro As Long
Dim ObjRelMovCxCentral As ClassRelMovCxCentral

On Error GoTo Erro_Obtem_MvCaixaCentral_SldIni_TMP

    Set ObjRelMovCxCentral = ObtemObj("RelMovCxCentral")
    If ObjRelMovCxCentral Is Nothing Then

        Set ObjRelMovCxCentral = New ClassRelMovCxCentral
        lErro = GuardaObj("RelMovCxCentral", ObjRelMovCxCentral)
        If lErro <> SUCESSO Then gError 113166

    End If
    
    Obtem_MvCaixaCentral_SldIni_TMP = ObjRelMovCxCentral.Obtem_MvCaixaCentral_SldInicial_TMP(dtDataIn, dSaldoInicial, iCodCaixa, iFilialEmpresa, iTipoMeioPagto)

    Exit Function

Erro_Obtem_MvCaixaCentral_SldIni_TMP:

    Obtem_MvCaixaCentral_SldIni_TMP = gErr

    Select Case gErr

        Case 113166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153001)

    End Select

End Function

Function Obtem_MvCaixaCentral_SldIni_TMP_Cust(dSaldoInicial As Double, dtDataIn As Date, iCodCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long

Dim lErro As Long
Dim ObjRelMovCxCentral As ClassRelMovCxCentral

On Error GoTo Erro_Obtem_MvCaixaCentral_SldIni_TMP_Cust

    Set ObjRelMovCxCentral = ObtemObj("RelMovCxCentral")
    If ObjRelMovCxCentral Is Nothing Then

        Set ObjRelMovCxCentral = New ClassRelMovCxCentral
        lErro = GuardaObj("RelMovCxCentral", ObjRelMovCxCentral)
        If lErro <> SUCESSO Then gError 113166

    End If
    
    Obtem_MvCaixaCentral_SldIni_TMP_Cust = ObjRelMovCxCentral.Obtem_MvCaixaCentral_SldInicial_TMP_Cust(dtDataIn, dSaldoInicial, iCodCaixa, iFilialEmpresa, iTipoMeioPagto)

    Exit Function

Erro_Obtem_MvCaixaCentral_SldIni_TMP_Cust:

    Obtem_MvCaixaCentral_SldIni_TMP_Cust = gErr

    Select Case gErr

        Case 113166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153001)

    End Select

End Function

Function Obtem_MvCaixaCentral_Cred_TMP_Cust(dCred As Double, dtDataIn As Date, iCodCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long

Dim lErro As Long
Dim ObjRelMovCxCentral As ClassRelMovCxCentral

On Error GoTo Erro_Obtem_MvCaixaCentral_Cred_TMP_Cust

    Set ObjRelMovCxCentral = ObtemObj("RelMovCxCentral")
    If ObjRelMovCxCentral Is Nothing Then

        Set ObjRelMovCxCentral = New ClassRelMovCxCentral
        lErro = GuardaObj("RelMovCxCentral", ObjRelMovCxCentral)
        If lErro <> SUCESSO Then gError 113166

    End If
    
    Obtem_MvCaixaCentral_Cred_TMP_Cust = ObjRelMovCxCentral.Obtem_MvCaixaCentral_Cred_TMP_Cust(dtDataIn, dCred, iCodCaixa, iFilialEmpresa, iTipoMeioPagto)

    Exit Function

Erro_Obtem_MvCaixaCentral_Cred_TMP_Cust:

    Obtem_MvCaixaCentral_Cred_TMP_Cust = gErr

    Select Case gErr

        Case 113166

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153001)

    End Select

End Function

Function Obtem_MvCaixaCentral_Deb_TMP_Cust(dDeb As Double, dtDataIn As Date, iCodCaixa As Integer, iFilialEmpresa As Integer, iTipoMeioPagto As Integer) As Long

Dim lErro As Long
Dim ObjRelMovCxCentral As ClassRelMovCxCentral

On Error GoTo Erro_Obtem_MvCaixaCentral_Deb_TMP_Cust

    Set ObjRelMovCxCentral = ObtemObj("RelMovCxCentral")
    If ObjRelMovCxCentral Is Nothing Then

        Set ObjRelMovCxCentral = New ClassRelMovCxCentral
        lErro = GuardaObj("RelMovCxCentral", ObjRelMovCxCentral)
        If lErro <> SUCESSO Then gError 113166

    End If
    
    Obtem_MvCaixaCentral_Deb_TMP_Cust = ObjRelMovCxCentral.Obtem_MvCaixaCentral_Deb_TMP_Cust(dtDataIn, dDeb, iCodCaixa, iFilialEmpresa, iTipoMeioPagto)

    Exit Function

Erro_Obtem_MvCaixaCentral_Deb_TMP_Cust:

    Obtem_MvCaixaCentral_Deb_TMP_Cust = gErr

    Select Case gErr

        Case 113166

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153001)

    End Select

End Function


'Funções para o Relatório de Flash de Vendas
'Incluidas por sergio dia 17/12/2002

Public Function Obtem_NumVendas_IntervHoras(dNumVendas As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
'retorna em lNumeVendas o Numero de Vendas para a Data e o Intervalo de Horas especificados

Dim lErro As Long
Dim objRelFlashVendas As ClassRelFlashVendas

On Error GoTo Erro_Obtem_NumVendas_IntervHoras

    Set objRelFlashVendas = ObtemObj("RelFlashVendas")
    If objRelFlashVendas Is Nothing Then

        Set objRelFlashVendas = New ClassRelFlashVendas

        lErro = GuardaObj("RelFlashVendas", objRelFlashVendas)
        If lErro <> SUCESSO Then gError 113260

    End If

    Obtem_NumVendas_IntervHoras = objRelFlashVendas.Obtem_NumVendas_IntervHorasAux(dNumVendas, sProduto, iFilialEmpresa, dtData, dHoraIni, dHoraFim, iCaixaIni, iCaixaFim)

    Exit Function

Erro_Obtem_NumVendas_IntervHoras:

    Obtem_NumVendas_IntervHoras = gErr

    Select Case gErr

        Case 113260 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153002)

    End Select

    Exit Function

End Function



Public Function Obtem_QuantVend_IntervHoras(dQuantVendida As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
'retorna em dQuantVendida o Numero de Vendas para a Data e o Intervalo de Horas especificados

Dim lErro As Long
Dim objRelFlashVendas As ClassRelFlashVendas

On Error GoTo Erro_Obtem_QuantVend_IntervHoras

    Set objRelFlashVendas = ObtemObj("RelFlashVendas")
    If objRelFlashVendas Is Nothing Then

        Set objRelFlashVendas = New ClassRelFlashVendas

        lErro = GuardaObj("RelFlashVendas", objRelFlashVendas)
        If lErro <> SUCESSO Then gError 113287

    End If

    Obtem_QuantVend_IntervHoras = objRelFlashVendas.Obtem_QuantVend_IntervHorasAux(dQuantVendida, sProduto, iFilialEmpresa, dtData, dHoraIni, dHoraFim, iCaixaIni, iCaixaFim)

    Exit Function

Erro_Obtem_QuantVend_IntervHoras:

    Obtem_QuantVend_IntervHoras = gErr

    Select Case gErr

        Case 113287 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153003)

    End Select

    Exit Function

End Function

Public Function Obtem_VlrVend_IntervHoras(dValorVendido As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
'retorna em dQuantVendida o Numero de Vendas para a Data e o Intervalo de Horas especificados

Dim lErro As Long
Dim objRelFlashVendas As ClassRelFlashVendas

On Error GoTo Erro_Obtem_VlrVend_IntervHoras

    Set objRelFlashVendas = ObtemObj("RelFlashVendas")
    If objRelFlashVendas Is Nothing Then

        Set objRelFlashVendas = New ClassRelFlashVendas

        lErro = GuardaObj("RelFlashVendas", objRelFlashVendas)
        If lErro <> SUCESSO Then gError 113288

    End If

    Obtem_VlrVend_IntervHoras = objRelFlashVendas.Obtem_VlrVend_IntervHorasAux(dValorVendido, sProduto, iFilialEmpresa, dtData, dHoraIni, dHoraFim, iCaixaIni, iCaixaFim)

    Exit Function

Erro_Obtem_VlrVend_IntervHoras:

    Obtem_VlrVend_IntervHoras = gErr

    Select Case Err

        Case 113288 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153004)

    End Select

    Exit Function

End Function

'Funções incluidas para a Tela de Relatório RelOpFatProduto
'Incluidas por Sergio dia 24/01/03

Public Function Obtem_Quant_FatProdCx(dQuantidade As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iTipoDe As Integer, ByVal iTipoAte As Integer, ByVal iIncluiDevolucoes As Integer) As Long
'Retorna em dQuantidade a Quantidade Faturada de Produto Levando em Consideração os Filtros Passados

Dim lErro As Long
Dim objRelFatProd As ClassRelFatProd

On Error GoTo Erro_Obtem_Quant_FatProdCx

    Set objRelFatProd = ObtemObj("RelFatProd")
    If objRelFatProd Is Nothing Then

        Set objRelFatProd = New ClassRelFatProd

        lErro = GuardaObj("RelFatProd", objRelFatProd)
        If lErro <> SUCESSO Then gError 113441

    End If

    Obtem_Quant_FatProdCx = objRelFatProd.Obtem_Quant_FatProdCxAux(dQuantidade, sProduto, iFilialEmpresa, dtDataDe, dtDataAte, iTipoDe, iTipoAte, iIncluiDevolucoes)

    Exit Function

Erro_Obtem_Quant_FatProdCx:

    Obtem_Quant_FatProdCx = gErr

    Select Case gErr

        Case 113441 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153005)

    End Select

    Exit Function

End Function

Public Function Obtem_Total_FatProdCx(dValor As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iTipoDe As Integer, ByVal iTipoAte As Integer, ByVal iIncluiDevolucoes As Integer) As Long
'Retorna em dValor, o Valor do Faturmento de um Produto Levando em Consideração os Filtros Passados

Dim lErro As Long
Dim objRelFatProd As ClassRelFatProd

On Error GoTo Erro_Obtem_Total_FatProdCx

    Set objRelFatProd = ObtemObj("RelFatProd")
    If objRelFatProd Is Nothing Then

        Set objRelFatProd = New ClassRelFatProd

        lErro = GuardaObj("RelFatProd", objRelFatProd)
        If lErro <> SUCESSO Then gError 113442

    End If
    
    Obtem_Total_FatProdCx = objRelFatProd.Obtem_Total_FatProdCxAux(dValor, sProduto, iFilialEmpresa, dtDataDe, dtDataAte, iTipoDe, iTipoAte, iIncluiDevolucoes)

    Exit Function

Erro_Obtem_Total_FatProdCx:

    Obtem_Total_FatProdCx = gErr

    Select Case gErr

        Case 113442 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153006)

    End Select

    Exit Function

End Function

'' #### Fim das Funções relacionadas à Tela RelOpFatProduto ##### ''

Public Function ObterSaldoProdDtAlmox1(dQuantidade As Double, ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, ByVal dtDataFim As Date, ByVal dQuantMinima As Double) As Long
'Para obter o saldo em estoque da propria empresa de um produto em um almoxarifado em uma data

Dim lErro As Long, objRelInventario As ClassRelInventario

On Error GoTo Erro_ObterSaldoProdDtAlmox1

    Set objRelInventario = ObtemObj("RelInventario")
    If objRelInventario Is Nothing Then

        Set objRelInventario = New ClassRelInventario
        lErro = GuardaObj("RelInventario", objRelInventario)
        If lErro <> SUCESSO Then gError 105178

    End If

    ObterSaldoProdDtAlmox1 = objRelInventario.ObterSaldoProdDtAlmox1(dQuantidade, sProduto, iAlmoxarifado, dtData, iFilialEmpresa, dtDataFim, dQuantMinima)

    Exit Function

Erro_ObterSaldoProdDtAlmox1:

    ObterSaldoProdDtAlmox1 = gErr

    Select Case gErr

        Case 105178

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153007)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPedCot(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPedCot

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111893

    End If

    MapaCot_ObterPedCot = objRelMapaCot.MapaCot_ObterPedCot(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPedCot:

    MapaCot_ObterPedCot = gErr

    Select Case gErr

        Case 111893

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153008)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPedCotQtd(sRetorno As String, ByVal sProduto As String, ByVal lNumIntMapaCot As Long, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPedCotQtd

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError 111893

    End If

    MapaCot_ObterPedCotQtd = objRelMapaCot.MapaCot_ObterPedCotQtd(sAux, sProduto, lNumIntMapaCot, iSequencial, iQtdMaxColecao)

    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPedCotQtd:

    MapaCot_ObterPedCotQtd = gErr

    Select Case gErr

        Case 111893

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153009)

    End Select

    Exit Function

End Function

Public Function Obter_Adicao_PropProd(dRetorno As Double, ByVal sProduto As String, ByVal sCodOPBaixada As String) As Long

Dim lErro As Long, objRelCustoProd As ClassRelCustoProd

On Error GoTo Erro_Obter_Adicao_PropProd

    Set objRelCustoProd = ObtemObj("RelCustoProd")
    If objRelCustoProd Is Nothing Then

        Set objRelCustoProd = New ClassRelCustoProd
        lErro = GuardaObj("RelCustoProd", objRelCustoProd)
        If lErro <> SUCESSO Then Error 43323

    End If

    dRetorno = objRelCustoProd.Obter_QuantidadePropProd(sProduto, sCodOPBaixada)
    
    Obter_Adicao_PropProd = SUCESSO

    Exit Function

Erro_Obter_Adicao_PropProd:

    Obter_Adicao_PropProd = Err

    Select Case Err

        Case 43323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153010)

    End Select
    
End Function

Public Function ObterSaldoInicialInventarioTerc(dQuantidade As Double, dtDataInicial As Date, iFilialEmpresa As Integer, iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer, sProduto As String, iCodEscaninho As Integer) As Long
'para obter o saldo inicial

Dim lErro As Long, objRelInventarioTerc As ClassRelInventarioTerc

On Error GoTo Erro_ObterSaldoInicialInventarioterc

    Set objRelInventarioTerc = ObtemObj("RelKadeFilDet")
    If objRelInventarioTerc Is Nothing Then

        Set objRelInventarioTerc = New ClassRelInventarioTerc
        lErro = GuardaObj("RelKadeFilDet", objRelInventarioTerc)
        If lErro <> SUCESSO Then gError 110260

    End If

    ObterSaldoInicialInventarioTerc = objRelInventarioTerc.ObterSaldoInicial(dQuantidade, dtDataInicial, iFilialEmpresa, iTipoTerc, lCodTerc, iFilialTerc, sProduto, iCodEscaninho)

    Exit Function

Erro_ObterSaldoInicialInventarioterc:

    ObterSaldoInicialInventarioTerc = gErr

    Select Case gErr

        Case 110260

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153011)

    End Select

    Exit Function

End Function

Public Function Obter_HoraAbertura_Caixa(dHoraAbertura As Double, dtDataHoje As Date, iCaixa As Integer, iFilialEmpresa As Integer) As Long
'Devolve para o relatório a hora de abertura do caixa

Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa
Dim objRelCaixa As ClassRelCaixa

On Error GoTo Erro_Obter_HoraAbertura_Caixa

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelCaixa = ObtemObj("RelCaixa")

    'Se o obj estiver vazio
    If objRelCaixa Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelCaixa = New ClassRelCaixa

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelCaixa", objRelCaixa)
        If lErro <> SUCESSO Then gError 116231

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objMovCaixa.dtDataMovimento = dtDataHoje
    objMovCaixa.iCaixa = iCaixa
    objMovCaixa.iFilialEmpresa = iFilialEmpresa

    'Obtém a hora de abertura do caixa, de acordo com os parâmetros passados
    Obter_HoraAbertura_Caixa = objRelCaixa.Obtem_HoraAbertura_Caixa(objMovCaixa)

    'devolve para o relatório a hora encontrada
    dHoraAbertura = objMovCaixa.dHora
    
    Exit Function

Erro_Obter_HoraAbertura_Caixa:

    Obter_HoraAbertura_Caixa = gErr

    Select Case gErr

        Case 116231

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153012)

    End Select

    Exit Function

End Function

Public Function Obter_Operador_Caixa(sOperador As String, dtDataHoje As Date, iCaixa As Integer, iFilialEmpresa As Integer) As Long
'Devolve p/ o gerador de relatório o operador do caixa

Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa
Dim objRelCaixa As ClassRelCaixa

On Error GoTo Erro_Obter_Operador_Caixa

    'Carrega o obj, caso ele já tenha sido preenchido
    Set objRelCaixa = ObtemObj("RelCaixa")

    'Se o obj estiver vazio
    If objRelCaixa Is Nothing Then

        'Seta o obj como uma nova classe com os dados do Relatório
        Set objRelCaixa = New ClassRelCaixa

        'Guarda o obj em uma coleção
        lErro = GuardaObj("RelCaixa", objRelCaixa)
        If lErro <> SUCESSO Then gError 116230

    End If

    'Preenche o obj que será passado como parâmetro para a função
    objMovCaixa.dtDataMovimento = dtDataHoje
    objMovCaixa.iCaixa = iCaixa
    objMovCaixa.iFilialEmpresa = iFilialEmpresa
    
    'Obtém o operador do caixa, de acordo com os parâmetros passados
    Obter_Operador_Caixa = objRelCaixa.Obtem_Operador_Caixa(objMovCaixa, sOperador)
    
    Exit Function

Erro_Obter_Operador_Caixa:

    Obter_Operador_Caixa = gErr

    Select Case gErr

        Case 116230

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153013)

    End Select

    Exit Function

End Function

Public Function ApuracaoICMSRes_GuiaNumero(sNumero As String, ByVal iIndice As Integer, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, sAux As String
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_GuiaNumero

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_GuiaNumero = objRelResApur.GuiaNumero(sAux, iIndice, dtDataInicio, dtDataFinal, iFilialEmpresa)
    Call Copia_String(sNumero, sAux)
     
    Exit Function
    
Erro_ApuracaoICMSRes_GuiaNumero:

    ApuracaoICMSRes_GuiaNumero = gErr
     
    Select Case gErr
          
        Case 78027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153014)
     
    End Select
     
    Exit Function

End Function

Public Function ApuracaoICMSRes_GuiaData(dtData As Date, ByVal iIndice As Integer, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_GuiaData

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_GuiaData = objRelResApur.GuiaData(dtData, iIndice, dtDataInicio, dtDataFinal, iFilialEmpresa)
     
    Exit Function
    
Erro_ApuracaoICMSRes_GuiaData:

    ApuracaoICMSRes_GuiaData = gErr
     
    Select Case gErr
          
        Case 78027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153015)
     
    End Select
     
    Exit Function

End Function

Public Function ApuracaoICMSRes_GuiaValor(dValor As Double, ByVal iIndice As Integer, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_GuiaValor

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_GuiaValor = objRelResApur.GuiaValor(dValor, iIndice, dtDataInicio, dtDataFinal, iFilialEmpresa)
     
    Exit Function
    
Erro_ApuracaoICMSRes_GuiaValor:

    ApuracaoICMSRes_GuiaValor = gErr
     
    Select Case gErr
          
        Case 78027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153016)
     
    End Select
     
    Exit Function

End Function

Public Function ApuracaoICMSRes_GuiaOrgao(sOrgao As String, ByVal iIndice As Integer, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, sAux As String
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_GuiaOrgao

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_GuiaOrgao = objRelResApur.GuiaOrgao(sAux, iIndice, dtDataInicio, dtDataFinal, iFilialEmpresa)
    Call Copia_String(sOrgao, sAux)
          
    Exit Function
    
Erro_ApuracaoICMSRes_GuiaOrgao:

    ApuracaoICMSRes_GuiaOrgao = gErr
     
    Select Case gErr
          
        Case 78027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153017)
     
    End Select
     
    Exit Function

End Function

Public Function ApuracaoICMSRes_GuiaDataEnt(dtDataEnt As Date, ByVal iIndice As Integer, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_GuiaDataEnt

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_GuiaDataEnt = objRelResApur.GuiaDataEnt(dtDataEnt, iIndice, dtDataInicio, dtDataFinal, iFilialEmpresa)
     
    Exit Function
    
Erro_ApuracaoICMSRes_GuiaDataEnt:

    ApuracaoICMSRes_GuiaDataEnt = gErr
     
    Select Case gErr
          
        Case 78027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153018)
     
    End Select
     
    Exit Function

End Function

Public Function ApuracaoICMSRes_GuiaLocal(sLocal As String, ByVal iIndice As Integer, ByVal dtDataInicio As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, sAux As String
Dim objRelResApur As ClassRelResApur

On Error GoTo Erro_ApuracaoICMSRes_GuiaLocal

    Set objRelResApur = ObtemObj("objRelResApur")

    If objRelResApur Is Nothing Then

        Set objRelResApur = New ClassRelResApur
        lErro = GuardaObj("objRelResApur", objRelResApur)
        If lErro <> SUCESSO Then gError 78027

    End If

    ApuracaoICMSRes_GuiaLocal = objRelResApur.GuiaLocal(sAux, iIndice, dtDataInicio, dtDataFinal, iFilialEmpresa)
    Call Copia_String(sLocal, sAux)
          
    Exit Function
    
Erro_ApuracaoICMSRes_GuiaLocal:

    ApuracaoICMSRes_GuiaLocal = gErr
     
    Select Case gErr
          
        Case 78027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153019)
     
    End Select
     
    Exit Function

End Function

Public Function NF_ObterCodProd(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Código do Produto

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterCodProd

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128112

    End If

    NF_ObterCodProd = objRelNFProd.NF_ObterCodProd(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterCodProd:

    NF_ObterCodProd = gErr

    Select Case gErr

        Case 128112

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153020)

    End Select

    Exit Function

End Function

Public Function NF_ObterDescProd(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Descrição do Produto

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterDescProd

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128113

    End If

    NF_ObterDescProd = objRelNFProd.NF_ObterDescProd(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterDescProd:

    NF_ObterDescProd = gErr

    Select Case gErr

        Case 128113

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153021)

    End Select

    Exit Function

End Function

Public Function NF_ObterCF(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a Classificação Fiscal

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterCF

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128114

    End If

    NF_ObterCF = objRelNFProd.NF_ObterCF(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterCF:

    NF_ObterCF = gErr

    Select Case gErr

        Case 128114

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153022)

    End Select

    Exit Function

End Function

Public Function NF_ObterST(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a Situação Tributária

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterST

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128115

    End If

    NF_ObterST = objRelNFProd.NF_ObterST(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterST:

    NF_ObterST = gErr

    Select Case gErr

        Case 128115

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153023)

    End Select

    Exit Function

End Function

Public Function NF_ObterUM(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a Unidade de Medida

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterUM

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128116

    End If

    NF_ObterUM = objRelNFProd.NF_ObterUM(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterUM:

    NF_ObterUM = gErr

    Select Case gErr

        Case 128116

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153024)

    End Select

    Exit Function

End Function

Public Function NF_ObterQuantidade(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a Quantidade

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterQuantidade

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128117

    End If

    NF_ObterQuantidade = objRelNFProd.NF_ObterQuantidade(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterQuantidade:

    NF_ObterQuantidade = gErr

    Select Case gErr

        Case 128117

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153025)

    End Select

    Exit Function

End Function

Public Function NF_ObterPrecoUnitario(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Preço Unitário

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterPrecoUnitario

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128118

    End If

    NF_ObterPrecoUnitario = objRelNFProd.NF_ObterPrecoUnitario(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterPrecoUnitario:

    NF_ObterPrecoUnitario = gErr

    Select Case gErr

        Case 128118

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153026)

    End Select

    Exit Function

End Function

Public Function NF_ObterPrecoTotal(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Preço Total

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterPrecoTotal

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128119

    End If

    NF_ObterPrecoTotal = objRelNFProd.NF_ObterPrecoTotal(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterPrecoTotal:

    NF_ObterPrecoTotal = gErr

    Select Case gErr

        Case 128119

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153027)

    End Select

    Exit Function

End Function

Public Function NF_ObterICMSAliquota(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a Alíquota ICMS

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterICMSAliquota

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128120

    End If

    NF_ObterICMSAliquota = objRelNFProd.NF_ObterICMSAliquota(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterICMSAliquota:

    NF_ObterICMSAliquota = gErr

    Select Case gErr

        Case 128120

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153028)

    End Select

    Exit Function

End Function

Public Function NF_ObterIPIAliquota(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a Alíquota IPI

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterIPIAliquota

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128121

    End If

    NF_ObterIPIAliquota = objRelNFProd.NF_ObterIPIAliquota(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterIPIAliquota:

    NF_ObterIPIAliquota = gErr

    Select Case gErr

        Case 128121

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153029)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorIPI(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Valor IPI

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterValorIPI

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128122

    End If

    NF_ObterValorIPI = objRelNFProd.NF_ObterValorIPI(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterValorIPI:

    NF_ObterValorIPI = gErr

    Select Case gErr

        Case 128122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153030)

    End Select

    Exit Function

End Function

Public Function NF_ObterDescServ(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a descrição do Serviço

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterDescServ

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128123

    End If

    NF_ObterDescServ = objRelNFProd.NF_ObterDescServ(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterDescServ:

    NF_ObterDescServ = gErr

    Select Case gErr

        Case 128123

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153031)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorServ(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Valor do Serviço

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterValorServ

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128124

    End If

    NF_ObterValorServ = objRelNFProd.NF_ObterValorServ(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterValorServ:

    NF_ObterValorServ = gErr

    Select Case gErr

        Case 128124

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153032)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorISS(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Valor do ISS

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterValorISS

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128125

    End If

    NF_ObterValorISS = objRelNFProd.NF_ObterValorISS(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterValorISS:

    NF_ObterValorISS = gErr

    Select Case gErr

        Case 128125

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153033)

    End Select

    Exit Function

End Function

Function Fonte_I25_ConverteString(sRetorno As String, ByVal sCodigo As String) As Long

  Dim sCod As String, sTmp As String
  Dim ii As Integer, jj As Integer, iTmp As Integer

  'Pressupoe-se que existe um numero par de digitos em sCodigo

  'Inicializacao
  sTmp = ""
  iTmp = 0
  ii = 1
  jj = Len(sCodigo)
  sCod = ""

  Do While ii < jj
    'Separando os digitos dois a dois
    sTmp = Mid(sCodigo, ii, 2)
    ii = ii + 2
    iTmp = CInt(sTmp)

    'A + ABS(A<=49)*48 + ABS(A>=50)*142
    iTmp = iTmp + Abs(iTmp <= 49) * 48 + Abs(iTmp >= 50) * 142

    'Pegando o caracter da conta acima.
    sCod = sCod & Chr(iTmp)
  Loop

  'Delimitadores - inicial e final
  If (sCod <> vbNullString) Then
    sCod = "(" + sCod + ")"
  End If

  'Retorno da funcao
  Call Copia_String(sRetorno, sCod)

    Fonte_I25_ConverteString = SUCESSO
    
End Function

Public Function ObterComissNumPV(lNumPV As Long, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna o numero do pedido associado à uma comissao se houver

Dim lErro As Long, objRelComissao As ClassRelComissao

On Error GoTo Erro_ObterComissNumPV

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then gError 133095

    End If

    ObterComissNumPV = objRelComissao.Comissao_Obter_NumPV(iTipoComiss, lNumIntComiss, lNumPV, dtDataGeracao)

    Exit Function

Erro_ObterComissNumPV:

    ObterComissNumPV = gErr

    Select Case gErr

        Case 133095

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153034)

    End Select

    Exit Function

End Function

Public Function NF_ObterValorUnitServ(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Valor do Serviço

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterValorUnitServ

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128124

    End If

    NF_ObterValorUnitServ = objRelNFProd.NF_ObterValorUnitServ(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterValorUnitServ:

    NF_ObterValorUnitServ = gErr

    Select Case gErr

        Case 128124

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153035)

    End Select

    Exit Function

End Function

Public Function NF_ObterQtdeServ(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Valor do Serviço

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterQtdeServ

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128124

    End If

    NF_ObterQtdeServ = objRelNFProd.NF_ObterQtdeServ(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterQtdeServ:

    NF_ObterQtdeServ = gErr

    Select Case gErr

        Case 128124

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153036)

    End Select

    Exit Function

End Function

Public Function NF_ObterUMServ(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Valor do Serviço

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterUMServ

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 128124

    End If

    NF_ObterUMServ = objRelNFProd.NF_ObterUMServ(sAux, lNumIntDoc, iLinha)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterUMServ:

    NF_ObterUMServ = gErr

    Select Case gErr

        Case 128124

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153037)

    End Select

    Exit Function

End Function

Function MargContr_ObtemPrecoCli(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer) As Long

Dim lErro As Long, objRelMargContr As ClassRelMargContr

On Error GoTo Erro_MargContr_ObtemPrecoCli

    Set objRelMargContr = ObtemObj("RelMargContr")
    If objRelMargContr Is Nothing Then

        Set objRelMargContr = New ClassRelMargContr
        lErro = GuardaObj("RelMargContr", objRelMargContr)
        If lErro <> SUCESSO Then gError 130040

    End If

    MargContr_ObtemPrecoCli = objRelMargContr.MargContr_ObtemPrecoCli(dValor, iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, 3, 4, 17, 20)
     
    Exit Function
    
Erro_MargContr_ObtemPrecoCli:

    MargContr_ObtemPrecoCli = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153038)
     
    End Select
     
    Exit Function

End Function

Function MargContr_ObtemPrecoPrat(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer) As Long

Dim lErro As Long, objRelMargContr As ClassRelMargContr

On Error GoTo Erro_MargContr_ObtemPrecoPrat

    Set objRelMargContr = ObtemObj("RelMargContr")
    If objRelMargContr Is Nothing Then

        Set objRelMargContr = New ClassRelMargContr
        lErro = GuardaObj("RelMargContr", objRelMargContr)
        If lErro <> SUCESSO Then gError 130040

    End If

    MargContr_ObtemPrecoPrat = objRelMargContr.MargContr_ObtemPrecoPrat(dValor, iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, 3, 4, 17, 20)
     
    Exit Function
    
Erro_MargContr_ObtemPrecoPrat:

    MargContr_ObtemPrecoPrat = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153039)
     
    End Select
     
    Exit Function

End Function

Function MargContr_ObtemMargPrat(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer) As Long

Dim lErro As Long, objRelMargContr As ClassRelMargContr

On Error GoTo Erro_MargContr_ObtemMargPrat

    Set objRelMargContr = ObtemObj("RelMargContr")
    If objRelMargContr Is Nothing Then

        Set objRelMargContr = New ClassRelMargContr
        lErro = GuardaObj("RelMargContr", objRelMargContr)
        If lErro <> SUCESSO Then gError 130040

    End If

    MargContr_ObtemMargPrat = objRelMargContr.MargContr_ObtemMargPrat(dValor, iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, 3, 4, 17, 20)
     
    Exit Function
    
Erro_MargContr_ObtemMargPrat:

    MargContr_ObtemMargPrat = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153040)
     
    End Select
     
    Exit Function

End Function

Function MargContr_ObtemLucroPrat(dValor As Double, ByVal iFilialFaturamento As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, ByVal iTabelaPreco As Integer) As Long

Dim lErro As Long, objRelMargContr As ClassRelMargContr

On Error GoTo Erro_MargContr_ObtemLucroPrat

    Set objRelMargContr = ObtemObj("RelMargContr")
    If objRelMargContr Is Nothing Then

        Set objRelMargContr = New ClassRelMargContr
        lErro = GuardaObj("RelMargContr", objRelMargContr)
        If lErro <> SUCESSO Then gError 130040

    End If

    MargContr_ObtemLucroPrat = objRelMargContr.MargContr_ObtemLucroPrat(dValor, iFilialFaturamento, lCliente, iFilialCli, sProduto, dQtde, sUM, iTabelaPreco, 3, 4, 17, 20)
     
    Exit Function
    
Erro_MargContr_ObtemLucroPrat:

    MargContr_ObtemLucroPrat = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153041)
     
    End Select
     
    Exit Function

End Function

Function ProdutoCli_Obtem_PrecoDesde(sRetorno As String, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal iFilial As Integer, ByVal sProduto As String, ByVal dtData As Date, ByVal dValor As Double) As Long

Dim lErro As Long, objRelMargContr As ClassRelMargContr
Dim sDesde As String

On Error GoTo Erro_ProdutoCli_Obtem_PrecoDesde

    Set objRelMargContr = ObtemObj("RelMargContr")
    If objRelMargContr Is Nothing Then

        Set objRelMargContr = New ClassRelMargContr
        lErro = GuardaObj("RelMargContr", objRelMargContr)
        If lErro <> SUCESSO Then gError 130040

    End If

    ProdutoCli_Obtem_PrecoDesde = objRelMargContr.ProdutoCli_Obtem_PrecoDesde(iFilialEmpresa, lCliente, iFilial, sProduto, dtData, dValor, sDesde)
     
    Call Copia_String(sRetorno, sDesde)
     
    Exit Function
    
Erro_ProdutoCli_Obtem_PrecoDesde:

    ProdutoCli_Obtem_PrecoDesde = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153038)
     
    End Select
     
    Exit Function

End Function

Public Function ObterPaginasNFCromaton(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal iMaxLinhasPagina As Integer, ByVal iMaxCaracLinha As Integer, ByVal iLinhasDesconto As Integer, ByVal iLinhasSuframa As Integer) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim iNumPagina As Integer

On Error GoTo Erro_ObterPaginasNFCromaton

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 69091

    End If

    ObterPaginasNFCromaton = objRelNF.ObterPaginasNFCromaton(lNumIntDoc, iMaxLinhasPagina, iNumPagina, iMaxCaracLinha, iLinhasDesconto, iLinhasSuframa)

    dNumPagina = iNumPagina

    Exit Function

Erro_ObterPaginasNFCromaton:

    ObterPaginasNFCromaton = gErr

    Select Case gErr

        Case 69091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153042)

    End Select

    Exit Function

End Function

Public Function ItemNF_ObterLote(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iSeq As Integer) As Long
'Para obter a Base de Calculo

Dim lErro As Long, objRelItemNFLote As ClassRelItemNFLote, sAux As String

On Error GoTo Erro_ItemNF_ObterLote

    Set objRelItemNFLote = ObtemObj("RelItemNFLote")
    If objRelItemNFLote Is Nothing Then

        Set objRelItemNFLote = New ClassRelItemNFLote
        lErro = GuardaObj("RelItemNFLote", objRelItemNFLote)
        If lErro <> SUCESSO Then gError 90342

    End If

    ItemNF_ObterLote = objRelItemNFLote.ItemNF_ObterLote(lNumIntDoc, iSeq, sAux)
    Call Copia_String(sRetorno, sAux)
    
    Exit Function

Erro_ItemNF_ObterLote:

    ItemNF_ObterLote = gErr

    Select Case gErr

        Case 90342

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153043)

    End Select

    Exit Function

End Function

Public Function ItemNF_ObterLoteQtd(dQuantidade As Double, ByVal lNumIntDoc As Long, ByVal iSeq As Integer) As Long
'Para obter a Base de Calculo

Dim lErro As Long, objRelItemNFLote As ClassRelItemNFLote

On Error GoTo Erro_ItemNF_ObterLoteQtd

    Set objRelItemNFLote = ObtemObj("RelItemNFLote")
    If objRelItemNFLote Is Nothing Then

        Set objRelItemNFLote = New ClassRelItemNFLote
        lErro = GuardaObj("RelItemNFLote", objRelItemNFLote)
        If lErro <> SUCESSO Then gError 90342

    End If

    ItemNF_ObterLoteQtd = objRelItemNFLote.ItemNF_ObterLoteQtd(lNumIntDoc, iSeq, dQuantidade)
    
    Exit Function

Erro_ItemNF_ObterLoteQtd:

    ItemNF_ObterLoteQtd = gErr

    Select Case gErr

        Case 90342

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153044)

    End Select

    Exit Function

End Function

'####################################################
'Inserido por Wagner
Public Function ObterSldIniCCIConc(dSldIni As Double, ByVal iCodCCI As Integer, ByVal dtData As Date) As Long
'Para obter o saldo inicial de uma conta numa determinada data

Dim lErro As Long
Dim objRelTes As ClassRelTes

On Error GoTo Erro_ObterSldIniCCIConc

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 132071

    End If

    ObterSldIniCCIConc = objRelTes.ObterSldIniConc(iCodCCI, dtData, dSldIni)

    Exit Function

Erro_ObterSldIniCCIConc:

    ObterSldIniCCIConc = gErr

    Select Case gErr

        Case 132071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153045)

    End Select

    Exit Function

End Function

Public Function ObterCCIHistContabil(sRetorno As String, ByVal lNumMovto) As Long
'Para obter o histórico contábil

Dim lErro As Long
Dim objRelTes As ClassRelTes
Dim sAux As String

On Error GoTo Erro_ObterCCIHistContabil

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError 132181

    End If

    ObterCCIHistContabil = objRelTes.ObterHistContabil(sAux, lNumMovto)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_ObterCCIHistContabil:

    ObterCCIHistContabil = gErr

    Select Case gErr

        Case 132181

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153046)

    End Select

    Exit Function

End Function

'####################################
'Inserido por Wagner
Public Function NF_ObterCodServ(sRetorno As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Código do Produto

Dim lErro As Long, objRelNFProd As ClassRelNFProd, sAux As String

On Error GoTo Erro_NF_ObterCodServ

    Set objRelNFProd = ObtemObj("RelNFProd")
    If objRelNFProd Is Nothing Then

        Set objRelNFProd = New ClassRelNFProd
        lErro = GuardaObj("RelNFProd", objRelNFProd)
        If lErro <> SUCESSO Then gError 138385

    End If

    NF_ObterCodServ = objRelNFProd.NF_ObterCodServ(sAux, lNumIntDoc, iLinha)

    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterCodServ:

    NF_ObterCodServ = gErr

    Select Case gErr

        Case 138385

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153047)

    End Select

    Exit Function

End Function
'####################################

Public Function NF_ObterAnotacao(sRetorno As String, ByVal lNumIntDoc As Long, ByVal sTitulo As String, ByVal iPosInicial As Integer, ByVal iTam As Integer) As Long
'Devolve p/ o gerador de relatório o Código do Produto

Dim lErro As Long, objRelAnotacaoNF As ClassRelAnotacaoNF, sAux As String

On Error GoTo Erro_NF_ObterAnotacao

    Set objRelAnotacaoNF = ObtemObj("RelAnotacaoNF")
    If objRelAnotacaoNF Is Nothing Then

        Set objRelAnotacaoNF = New ClassRelAnotacaoNF
        lErro = GuardaObj("RelAnotacaoNF", objRelAnotacaoNF)
        If lErro <> SUCESSO Then gError 130407

    End If

    NF_ObterAnotacao = objRelAnotacaoNF.NF_ObterAnotacao(sAux, lNumIntDoc, sTitulo, iPosInicial, iTam)

    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterAnotacao:

    NF_ObterAnotacao = gErr

    Select Case gErr

        Case 130407

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153048)

    End Select

    Exit Function

End Function

'####################################################################
'Inserido por Wagner 27/10/2005
Public Function MapaECF_ObterValorAliquota(dValor As Double, ByVal dtData As Date, ByVal iCodECF As Integer, ByVal sSituacaoTrib As String) As Long

Dim lErro As Long
Dim objRelMapaECF As ClassRelMapaECF

On Error GoTo Erro_MapaECF_ObterValorAliquota

    Set objRelMapaECF = ObtemObj("RelMapaECF")
    If objRelMapaECF Is Nothing Then

        Set objRelMapaECF = New ClassRelMapaECF
        lErro = GuardaObj("RelMapaECF", objRelMapaECF)
        If lErro <> SUCESSO Then gError 140613

    End If

    MapaECF_ObterValorAliquota = objRelMapaECF.MapaECF_ObterValorAliquota(dValor, dtData, iCodECF, sSituacaoTrib)

    Exit Function

Erro_MapaECF_ObterValorAliquota:

    MapaECF_ObterValorAliquota = gErr

    Select Case gErr

        Case 140613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153049)

    End Select

    Exit Function

End Function

Public Function MapaECF_ObterAliquota(dAliquota As Double, ByVal dtData As Date, ByVal iPOS As Integer) As Long

Dim lErro As Long
Dim objRelMapaECF As ClassRelMapaECF

On Error GoTo Erro_MapaECF_ObterAliquota

    Set objRelMapaECF = ObtemObj("RelMapaECF")
    If objRelMapaECF Is Nothing Then

        Set objRelMapaECF = New ClassRelMapaECF
        lErro = GuardaObj("RelMapaECF", objRelMapaECF)
        If lErro <> SUCESSO Then gError 140614

    End If

    MapaECF_ObterAliquota = objRelMapaECF.MapaECF_ObterAliquota(dAliquota, dtData, iPOS)

    Exit Function

Erro_MapaECF_ObterAliquota:

    MapaECF_ObterAliquota = gErr

    Select Case gErr

        Case 140614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153050)

    End Select

    Exit Function

End Function
'############################################################################

'####################################################################
'Inserido por Wagner 09/02/2006
Public Function Obter_Bloqueios_PV(iBloqueios As Integer, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde de bloqueios ativos (nao liberados) de um pedido de venda

Dim lErro As Long
Dim objRelPedVenda As ClassRelPedVenda

On Error GoTo Erro_Obter_Bloqueios_PV

    Set objRelPedVenda = ObtemObj("RelPedVenda")

    If objRelPedVenda Is Nothing Then

        Set objRelPedVenda = New ClassRelPedVenda

        lErro = GuardaObj("RelPedVenda", objRelPedVenda)
        If lErro <> SUCESSO Then gError 141739

    End If

    Obter_Bloqueios_PV = objRelPedVenda.Obter_Bloqueios_PV(iBloqueios, lCodigo, iFilialEmpresa)

    Exit Function

Erro_Obter_Bloqueios_PV:

    Obter_Bloqueios_PV = gErr

    Select Case gErr

        Case 141739 'Tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 130462)

    End Select

    Exit Function

End Function
'####################################################################

Public Function PedVendaObterDataParc(dtData As Date, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, ByVal iParcela As Integer) As Long
'para obter a data de vencimento da parcela a receber

Dim lErro As Long
Dim objRelPedVenda As ClassRelPedVenda

On Error GoTo Erro_PedVendaObterDataParc

    Set objRelPedVenda = ObtemObj("RelPedVenda")
    If objRelPedVenda Is Nothing Then

        Set objRelPedVenda = New ClassRelPedVenda
        lErro = GuardaObj("RelPedVenda", objRelPedVenda)
        If lErro <> SUCESSO Then Error 58221

    End If

    PedVendaObterDataParc = objRelPedVenda.ObterDataParcRec(dtData, lCodigo, iFilialEmpresa, iParcela)

    Exit Function

Erro_PedVendaObterDataParc:

    PedVendaObterDataParc = Err

    Select Case Err

        Case 58221

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 130464)

    End Select

    Exit Function

End Function

Public Function PedVendaObterValorParc(dValor As Double, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, ByVal iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long
Dim objRelPedVenda As ClassRelPedVenda

On Error GoTo Erro_PedVendaObterValorParc

    Set objRelPedVenda = ObtemObj("RelPedVenda")
    If objRelPedVenda Is Nothing Then

        Set objRelPedVenda = New ClassRelPedVenda
        lErro = GuardaObj("RelPedVenda", objRelPedVenda)
        If lErro <> SUCESSO Then Error 58222

    End If

    PedVendaObterValorParc = objRelPedVenda.ObterValorParcRec(dValor, lCodigo, iFilialEmpresa, iParcela)

    Exit Function

Erro_PedVendaObterValorParc:

    PedVendaObterValorParc = Err

    Select Case Err

        Case 58222

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 130465)

    End Select

    Exit Function

End Function

Function RCPEMod3_ObterCodigo(dCodigo As Double, ByVal iQuebra As Integer, ByVal sEntradaSaida As String, ByVal iTipoMov As Integer, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal iNatureza As Integer, ByVal iProdNaFilial As Integer) As Long

Dim lErro As Long, iCodigo As Integer
Dim objRelRCPEMod3 As ClassRelRCPEMod3

On Error GoTo Erro_RCPEMod3_ObterCodigo

    Set objRelRCPEMod3 = ObtemObj("RelRCPEMod3")
    If objRelRCPEMod3 Is Nothing Then

        Set objRelRCPEMod3 = New ClassRelRCPEMod3
        lErro = GuardaObj("RelRCPEMod3", objRelRCPEMod3)
        If lErro <> SUCESSO Then gError 140614

    End If

    RCPEMod3_ObterCodigo = objRelRCPEMod3.ObterCodigo(iCodigo, iQuebra, sEntradaSaida, iTipoMov, sProduto, iFilialEmpresa, iNatureza, iProdNaFilial)
    dCodigo = iCodigo
    
    Exit Function

Erro_RCPEMod3_ObterCodigo:

    RCPEMod3_ObterCodigo = gErr

    Select Case gErr

        Case 140614

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 130494)

    End Select

    Exit Function

End Function

'##################################################
'Inserido por Wagner
Public Function ObterComissRazaoSocialCli(sRazaoSocial As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long, ByVal dtDataGeracao As Date) As Long
'Retorna o nome reduzido de um cliente associado à uma comissao

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sAux As String

On Error GoTo Erro_ObterComissRazaoSocialCli

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then gError 132386

    End If

    ObterComissRazaoSocialCli = objRelComissao.Comissao_Obter_RazaoSocialCli(iTipoComiss, lNumIntComiss, sAux, dtDataGeracao)

    Call Copia_String(sRazaoSocial, sAux)

    Exit Function

Erro_ObterComissRazaoSocialCli:

    ObterComissRazaoSocialCli = gErr

    Select Case gErr

        Case 132386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function
'###############################################################

'Public Function Obter_Classificacao_FiscalHar(sClass As String, ByVal iPOS As Integer) As Long
'
'Dim lErro As Long
'Dim objClassiFiscal As ClassRelClassiFiscalHar
'Dim sAux As String
'
'On Error GoTo Erro_Obter_Classificacao_FiscalHar
'
'    Set objClassiFiscal = ObtemObj("ClassiFiscalHar")
'
'    If objClassiFiscal Is Nothing Then
'
'        Set objClassiFiscal = New ClassRelClassiFiscalHar
'
'        lErro = GuardaObj("ClassiFiscalHar", objClassiFiscal)
'        If lErro <> SUCESSO Then Error 61277
'
'    End If
'
'    Call objClassiFiscal.Obter_Classificacao_Fiscal(sAux, iPOS)
'
'    Call Copia_String(sClass, sAux)
'
'    Obter_Classificacao_FiscalHar = SUCESSO
'
'    Exit Function
'
'Erro_Obter_Classificacao_FiscalHar:
'
'    Obter_Classificacao_FiscalHar = Err
'
'    Select Case Err
'
'        Case 61277
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152853)
'
'    End Select
'
'    Exit Function
'
'End Function

'Public Function Obter_LetraHar(sLetra As String, ByVal sClass As String, ByVal lNumIntDoc As Long) As Long
''para os itens da nota fiscal
'
'Dim lErro As Long
'Dim objClassiFiscal As ClassRelClassiFiscalHar
'Dim sAux As String
'
'On Error GoTo Erro_Obter_LetraHar
'
'    Set objClassiFiscal = ObtemObj("ClassiFiscalHar")
'
'    If objClassiFiscal Is Nothing Then
'
'        Set objClassiFiscal = New ClassRelClassiFiscalHar
'
'        lErro = GuardaObj("ClassiFiscalHar", objClassiFiscal)
'        If lErro <> SUCESSO Then Error 61278
'
'    End If
'
'    Call objClassiFiscal.Obter_Letra(sAux, sClass, lNumIntDoc)
'
'    Call Copia_String(sLetra, sAux)
'
'    Obter_LetraHar = SUCESSO
'
'    Exit Function
'
'Erro_Obter_LetraHar:
'
'    Obter_LetraHar = Err
'
'    Select Case Err
'
'        Case 61278
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152854)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function ObterQuantInicialLote(dQuantInicial As Double, ByVal sProduto As String, ByVal sLote As String, ByVal iFilialOP As Integer, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date) As Long
'obtem  a qtde inicial de um produto/lote/almoxarifado antes da data inicial
'dQuantInicial é a quantidade inicial do produto sProduto do lote sLote/iFilialOP no Almoxarifado iAlmoxarifado antes da data inicial dtDataInicial, que é retornada
'Caso dQuantInicial = -1 significa que a quantidade não está disponível.
'sProduto - é o codigo do produto em questao
'iAlmoxarifado é o codigo do almoxarifado em questão.
'dtDataInicial é a data inicial do periodo que está sendo consultado ou data nula

Dim lErro As Long, objRelSaldoLoteRastro As ClassRelSaldoLoteRastro

On Error GoTo Erro_ObterQuantInicialLote

    Set objRelSaldoLoteRastro = ObtemObj("RelSaldoLoteRastro")
    If objRelSaldoLoteRastro Is Nothing Then

        Set objRelSaldoLoteRastro = New ClassRelSaldoLoteRastro
        lErro = GuardaObj("RelSaldoLoteRastro", objRelSaldoLoteRastro)
        If lErro <> SUCESSO Then Error 130538

    End If

    ObterQuantInicialLote = objRelSaldoLoteRastro.ObterQuantInicialLote(dQuantInicial, sProduto, sLote, iFilialOP, iAlmoxarifado, dtDataInicial)

    Exit Function

Erro_ObterQuantInicialLote:

    ObterQuantInicialLote = Err

    Select Case Err

        Case 130538

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 130539)

    End Select

    Exit Function

End Function

Public Function ObterPaginasNFHar(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal iMaxLinhasPagina As Integer, ByVal iMaxCaracLinha As Integer, ByVal iLinhasDesconto As Integer, ByVal iLinhasSuframa As Integer) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim iNumPagina As Integer

On Error GoTo Erro_ObterPaginasNFHar

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 181820

    End If

    ObterPaginasNFHar = objRelNF.ObterPaginasNFHar(lNumIntDoc, iMaxLinhasPagina, iNumPagina, iMaxCaracLinha, iLinhasDesconto, iLinhasSuframa)

    dNumPagina = iNumPagina

    Exit Function

Erro_ObterPaginasNFHar:

    ObterPaginasNFHar = gErr

    Select Case gErr

        Case 181820

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181821)

    End Select

    Exit Function

End Function

Public Function ObterQuantDispLoteFilial(dQuantDisp As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long
'obtem  a qtde disponivel associada a lotes de rastreamento em uma filial

Dim lErro As Long, objRelSaldoLoteRastro As ClassRelSaldoLoteRastro

On Error GoTo Erro_ObterQuantDispLoteFilial

    Set objRelSaldoLoteRastro = ObtemObj("RelSaldoLoteRastro")
    If objRelSaldoLoteRastro Is Nothing Then

        Set objRelSaldoLoteRastro = New ClassRelSaldoLoteRastro
        lErro = GuardaObj("RelSaldoLoteRastro", objRelSaldoLoteRastro)
        If lErro <> SUCESSO Then gError 184104

    End If

    ObterQuantDispLoteFilial = objRelSaldoLoteRastro.ObterQuantDispLoteFilial(dQuantDisp, sProduto, iFilialEmpresa)

    Exit Function

Erro_ObterQuantDispLoteFilial:

    ObterQuantDispLoteFilial = gErr

    Select Case gErr

        Case 184104

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184105)

    End Select

    Exit Function

End Function

Function Fonte_EAN13_ConverteString(sRetorno As String, ByVal sCodigo As String) As Long

Dim sCod As String

    sCod = sCodigo
    sCod = EAN13(sCod)
    
    'Retorno da funcao
    Call Copia_String(sRetorno, sCod)

    Fonte_EAN13_ConverteString = SUCESSO
    
End Function

Public Function ObterSaldoParcRec(dSaldo As Double, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter o saldo da parcela a receber

Dim lErro As Long, objRelTitRec As ClassRelTitRec

On Error GoTo Erro_ObterSaldoParcRec

    Set objRelTitRec = ObtemObj("RelTitRec")
    If objRelTitRec Is Nothing Then

        Set objRelTitRec = New ClassRelTitRec
        lErro = GuardaObj("RelTitRec", objRelTitRec)
        If lErro <> SUCESSO Then Error 58222

    End If

    ObterSaldoParcRec = objRelTitRec.ObterSaldoParcRec(dSaldo, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterSaldoParcRec:

    ObterSaldoParcRec = Err

    Select Case Err

        Case 58222

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152834)

    End Select

    Exit Function

End Function

Public Function ObterValorFatProdPer_SemDev(dValorFaturado As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o valor faturado por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterValorFatProdPer_SemDev

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then Error 43318

    End If

    ObterValorFatProdPer_SemDev = objRelMesFat.ObterValorFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorFaturado, True)

    Exit Function

Erro_ObterValorFatProdPer_SemDev:

    ObterValorFatProdPer_SemDev = Err

    Select Case Err

        Case 43318

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152789)

    End Select

    Exit Function

End Function

Public Function ObterRendAplicPer(dRendimentos As Double, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o valor dos rendimentos das aplicacoes de uma filialempresa num periodo

Dim lErro As Long, objRelRendAplic As ClassRelRendAplic

On Error GoTo Erro_ObterRendAplicPer

    Set objRelRendAplic = ObtemObj("RelRendAplic")
    If objRelRendAplic Is Nothing Then

        Set objRelRendAplic = New ClassRelRendAplic
        lErro = GuardaObj("RelRendAplic", objRelRendAplic)
        If lErro <> SUCESSO Then gError 184220

    End If

    ObterRendAplicPer = objRelRendAplic.ObterRendAplicPer(dRendimentos, dtDataInicial, dtDataFinal, iFilialEmpresa)

    Exit Function

Erro_ObterRendAplicPer:

    ObterRendAplicPer = gErr

    Select Case gErr

        Case 184220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184221)

    End Select

    Exit Function

End Function

Public Function RelFluxoCaixaCtb_ObterInfo(dValor As Double, ByVal iTipoInfo As Integer, ByVal lNumIntRel As Long) As Long
'Para obter informacoes auxliares para o relatorio de fluxo de caixa contabil

Dim lErro As Long, objRelFluxoCaixaCtb As ClassRelFluxoCaixaCtb

On Error GoTo Erro_RelFluxoCaixaCtb_ObterInfo

    Set objRelFluxoCaixaCtb = ObtemObj("RelFluxoCaixaCtb")
    If objRelFluxoCaixaCtb Is Nothing Then

        Set objRelFluxoCaixaCtb = New ClassRelFluxoCaixaCtb
        lErro = GuardaObj("RelFluxoCaixaCtb", objRelFluxoCaixaCtb)
        If lErro <> SUCESSO Then gError 184258

    End If

    RelFluxoCaixaCtb_ObterInfo = objRelFluxoCaixaCtb.ObterInfo(dValor, iTipoInfo, lNumIntRel)

    Exit Function

Erro_RelFluxoCaixaCtb_ObterInfo:

    RelFluxoCaixaCtb_ObterInfo = gErr

    Select Case gErr

        Case 184258

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184259)

    End Select

    Exit Function

End Function

Public Function Formata_Produto_Rel(sProdutoFormatado As String, ByVal sProdutoMascarado As String) As Long
'Função para obter o numero do Lote atravez do Numero Interno da Nota Fiscal

Dim lErro As Long
Dim sAux As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Formata_Produto_Rel

    'Preenche o código de objProduto
    lErro = CF("Produto_Formata", sProdutoMascarado, sAux, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 111300
        
    Call Copia_String(sProdutoFormatado, sAux)
    
    Formata_Produto_Rel = SUCESSO

    Exit Function

Erro_Formata_Produto_Rel:

    Formata_Produto_Rel = gErr

    Select Case gErr

        Case 111300

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152999)

    End Select

End Function

Public Function Obter_Produto_Referencia(sReferencia As String, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Produto_Referencia

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError 196914

    End If

    Obter_Produto_Referencia = objRelProduto.Obter_Referencia(sAux, sProduto)

    Call Copia_String(sReferencia, sAux)

    Exit Function

Erro_Obter_Produto_Referencia:

    Obter_Produto_Referencia = gErr

    Select Case gErr

        Case 196914

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196915)

    End Select

    Exit Function

End Function

Public Function Obter_Produto_NomeReduzido(sNomeReduzido As String, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Produto_NomeReduzido

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError 196916

    End If

    Obter_Produto_NomeReduzido = objRelProduto.Obter_NomeReduzido(sAux, sProduto)

    Call Copia_String(sNomeReduzido, sAux)

    Exit Function

Erro_Obter_Produto_NomeReduzido:

    Obter_Produto_NomeReduzido = gErr

    Select Case gErr

        Case 196916

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196917)

    End Select

    Exit Function

End Function


Public Function ObterListaCFOP(sListaCFOP As String, ByVal lNumIntDoc As Long) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim sAux As String


On Error GoTo Erro_ObterListaCFOP

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 194857

    End If

    ObterListaCFOP = objRelNF.ObterListaCFOP(lNumIntDoc, sAux)
    Call Copia_String(sListaCFOP, sAux)

    Exit Function

Erro_ObterListaCFOP:

    ObterListaCFOP = gErr

    Select Case gErr

        Case 194857

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194858)

    End Select

    Exit Function

End Function

Public Function OP_ObterQtdEmb_Cro(dQtdEmb As Double, ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long, objRelOpCromaton As ClassRelOpCromaton

On Error GoTo Erro_OP_ObterQtdEmb_Cro

    Set objRelOpCromaton = ObtemObj("RelOpCromaton")
    If objRelOpCromaton Is Nothing Then

        Set objRelOpCromaton = New ClassRelOpCromaton
        lErro = GuardaObj("RelOpCromaton", objRelOpCromaton)
        If lErro <> SUCESSO Then gError 184220

    End If

    OP_ObterQtdEmb_Cro = objRelOpCromaton.ObterQtdEmb_Cro(dQtdEmb, sProduto, iEmbalagem, dQtdProd, iFilialEmpresa, sCodigoOP)

    Exit Function

Erro_OP_ObterQtdEmb_Cro:

    OP_ObterQtdEmb_Cro = gErr

    Select Case gErr

        Case 184220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184221)

    End Select

    Exit Function

End Function

Public Function OP_ObterQtdFrac_Cro(dQtdFrac As Double, ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long, objRelOpCromaton As ClassRelOpCromaton

On Error GoTo Erro_OP_ObterQtdFrac_Cro

    Set objRelOpCromaton = ObtemObj("RelOpCromaton")
    If objRelOpCromaton Is Nothing Then

        Set objRelOpCromaton = New ClassRelOpCromaton
        lErro = GuardaObj("RelOpCromaton", objRelOpCromaton)
        If lErro <> SUCESSO Then gError 184220

    End If

    OP_ObterQtdFrac_Cro = objRelOpCromaton.ObterQtdFrac_Cro(dQtdFrac, sProduto, iEmbalagem, dQtdProd, iFilialEmpresa, sCodigoOP)

    Exit Function

Erro_OP_ObterQtdFrac_Cro:

    OP_ObterQtdFrac_Cro = gErr

    Select Case gErr

        Case 184220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184221)

    End Select

    Exit Function

End Function

Public Function OP_ObterDescEmb_Cro(sDescEmb As String, ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long, objRelOpCromaton As ClassRelOpCromaton
Dim sAux As String

On Error GoTo Erro_OP_ObterDescEmb_Cro

    Set objRelOpCromaton = ObtemObj("RelOpCromaton")
    If objRelOpCromaton Is Nothing Then

        Set objRelOpCromaton = New ClassRelOpCromaton
        lErro = GuardaObj("RelOpCromaton", objRelOpCromaton)
        If lErro <> SUCESSO Then gError 184220

    End If

    OP_ObterDescEmb_Cro = objRelOpCromaton.ObterDescEmb_Cro(sAux, sProduto, iEmbalagem, dQtdProd, iFilialEmpresa, sCodigoOP)
    Call Copia_String(sDescEmb, sAux)

    Exit Function

Erro_OP_ObterDescEmb_Cro:

    OP_ObterDescEmb_Cro = gErr

    Select Case gErr

        Case 184220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184221)

    End Select

    Exit Function

End Function

Public Function Obter_Bloqueios_PV2(dBloqueios As Double, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer) As Long
'Para obter a qtde de bloqueios ativos (nao liberados) de um pedido de venda

Dim lErro As Long, iBloqueios As Integer
Dim objRelPedVenda As ClassRelPedVenda

On Error GoTo Erro_Obter_Bloqueios_PV2

    Set objRelPedVenda = ObtemObj("RelPedVenda")

    If objRelPedVenda Is Nothing Then

        Set objRelPedVenda = New ClassRelPedVenda

        lErro = GuardaObj("RelPedVenda", objRelPedVenda)
        If lErro <> SUCESSO Then gError 141739

    End If

    Obter_Bloqueios_PV2 = objRelPedVenda.Obter_Bloqueios_PV(iBloqueios, lCodigo, iFilialEmpresa)
    dBloqueios = iBloqueios
    
    Exit Function

Erro_Obter_Bloqueios_PV2:

    Obter_Bloqueios_PV2 = gErr

    Select Case gErr

        Case 141739 'Tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 130462)

    End Select

    Exit Function

End Function

Public Function ObterQuantDispLoteFilialComLoc(dQuantDisp As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long
'obtem  a qtde disponivel associada a lotes de rastreamento em uma filial

Dim lErro As Long, objRelSaldoLoteRastro As ClassRelSaldoLoteRastro

On Error GoTo Erro_ObterQuantDispLoteFilialComLoc

    Set objRelSaldoLoteRastro = ObtemObj("RelSaldoLoteRastro")
    If objRelSaldoLoteRastro Is Nothing Then

        Set objRelSaldoLoteRastro = New ClassRelSaldoLoteRastro
        lErro = GuardaObj("RelSaldoLoteRastro", objRelSaldoLoteRastro)
        If lErro <> SUCESSO Then gError 190920

    End If

    ObterQuantDispLoteFilialComLoc = objRelSaldoLoteRastro.ObterQuantDispLoteFilialComLoc(dQuantDisp, sProduto, iFilialEmpresa)

    Exit Function

Erro_ObterQuantDispLoteFilialComLoc:

    ObterQuantDispLoteFilialComLoc = gErr

    Select Case gErr

        Case 190920

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190921)

    End Select

    Exit Function

End Function

Public Function Anotacao_ObterTrecho(sRetorno As String, ByVal iTipoDocOrigem As Integer, ByVal sID As String, ByVal sTitulo As String, ByVal iPosInicial As Integer, ByVal iTam As Integer) As Long
'Devolve p/ o gerador de relatório um trecho da anotacao associada ao objeto

Dim lErro As Long, objRelAnotacao As ClassRelAnotacao, sAux As String

On Error GoTo Erro_Anotacao_ObterTrecho

    Set objRelAnotacao = ObtemObj("RelAnotacao")
    If objRelAnotacao Is Nothing Then

        Set objRelAnotacao = New ClassRelAnotacao
        lErro = GuardaObj("RelAnotacao", objRelAnotacao)
        If lErro <> SUCESSO Then gError 130407

    End If

    Anotacao_ObterTrecho = objRelAnotacao.Anotacao_ObterTrecho(sAux, iTipoDocOrigem, sID, sTitulo, iPosInicial, iTam)

    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_Anotacao_ObterTrecho:

    Anotacao_ObterTrecho = gErr

    Select Case gErr

        Case 130407

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153048)

    End Select

    Exit Function

End Function

Public Function ObterPaginasNFPhar(dNumPagina As Double, ByVal lNumIntDoc As Long, ByVal lMaximoDeLinhas As Long) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim lNumPagina As Long

On Error GoTo Erro_ObterPaginasNFPhar

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 69091

    End If

    ObterPaginasNFPhar = objRelNF.ObterPaginasNFPhar(lNumIntDoc, lMaximoDeLinhas, lNumPagina)

    dNumPagina = lNumPagina

    Exit Function

Erro_ObterPaginasNFPhar:

    ObterPaginasNFPhar = gErr

    Select Case gErr

        Case 69091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152868)

    End Select

    Exit Function

End Function

Public Function ObterQuantBenef3LoteFilialComLoc(dQuantBenef3 As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long
'obtem  a qtde disponivel associada a lotes de rastreamento em uma filial

Dim lErro As Long, objRelSaldoLoteRastro As ClassRelSaldoLoteRastro

On Error GoTo Erro_ObterQuantBenef3LoteFilialComLoc

    Set objRelSaldoLoteRastro = ObtemObj("RelSaldoLoteRastro")
    If objRelSaldoLoteRastro Is Nothing Then

        Set objRelSaldoLoteRastro = New ClassRelSaldoLoteRastro
        lErro = GuardaObj("RelSaldoLoteRastro", objRelSaldoLoteRastro)
        If lErro <> SUCESSO Then gError 190920

    End If

    ObterQuantBenef3LoteFilialComLoc = objRelSaldoLoteRastro.ObterQuantBenef3LoteFilialComLoc(dQuantBenef3, sProduto, iFilialEmpresa)

    Exit Function

Erro_ObterQuantBenef3LoteFilialComLoc:

    ObterQuantBenef3LoteFilialComLoc = gErr

    Select Case gErr

        Case 190920

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190921)

    End Select

    Exit Function

End Function

Public Function NF_ObterAliquotaICMSDAS(dValor As Double, ByVal dtDataEmissao As Date) As Long
'Para obter informacoes auxliares para o relatorio de fluxo de caixa contabil

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NF_ObterAliquotaICMSDAS

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 200969

    End If

    NF_ObterAliquotaICMSDAS = objRelNF.Obtem_AliquotaICMSDAS(dtDataEmissao, dValor)

    Exit Function

Erro_NF_ObterAliquotaICMSDAS:

    NF_ObterAliquotaICMSDAS = gErr

    Select Case gErr

        Case 200969

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200970)

    End Select

    Exit Function

End Function

Public Function Cliente_PossuiCatItem(iPossuiCatItem As Double, ByVal lCliente As Long, ByVal iFilial As Integer, ByVal sCategoria As String, ByVal sItem As String) As Long
'Verifica se o cliente possui a categoria\item em questão

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_Cliente_PossuiCatItem

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 200971

    End If

    Cliente_PossuiCatItem = objRelNF.Cliente_Possui_CatItem(lCliente, iFilial, sCategoria, sItem, iPossuiCatItem)

    Exit Function

Erro_Cliente_PossuiCatItem:

    Cliente_PossuiCatItem = gErr

    Select Case gErr

        Case 200971

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200972)

    End Select

    Exit Function

End Function

Public Function Formata_Num(sTextoFormatado As String, ByVal dValor As Double, ByVal sFormato As String) As Long
'Verifica se o cliente possui a categoria\item em questão

Dim lErro As Long
Dim sTextoAux As String

On Error GoTo Erro_Formata_Num

    sTextoAux = Format(dValor, sFormato)
    Call Copia_String(sTextoFormatado, sTextoAux)

    Formata_Num = SUCESSO

    Exit Function

Erro_Formata_Num:

    Formata_Num = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200972)

    End Select

    Exit Function

End Function

Public Function NF_ObterCredICMSSaidaSimples(dValor As Double, ByVal lNumIntNF As Long, ByVal dtDataEmissao As Date) As Long
'retorna em dValor o credito do ICMS para emissao de nf de venda para empresas no simples

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NF_ObterCredICMSSaidaSimples

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59442

    End If

    NF_ObterCredICMSSaidaSimples = objRelNF.ObterCredICMSSaidaSimples(dValor, lNumIntNF, dtDataEmissao)

    Exit Function

Erro_NF_ObterCredICMSSaidaSimples:

    NF_ObterCredICMSSaidaSimples = Err

    Select Case Err

        Case 59442 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152860)

    End Select

    Exit Function

End Function

Public Function Formata_Code128(sSaida As String, ByVal sEntrada As String) As Long

Dim lErro As Long, objLib As Object

On Error GoTo Erro_Formata_Code128

    Set objLib = CreateObject("Code128DLLCOM.Code128Encoding")
    
    If Not (objLib Is Nothing) Then
        
        sSaida = objLib.GetCode128DataString(sEntrada)
    
    Else
        sSaida = sEntrada
    End If

    Formata_Code128 = SUCESSO

    Exit Function

Erro_Formata_Code128:

    Formata_Code128 = Err

    Select Case Err

        Case 429
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152832)

    End Select

    Exit Function

End Function

Function UltVendasPV_ObtemLinha(sRetorno As String, ByVal iLinha As Integer, ByVal lNumIntPV As Long) As Long

Dim lErro As Long, objRelUltVendasPV As ClassRelUltVendasPV, sAux As String

On Error GoTo Erro_UltVendasPV_ObtemLinha

    Set objRelUltVendasPV = ObtemObj("RelUltVendasPV")
    If objRelUltVendasPV Is Nothing Then

        Set objRelUltVendasPV = New ClassRelUltVendasPV
        lErro = GuardaObj("RelUltVendasPV", objRelUltVendasPV)
        If lErro <> SUCESSO Then gError 130783

    End If

    UltVendasPV_ObtemLinha = objRelUltVendasPV.UltVendasPV_ObtemLinha(sAux, iLinha, lNumIntPV)

    Call Copia_String(sRetorno, sAux)
     
    Exit Function
    
Erro_UltVendasPV_ObtemLinha:

    UltVendasPV_ObtemLinha = gErr
     
    Select Case gErr
          
        Case 130783
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
     
    Exit Function

End Function

Public Function ObterCustoFatProdPer(dValorCusto As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer) As Long
'Para obter o custo por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelMesFat As ClassRelMesFat

On Error GoTo Erro_ObterCustoFatProdPer

'    MsgBox ("entrou em ObterCustoFatProdPer")

    Set objRelMesFat = ObtemObj("RelMesFat")
    If objRelMesFat Is Nothing Then

        Set objRelMesFat = New ClassRelMesFat
        lErro = GuardaObj("RelMesFat", objRelMesFat)
        If lErro <> SUCESSO Then gError 204877

    End If

    ObterCustoFatProdPer = objRelMesFat.ObterCustoFatProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorCusto)

'    MsgBox ("saiu em ObterCustoFatProdPer")

    Exit Function

Erro_ObterCustoFatProdPer:

    ObterCustoFatProdPer = gErr

    Select Case gErr

        Case 204877

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204878)

    End Select

    Exit Function

End Function

Public Function Obter_Produto_UltVlrUnitCompra(dValor As Double, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Produto_UltVlrUnitCompra

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError 196916

    End If

    Obter_Produto_UltVlrUnitCompra = objRelProduto.Obter_UltVlrUnitCompra(dValor, sProduto)

    Exit Function

Erro_Obter_Produto_UltVlrUnitCompra:

    Obter_Produto_UltVlrUnitCompra = gErr

    Select Case gErr

        Case 196916

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196917)

    End Select

    Exit Function

End Function

Public Function CpoGenerico_ObterDesc(sTexto As String, ByVal lCodCampo As Long, ByVal lCodValor As Long) As Long
'Verifica se o cliente possui a categoria\item em questão

Dim lErro As Long, lComando As Long
Dim sTextoAux As String

On Error GoTo Erro_CpoGenerico_ObterDesc

    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 84687
    
    sTextoAux = String(STRING_CAMPOSGENERICOSVALORES_VALOR, 0)
    lErro = Comando_Executar(lComando, "SELECT Valor FROM CamposGenericosValores WHERE CodCampo = ? AND CodValor = ?", _
        sTextoAux, lCodCampo, lCodValor)
    If lErro <> AD_SQL_SUCESSO Then gError 84686
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84686
    
    If lErro = AD_SQL_SEM_DADOS Then sTextoAux = CStr(lCodValor)
    
    Call Comando_Fechar(lComando)
    
    Call Copia_String(sTexto, sTextoAux)

    CpoGenerico_ObterDesc = SUCESSO

    Exit Function

Erro_CpoGenerico_ObterDesc:

    CpoGenerico_ObterDesc = gErr

    Select Case gErr

        Case 84687
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200972)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Sub Class_Terminate()
Dim iIndice As Integer
    For iIndice = 1 To 64
        If Not (gcolChv(iIndice) Is Nothing) Then Set gcolChv(iIndice) = Nothing
    Next
End Sub

Public Function NF_ObterBaseICMSSaidaSimples(dValor As Double, ByVal lNumIntNF As Long) As Long
'retorna em dValor a base do credito do ICMS para emissao de nf de venda para empresas no simples

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NF_ObterBaseICMSSaidaSimples

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59442

    End If

    NF_ObterBaseICMSSaidaSimples = objRelNF.ObterBaseICMSSaidaSimples(dValor, lNumIntNF)

    Exit Function

Erro_NF_ObterBaseICMSSaidaSimples:

    NF_ObterBaseICMSSaidaSimples = Err

    Select Case Err

        Case 59442 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152860)

    End Select

    Exit Function

End Function

Public Function NF_ObterBaseICMSSaidaSimples2(dValor As Double, ByVal lNumIntNF As Long) As Long
'retorna em dValor a base do credito do ICMS para emissao de nf de venda para empresas no simples

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NF_ObterBaseICMSSaidaSimples2

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59442

    End If

    NF_ObterBaseICMSSaidaSimples2 = objRelNF.ObterBaseICMSSaidaSimples2(dValor, lNumIntNF)

    Exit Function

Erro_NF_ObterBaseICMSSaidaSimples2:

    NF_ObterBaseICMSSaidaSimples2 = Err

    Select Case Err

        Case 59442 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152860)

    End Select

    Exit Function

End Function

Public Function NF_ObterCredICMSSaidaSimples2(dValor As Double, ByVal lNumIntNF As Long, ByVal dtDataEmissao As Date) As Long
'retorna em dValor o credito do ICMS para emissao de nf de venda para empresas no simples

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NF_ObterCredICMSSaidaSimples2

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF

        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then Error 59442

    End If

    NF_ObterCredICMSSaidaSimples2 = objRelNF.ObterCredICMSSaidaSimples2(dValor, lNumIntNF, dtDataEmissao)

    Exit Function

Erro_NF_ObterCredICMSSaidaSimples2:

    NF_ObterCredICMSSaidaSimples2 = Err

    Select Case Err

        Case 59442 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152860)

    End Select

    Exit Function

End Function

Public Function NF_ObterTransp(sRetorno As String, ByVal iCodTransp As Integer) As Long
'Devolve p/ o gerador de relatório o Descrição do Produto

Dim lErro As Long, objRelNF As ClassRelNF, sAux As String

On Error GoTo Erro_NF_ObterTransp

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NF_ObterTransp = objRelNF.ObterTransp(sAux, iCodTransp)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_NF_ObterTransp:

    NF_ObterTransp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_Msg_Linha(sRetorno As String, ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long, ByVal iTipoMsg As Integer, ByVal iTamMax As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long, objRelMsg As ClassRelMsg, sAux As String

On Error GoTo Erro_Obter_Msg_Linha

    Set objRelMsg = ObtemObj("RelMsg")
    If objRelMsg Is Nothing Then

        Set objRelMsg = New ClassRelMsg
        lErro = GuardaObj("RelMsg", objRelMsg)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Msg_Linha = objRelMsg.ObterMsg(sAux, iTipoDoc, lNumIntDoc, iTipoMsg, iTamMax, iLinha)
    Call Copia_String(sRetorno, sAux)
    
    Exit Function

Erro_Obter_Msg_Linha:

    Obter_Msg_Linha = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_Anot_Linha(sRetorno As String, ByVal iTipoDoc As Integer, ByVal sID As String, ByVal iTamMax As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long, objRelMsg As ClassRelMsg, sAux As String

On Error GoTo Erro_Obter_Anot_Linha

    Set objRelMsg = ObtemObj("RelMsg")
    If objRelMsg Is Nothing Then

        Set objRelMsg = New ClassRelMsg
        lErro = GuardaObj("RelMsg", objRelMsg)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Anot_Linha = objRelMsg.ObterAnot(sAux, iTipoDoc, sID, iTamMax, iLinha)
    Call Copia_String(sRetorno, sAux)
    
    Exit Function

Erro_Obter_Anot_Linha:

    Obter_Anot_Linha = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_Logo(sRetorno As String) As Long

Dim lErro As Long
Dim sAux As String

On Error GoTo Erro_Obter_Logo

    sAux = ""
    
    lErro = CF("Config_Le", "AdmConfig", "LOCALIZACAO_LOGO", giFilialEmpresa, sAux)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then

        'Le a versão do sistema
        lErro = CF("Config_Le", "AdmConfig", "LOCALIZACAO_LOGO", EMPRESA_TODA, sAux)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
    End If

    Call Copia_String(sRetorno, sAux)
    
    Obter_Logo = SUCESSO

    Exit Function

Erro_Obter_Logo:

    Obter_Logo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_Campo_Danfe(sRetorno As String, ByVal lNumIntItemNF As Long, ByVal iTipoCampo As Integer, ByVal sNomeCampo As String) As Long

Dim lErro As Long, objRelNF As ClassRelNF, sAux As String

On Error GoTo Erro_Obter_Campo_Danfe

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Campo_Danfe = objRelNF.Obter_Campo_Danfe(sAux, lNumIntItemNF, iTipoCampo, sNomeCampo)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_Obter_Campo_Danfe:

    Obter_Campo_Danfe = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function ItemNF_ObterLoteVal(dtValidade As Date, ByVal lNumIntDoc As Long, ByVal iSeq As Integer) As Long
'Para obter a Base de Calculo

Dim lErro As Long, objRelItemNFLote As ClassRelItemNFLote

On Error GoTo Erro_ItemNF_ObterLoteVal

    Set objRelItemNFLote = ObtemObj("RelItemNFLote")
    If objRelItemNFLote Is Nothing Then

        Set objRelItemNFLote = New ClassRelItemNFLote
        lErro = GuardaObj("RelItemNFLote", objRelItemNFLote)
        If lErro <> SUCESSO Then gError 90342

    End If

    ItemNF_ObterLoteVal = objRelItemNFLote.ItemNF_ObterLoteVal(lNumIntDoc, iSeq, dtValidade)
    
    Exit Function

Erro_ItemNF_ObterLoteVal:

    ItemNF_ObterLoteVal = gErr

    Select Case gErr

        Case 90342

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153044)

    End Select

    Exit Function

End Function

Public Function ObterPaginasDanfe(sPaginas As String, ByVal lNumIntDoc As Long, ByVal lMaximoDeLinhasPri As Long, ByVal lMaximoDeLinhasOut As Long) As Long
'Para Obter o Numero de Itens de Uma Nota passada

Dim lErro As Long
Dim objRelNF As ClassRelNF
Dim lNumPagina As Long, sAux As String

On Error GoTo Erro_ObterPaginasDanfe

    Set objRelNF = ObtemObj("RelNF")

    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 69091

    End If

    ObterPaginasDanfe = objRelNF.ObterPaginasDanfe(lNumIntDoc, lMaximoDeLinhasPri, lMaximoDeLinhasOut, sAux)
    Call Copia_String(sPaginas, sAux)

    Exit Function

Erro_ObterPaginasDanfe:

    ObterPaginasDanfe = gErr

    Select Case gErr

        Case 69091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152868)

    End Select

    Exit Function

End Function

Public Function Obter_Campo_Config(sRetorno As String, ByVal sTabela As String, ByVal sCodigo As String) As Long

Dim lErro As Long, objRelNF As ClassRelNF, sAux As String

On Error GoTo Erro_Obter_Campo_Config

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Campo_Config = objRelNF.Obter_Campo_Config(sAux, sTabela, sCodigo)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_Obter_Campo_Config:

    Obter_Campo_Config = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function FlCx_ObterValor(dValor As Double, ByVal lNumIntRel As Long, ByVal iTipo As Integer, ByVal dtData As Date, ByVal dLinha As Double) As Long

Dim lErro As Long, objRelFluxoCxMiguez As ClassRelFluxoCxMiguez

On Error GoTo Erro_FlCx_ObterValor

    Set objRelFluxoCxMiguez = ObtemObj("RelFluxoCxMiguez")
    If objRelFluxoCxMiguez Is Nothing Then
    
        Set objRelFluxoCxMiguez = New ClassRelFluxoCxMiguez
        lErro = GuardaObj("RelFluxoCxMiguez", objRelFluxoCxMiguez)
        If lErro <> SUCESSO Then gError 106975
        
    End If
    
    FlCx_ObterValor = objRelFluxoCxMiguez.FlCx_ObterValor(dValor, lNumIntRel, iTipo, dtData, dLinha)
     
    Exit Function
    
Erro_FlCx_ObterValor:

    FlCx_ObterValor = gErr
     
    Select Case gErr
          
        Case 106975
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function FlCx_ObterDescricao(sDescricao As String, ByVal lNumIntRel As Long, ByVal iTipo As Integer, ByVal dtData As Date, ByVal dLinha As Double) As Long

Dim lErro As Long, objRelFluxoCxMiguez As ClassRelFluxoCxMiguez, sAux As String

On Error GoTo Erro_FlCx_ObterDescricao

    Set objRelFluxoCxMiguez = ObtemObj("RelFluxoCxMiguez")
    If objRelFluxoCxMiguez Is Nothing Then
    
        Set objRelFluxoCxMiguez = New ClassRelFluxoCxMiguez
        lErro = GuardaObj("RelFluxoCxMiguez", objRelFluxoCxMiguez)
        If lErro <> SUCESSO Then gError 106976
        
    End If
    
    FlCx_ObterDescricao = objRelFluxoCxMiguez.FlCx_ObterDescricao(sAux, lNumIntRel, iTipo, dtData, dLinha)
    Call Copia_String(sDescricao, sAux)
     
    Exit Function
    
Erro_FlCx_ObterDescricao:

    FlCx_ObterDescricao = gErr
     
    Select Case gErr
          
        Case 106976
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function Obter_Campo_Texto(sRetorno As String, ByVal sTabela As String, ByVal sCampo As String, ByVal sFiltro As String) As Long

Dim lErro As Long, sAux As String

On Error GoTo Erro_Obter_Campo_Texto

    lErro = CF("Tabela_Le_Campo", sTabela, sCampo, sFiltro, TIPO_CAMPO_STRING, sAux)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    Call Copia_String(sRetorno, sAux)
    
    Obter_Campo_Texto = SUCESSO

    Exit Function

Erro_Obter_Campo_Texto:

    Obter_Campo_Texto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_Campo_Data(dtData As Date, ByVal sTabela As String, ByVal sCampo As String, ByVal sFiltro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_Campo_Data

    lErro = CF("Tabela_Le_Campo", sTabela, sCampo, sFiltro, TIPO_CAMPO_DATE, dtData)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    Obter_Campo_Data = SUCESSO
    
    Exit Function

Erro_Obter_Campo_Data:

    Obter_Campo_Data = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_Campo_Numero(dValor As Double, ByVal sTabela As String, ByVal sCampo As String, ByVal sFiltro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_Campo_Numero

    lErro = CF("Tabela_Le_Campo", sTabela, sCampo, sFiltro, TIPO_CAMPO_DOUBLE, dValor)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    Obter_Campo_Numero = SUCESSO
    
    Exit Function

Erro_Obter_Campo_Numero:

    Obter_Campo_Numero = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function NFSe_Obter_Discriminacao(sTexto As String, ByVal lNumIntDoc As Long, ByVal iLinha As Integer, ByVal iTamanho As Integer) As Long

Dim lErro As Long, objRelNFSE As ClassRelNFSe, sAux As String

On Error GoTo Erro_NFSe_Obter_Discriminacao

    Set objRelNFSE = ObtemObj("RelNFSe")
    If objRelNFSE Is Nothing Then
    
        Set objRelNFSE = New ClassRelNFSe
        lErro = GuardaObj("RelNFSe", objRelNFSE)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    NFSe_Obter_Discriminacao = objRelNFSE.Obter_Discriminacao(lNumIntDoc, iLinha, iTamanho, sAux)
    Call Copia_String(sTexto, sAux)
     
    Exit Function
    
Erro_NFSe_Obter_Discriminacao:

    NFSe_Obter_Discriminacao = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209779)
     
    End Select
     
    Exit Function

End Function

Public Function NFSe_Obter_Municipio(sTexto As String, ByVal lCodMunic As Long) As Long

Dim lErro As Long, objRelNFSE As ClassRelNFSe, sAux As String

On Error GoTo Erro_NFSe_Obter_Municipio

    Set objRelNFSE = ObtemObj("RelNFSe")
    If objRelNFSE Is Nothing Then
    
        Set objRelNFSE = New ClassRelNFSe
        lErro = GuardaObj("RelNFSe", objRelNFSE)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    NFSe_Obter_Municipio = objRelNFSE.Obter_NomeMunic(lCodMunic, sAux)
    Call Copia_String(sTexto, sAux)
     
    Exit Function
    
Erro_NFSe_Obter_Municipio:

    NFSe_Obter_Municipio = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209780)
     
    End Select
     
    Exit Function

End Function

Public Function NFSe_Obter_DescServ(sTexto As String, ByVal lCodMunic As Long, ByVal sCodServ As String) As Long

Dim lErro As Long, objRelNFSE As ClassRelNFSe, sAux As String

On Error GoTo Erro_NFSe_Obter_DescServ

    Set objRelNFSE = ObtemObj("RelNFSe")
    If objRelNFSE Is Nothing Then
    
        Set objRelNFSE = New ClassRelNFSe
        lErro = GuardaObj("RelNFSe", objRelNFSE)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    NFSe_Obter_DescServ = objRelNFSE.Obter_DescServ(lCodMunic, sCodServ, sAux)
    Call Copia_String(sTexto, sAux)
     
    Exit Function
    
Erro_NFSe_Obter_DescServ:

    NFSe_Obter_DescServ = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209781)
     
    End Select
     
    Exit Function

End Function

Public Function Obter_Qtd_PedCompra(dQtd As Double, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Qtd_PedCompra

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Qtd_PedCompra = objRelProduto.Obter_Qtd_PedCompra(dQtd, sProduto)

    'Obter_Qtd_PedCompra = SUCESSO

    Exit Function

Erro_Obter_Qtd_PedCompra:

    Obter_Qtd_PedCompra = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196917)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_ReqCompra(dQtd As Double, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Qtd_ReqCompra

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Qtd_ReqCompra = objRelProduto.Obter_Qtd_ReqCompra(dQtd, sProduto)

    'Obter_Qtd_ReqCompra = SUCESSO

    Exit Function

Erro_Obter_Qtd_ReqCompra:

    Obter_Qtd_ReqCompra = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196917)

    End Select

    Exit Function

End Function


Public Function Obter_Qtd_PedCompraT(dQtd As Double, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Qtd_PedCompraT

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Qtd_PedCompraT = objRelProduto.Obter_Qtd_PedCompraT(dQtd, sProduto)
    
    'Obter_Qtd_PedCompraT = SUCESSO

    Exit Function

Erro_Obter_Qtd_PedCompraT:

    Obter_Qtd_PedCompraT = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196917)

    End Select

    Exit Function

End Function

Public Function Obter_Qtd_ReqCompraT(dQtd As Double, ByVal sProduto As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_Qtd_ReqCompraT

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_Qtd_ReqCompraT = objRelProduto.Obter_Qtd_ReqCompraT(dQtd, sProduto)
    
    'Obter_Qtd_ReqCompraT = SUCESSO

    Exit Function

Erro_Obter_Qtd_ReqCompraT:

    Obter_Qtd_ReqCompraT = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196917)

    End Select

    Exit Function

End Function

Public Function Obter_FatorConvUM(dFator As Double, ByVal iClasseUM As Integer, ByVal sSiglaUMDe As String, ByVal sSiglaUMPara As String) As Long

Dim lErro As Long
Dim sAux As String
Dim objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_FatorConvUM

    Set objRelProduto = ObtemObj("RelProduto")

    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto

        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_FatorConvUM = objRelProduto.Obter_FatorConvUM(dFator, iClasseUM, sSiglaUMDe, sSiglaUMPara)

    Exit Function

Erro_Obter_FatorConvUM:

    Obter_FatorConvUM = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201223)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterCodProd(sRetorno As String, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o codigo do Produto

Dim lErro As Long, objRelOS As ClassRelOS, sAux As String

On Error GoTo Erro_MovEstSRV_ObterCodProd

    Set objRelOS = ObtemObj("RelOS")
    
    If objRelOS Is Nothing Then

        Set objRelOS = New ClassRelOS
        
        lErro = GuardaObj("RelOS", objRelOS)
        If lErro <> SUCESSO Then gError 210557

    End If

    MovEstSRV_ObterCodProd = objRelOS.MovEstSRV_ObterCodProd(sAux, iFilialEmpresa, sCodigoOS, iLinha)
    
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MovEstSRV_ObterCodProd:

    MovEstSRV_ObterCodProd = gErr

    Select Case gErr

        Case 210557

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210558)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterDescProd(sRetorno As String, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório o Descricao do Produto

Dim lErro As Long, objRelOS As ClassRelOS, sAux As String

On Error GoTo Erro_MovEstSRV_ObterDescProd

    Set objRelOS = ObtemObj("RelOS")
    
    If objRelOS Is Nothing Then

        Set objRelOS = New ClassRelOS
        
        lErro = GuardaObj("RelOS", objRelOS)
        If lErro <> SUCESSO Then gError 210559

    End If

    MovEstSRV_ObterDescProd = objRelOS.MovEstSRV_ObterDescProd(sAux, iFilialEmpresa, sCodigoOS, iLinha)
    
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MovEstSRV_ObterDescProd:

    MovEstSRV_ObterDescProd = gErr

    Select Case gErr

        Case 210559

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210560)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterSiglaUMProd(sRetorno As String, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a UM do Produto

Dim lErro As Long, objRelOS As ClassRelOS, sAux As String

On Error GoTo Erro_MovEstSRV_ObterSiglaUMProd

    Set objRelOS = ObtemObj("RelOS")
    
    If objRelOS Is Nothing Then

        Set objRelOS = New ClassRelOS
        
        lErro = GuardaObj("RelOS", objRelOS)
        If lErro <> SUCESSO Then gError 210561

    End If

    MovEstSRV_ObterSiglaUMProd = objRelOS.MovEstSRV_ObterSiglaUMProd(sAux, iFilialEmpresa, sCodigoOS, iLinha)
    
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MovEstSRV_ObterSiglaUMProd:

    MovEstSRV_ObterSiglaUMProd = gErr

    Select Case gErr

        Case 210561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210562)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterQuantidade(dQtde As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve p/ o gerador de relatório a quantidade do produto

Dim lErro As Long, objRelOS As ClassRelOS

On Error GoTo Erro_MovEstSRV_ObterQuantidade

    Set objRelOS = ObtemObj("RelOS")
    
    If objRelOS Is Nothing Then

        Set objRelOS = New ClassRelOS
        
        lErro = GuardaObj("RelOS", objRelOS)
        If lErro <> SUCESSO Then gError 210563

    End If

    MovEstSRV_ObterQuantidade = objRelOS.MovEstSRV_ObterQuantidade(dQtde, iFilialEmpresa, sCodigoOS, iLinha)
    
    Exit Function

Erro_MovEstSRV_ObterQuantidade:

    MovEstSRV_ObterQuantidade = gErr

    Select Case gErr

        Case 210563

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210564)

    End Select

    Exit Function

End Function

Function MargContr_ObtemMargItemPV(dValor As Double, ByVal lNumIntItemPV As Long, ByVal dColunaSimulacao As Double, ByVal dLinhaPercMarg As Double, ByVal dColunaDVVSimulacao As Double) As Long

Dim lErro As Long, objRelMargContr2 As ClassRelMargContr2

On Error GoTo Erro_MargContr_ObtemMargItemPV

    Set objRelMargContr2 = ObtemObj("RelMargContr2")
    If objRelMargContr2 Is Nothing Then

        Set objRelMargContr2 = New ClassRelMargContr2
        lErro = GuardaObj("RelMargContr2", objRelMargContr2)
        If lErro <> SUCESSO Then gError 130040

    End If

    MargContr_ObtemMargItemPV = objRelMargContr2.MargContr_ObtemMargItemPV(dValor, lNumIntItemPV, CInt(dColunaSimulacao), CInt(dLinhaPercMarg), CInt(dColunaDVVSimulacao))
     
    Exit Function
    
Erro_MargContr_ObtemMargItemPV:

    MargContr_ObtemMargItemPV = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153038)
     
    End Select
     
    Exit Function

End Function

Function NF_Possui_ParcRec_Bordero(dFlag As Double, ByVal lNumIntNF As Long, ByVal lNumBordero As Long) As Long

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_NF_Possui_ParcRec_Bordero

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NF_Possui_ParcRec_Bordero = objRelNF.NF_Possui_ParcRec_Bordero(dFlag, lNumIntNF, lNumBordero)
     
    Exit Function
    
Erro_NF_Possui_ParcRec_Bordero:

    NF_Possui_ParcRec_Bordero = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153038)
     
    End Select
     
    Exit Function

End Function

Function ParcRec_Obtem_Bordero(dNumBordero As Double, ByVal lNumIntParc As Long) As Long

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_ParcRec_Obtem_Bordero

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    ParcRec_Obtem_Bordero = objRelNF.ParcRec_Obtem_Bordero(dNumBordero, lNumIntParc)
     
    Exit Function
    
Erro_ParcRec_Obtem_Bordero:

    ParcRec_Obtem_Bordero = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153038)
     
    End Select
     
    Exit Function

End Function

Public Function Obter_Usuario(sUsuario As String) As Long

Dim lErro As Long

On Error GoTo Erro_Obter_Usuario

    sUsuario = gsUsuario

    Obter_Usuario = SUCESSO

    Exit Function

Erro_Obter_Usuario:

    Obter_Usuario = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210627)

    End Select

    Exit Function

End Function

Public Function Obter_LigadoRegiaoVenda(dLigado As Double) As Long

Dim lErro As Long, iLigado As Integer
Dim objRelUsuario As ClassRelUsuario

On Error GoTo Erro_Obter_LigadoRegiaoVenda

    Set objRelUsuario = ObtemObj("RelUsuario")

    If objRelUsuario Is Nothing Then

        Set objRelUsuario = New ClassRelUsuario

        lErro = GuardaObj("RelUsuario", objRelUsuario)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_LigadoRegiaoVenda = objRelUsuario.Obter_LigadoRegiaoVenda(iLigado)

    dLigado = iLigado
    
    Exit Function

Erro_Obter_LigadoRegiaoVenda:

    Obter_LigadoRegiaoVenda = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210634)

    End Select

    Exit Function

End Function

Public Function Obter_PedidoCompra_Codigo(dCodigoPC As Double, ByVal sCodProduto As String, ByVal iFilialEmpresa As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long
Dim objRelContEstP As ClassRelContEstP

On Error GoTo Erro_Obter_PedidoCompra_Codigo

    Set objRelContEstP = ObtemObj("RelContEstP")

    If objRelContEstP Is Nothing Then

        Set objRelContEstP = New ClassRelContEstP

        lErro = GuardaObj("RelContEstP", objRelContEstP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_PedidoCompra_Codigo = objRelContEstP.Obter_PedidoCompra_Codigo(dCodigoPC, sCodProduto, iFilialEmpresa, iColuna)

    Exit Function

Erro_Obter_PedidoCompra_Codigo:

    Obter_PedidoCompra_Codigo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210760)

    End Select

    Exit Function

End Function


Public Function Obter_PedidoCompra_Qtde(dQtde As Double, ByVal sCodProduto As String, ByVal iFilialEmpresa As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long
Dim objRelContEstP As ClassRelContEstP

On Error GoTo Erro_Obter_PedidoCompra_Qtde

    Set objRelContEstP = ObtemObj("RelContEstP")

    If objRelContEstP Is Nothing Then

        Set objRelContEstP = New ClassRelContEstP

        lErro = GuardaObj("RelContEstP", objRelContEstP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_PedidoCompra_Qtde = objRelContEstP.Obter_PedidoCompra_Qtde(dQtde, sCodProduto, iFilialEmpresa, iColuna)

    Exit Function

Erro_Obter_PedidoCompra_Qtde:

    Obter_PedidoCompra_Qtde = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210761)

    End Select

    Exit Function

End Function

Public Function ObterDataParcRecNFe(dtData As Date, ByVal schNFe As String, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter a data de vencimento da parcela a receber

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_ObterDataParcRecNFe

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    ObterDataParcRecNFe = objRelNF.ObterDataParcRecNFe(dtData, schNFe, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterDataParcRecNFe:

    ObterDataParcRecNFe = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152833)

    End Select

    Exit Function

End Function

Public Function ObterValorParcRecNFe(dValor As Double, ByVal schNFe As String, ByVal lNumIntTitulo As Long, ByVal iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long, objRelNF As ClassRelNF

On Error GoTo Erro_ObterValorParcRecNFe

    Set objRelNF = ObtemObj("RelNF")
    If objRelNF Is Nothing Then

        Set objRelNF = New ClassRelNF
        lErro = GuardaObj("RelNF", objRelNF)
        If lErro <> SUCESSO Then gError 58222

    End If

    ObterValorParcRecNFe = objRelNF.ObterValorParcRecNFe(dValor, schNFe, lNumIntTitulo, iParcela)

    Exit Function

Erro_ObterValorParcRecNFe:

    ObterValorParcRecNFe = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152834)

    End Select

    Exit Function

End Function

Public Function Obter_InfAdicDocItem_Linha(sRetorno As String, ByVal iTipoDoc As Integer, ByVal lNumIntDocItem As Long, ByVal iTamMax As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long, objRelInfoAdicDocItem As ClassRelInfoAdicDocItem, sAux As String

On Error GoTo Erro_Obter_InfAdicDocItem_Linha

    Set objRelInfoAdicDocItem = ObtemObj("RelInfoAdicDocItem")
    If objRelInfoAdicDocItem Is Nothing Then

        Set objRelInfoAdicDocItem = New ClassRelInfoAdicDocItem
        lErro = GuardaObj("RelInfoAdicDocItem", objRelInfoAdicDocItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_InfAdicDocItem_Linha = objRelInfoAdicDocItem.ObterMsg(sAux, iTipoDoc, lNumIntDocItem, iTamMax, iLinha)
    Call Copia_String(sRetorno, sAux)
    
    Exit Function

Erro_Obter_InfAdicDocItem_Linha:

    Obter_InfAdicDocItem_Linha = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208094)

    End Select

    Exit Function

End Function

Public Function Obter_QtdeOVProdPer(dQtdeOV As Double, ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal lStatusMin As Long) As Long
'Para obter a qtde em ov por uma filialempresa de um produto num periodo

Dim lErro As Long, objRelProduto As ClassRelProduto

On Error GoTo Erro_Obter_QtdeOVProdPer

    Set objRelProduto = ObtemObj("RelProduto")
    If objRelProduto Is Nothing Then

        Set objRelProduto = New ClassRelProduto
        lErro = GuardaObj("RelProduto", objRelProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_QtdeOVProdPer = objRelProduto.Obter_QtdeOVProdPer(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, lStatusMin, dQtdeOV)

    Exit Function

Erro_Obter_QtdeOVProdPer:

    Obter_QtdeOVProdPer = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152793)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCFrete(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCFrete

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    MapaCot_ObterPCFrete = objRelMapaCot.MapaCot_ObterPCFrete(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCFrete:

    MapaCot_ObterPCFrete = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152803)

    End Select

    Exit Function

End Function

Public Function MapaCot_ObterPCFreteTipo(sRetorno As String, ByVal sProduto As String, ByVal iSequencial As Integer, ByVal iQtdMaxColecao As Integer) As Long

Dim lErro As Long, objRelMapaCot As ClassRelMapaCotacao, sAux As String

On Error GoTo Erro_MapaCot_ObterPCFreteTipo

    Set objRelMapaCot = ObtemObj("RelMapaCot")
    If objRelMapaCot Is Nothing Then

        Set objRelMapaCot = New ClassRelMapaCotacao
        lErro = GuardaObj("RelMapaCot", objRelMapaCot)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    MapaCot_ObterPCFreteTipo = objRelMapaCot.MapaCot_ObterPCFreteTipo(sAux, sProduto, iSequencial, iQtdMaxColecao)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MapaCot_ObterPCFreteTipo:

    MapaCot_ObterPCFreteTipo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152803)

    End Select

    Exit Function

End Function

Public Function Obter_PedidoCompra_Data(dtDataLimite As Date, ByVal sCodProduto As String, ByVal iFilialEmpresa As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long
Dim objRelContEstP As ClassRelContEstP

On Error GoTo Erro_Obter_PedidoCompra_Data

    Set objRelContEstP = ObtemObj("RelContEstP")

    If objRelContEstP Is Nothing Then

        Set objRelContEstP = New ClassRelContEstP

        lErro = GuardaObj("RelContEstP", objRelContEstP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_PedidoCompra_Data = objRelContEstP.Obter_PedidoCompra_Data(dtDataLimite, sCodProduto, iFilialEmpresa, iColuna)

    Exit Function

Erro_Obter_PedidoCompra_Data:

    Obter_PedidoCompra_Data = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210761)

    End Select

    Exit Function

End Function

Public Function PosCP_ObterCtaForn(sRetorno As String, ByVal iTipoLcto As Integer, ByVal lNumIntDoc As Long, ByVal iDocFilEmp As Integer) As Long

Dim lErro As Long, objRelPosCtb As ClassRelPosCtb, sAux As String

On Error GoTo Erro_PosCP_ObterCtaForn

    Set objRelPosCtb = ObtemObj("RelPosCtb")
    If objRelPosCtb Is Nothing Then

        Set objRelPosCtb = New ClassRelPosCtb
        lErro = GuardaObj("RelPosCtb", objRelPosCtb)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    PosCP_ObterCtaForn = objRelPosCtb.PosCPObterCtaForn(sAux, iTipoLcto, lNumIntDoc, iDocFilEmp)
    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_PosCP_ObterCtaForn:

    PosCP_ObterCtaForn = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152803)

    End Select

    Exit Function

End Function

Function MargContr_ObtemMargItemNF(dValor As Double, ByVal lNumIntItemNF As Long, ByVal dColunaSimulacao As Double, ByVal dLinhaPercMarg As Double, ByVal dColunaDVVSimulacao As Double) As Long

Dim lErro As Long, objRelMargContr3 As ClassRelMargContr3

On Error GoTo Erro_MargContr_ObtemMargItemNF

    Set objRelMargContr3 = ObtemObj("RelMargContr3")
    If objRelMargContr3 Is Nothing Then

        Set objRelMargContr3 = New ClassRelMargContr3
        lErro = GuardaObj("RelMargContr3", objRelMargContr3)
        If lErro <> SUCESSO Then gError 130040

    End If

    MargContr_ObtemMargItemNF = objRelMargContr3.MargContr_ObtemMargItemNF(dValor, lNumIntItemNF, CInt(dColunaSimulacao), CInt(dLinhaPercMarg), CInt(dColunaDVVSimulacao))
     
    Exit Function
    
Erro_MargContr_ObtemMargItemNF:

    MargContr_ObtemMargItemNF = gErr
     
    Select Case gErr
          
        Case 130040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153038)
     
    End Select
     
    Exit Function

End Function

Public Function ObterComissTextoDataVenc(sTextoDataVenc As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long) As Long

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sAux As String

On Error GoTo Erro_ObterComissTextoDataVenc

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then gError 132386

    End If

    ObterComissTextoDataVenc = objRelComissao.Comissao_Obter_TextoDataVenc(iTipoComiss, lNumIntComiss, sAux)

    Call Copia_String(sTextoDataVenc, sAux)

    Exit Function

Erro_ObterComissTextoDataVenc:

    ObterComissTextoDataVenc = gErr

    Select Case gErr

        Case 132386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function ObterComissEmBordero(dEmBordero As Double, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long) As Long

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim iEmBordero As Integer

On Error GoTo Erro_ObterComissEmBordero

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then gError 132386

    End If

    ObterComissEmBordero = objRelComissao.Comissao_Obter_EmBordero(iTipoComiss, lNumIntComiss, iEmBordero)
    
    dEmBordero = iEmBordero

    Exit Function

Erro_ObterComissEmBordero:

    ObterComissEmBordero = gErr

    Select Case gErr

        Case 132386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function ObterComissTextoBaixa(sTextoBaixa As String, ByVal iTipoComiss As Integer, ByVal lNumIntComiss As Long) As Long
'obtem texto associado à baixa do titulo ligado à comissao

Dim lErro As Long, objRelComissao As ClassRelComissao
Dim sTextoBaixaAux As String

On Error GoTo Erro_ObterComissTextoBaixa

    Set objRelComissao = ObtemObj("RelComissao")
    If objRelComissao Is Nothing Then

        Set objRelComissao = New ClassRelComissao
        lErro = GuardaObj("RelComissao", objRelComissao)
        If lErro <> SUCESSO Then gError 132386

    End If

    ObterComissTextoBaixa = objRelComissao.Comissao_Obter_TextoBaixa(iTipoComiss, lNumIntComiss, sTextoBaixaAux)
    Call Copia_String(sTextoBaixa, sTextoBaixaAux)

    Exit Function

Erro_ObterComissTextoBaixa:

    ObterComissTextoBaixa = gErr

    Select Case gErr

        Case 132386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function NFCE_ObterImagemQRCode(sArquivo As String, ByVal lNumIntNF As Long) As Long
'obtem nome do arquivo que contém a imagem do qrcode associado à nfce

Dim lErro As Long, objRelNFCE As ClassRelNFCE
Dim sArqAux As String

On Error GoTo Erro_NFCE_ObterImagemQRCode

    Set objRelNFCE = ObtemObj("RelNFCE")
    If objRelNFCE Is Nothing Then

        Set objRelNFCE = New ClassRelNFCE
        lErro = GuardaObj("RelNFCE", objRelNFCE)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NFCE_ObterImagemQRCode = objRelNFCE.ObterImagemQRCode(lNumIntNF, sArqAux)
    Call Copia_String(sArquivo, sArqAux)

    Exit Function

Erro_NFCE_ObterImagemQRCode:

    NFCE_ObterImagemQRCode = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function Obter_ProdTrib_ICMSSTValor(dValor As Double, ByVal iFilialEmpresa As Integer, ByVal lCliente As Long, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dPreco As Double) As Long

Dim lErro As Long, objRelProdTrib As ClassRelProdTrib

On Error GoTo Erro_Obter_ProdTrib_ICMSSTValor

    Set objRelProdTrib = ObtemObj("RelProdTrib")
    If objRelProdTrib Is Nothing Then

        Set objRelProdTrib = New ClassRelProdTrib
        lErro = GuardaObj("RelProdTrib", objRelProdTrib)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Obter_ProdTrib_ICMSSTValor = objRelProdTrib.Obter_ICMSSTValor(iFilialEmpresa, lCliente, iFilialCli, sProduto, dPreco, dValor)

    Exit Function

Erro_Obter_ProdTrib_ICMSSTValor:

    Obter_ProdTrib_ICMSSTValor = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213850)

    End Select

    Exit Function

End Function

Public Function MovCtaCo_Obter_CliForn(sRetorno As String, ByVal iTipo As Integer, ByVal lNumMovto As Long) As Long

Dim lErro As Long, objRelTes As ClassRelTes, sAux As String

On Error GoTo Erro_MovCtaCo_Obter_CliForn

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    MovCtaCo_Obter_CliForn = objRelTes.Obter_CliForn(iTipo, lNumMovto, sAux)

    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MovCtaCo_Obter_CliForn:

    MovCtaCo_Obter_CliForn = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213851)

    End Select

    Exit Function

End Function

Public Function MovCtaCo_Obter_HistoricoTit(sRetorno As String, ByVal iTipo As Integer, ByVal lNumMovto As Long) As Long

Dim lErro As Long, objRelTes As ClassRelTes, sAux As String

On Error GoTo Erro_MovCtaCo_Obter_HistoricoTit

    Set objRelTes = ObtemObj("RelTes")
    If objRelTes Is Nothing Then

        Set objRelTes = New ClassRelTes
        lErro = GuardaObj("RelTes", objRelTes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    MovCtaCo_Obter_HistoricoTit = objRelTes.Obter_HistoricoTit(iTipo, lNumMovto, sAux)

    Call Copia_String(sRetorno, sAux)

    Exit Function

Erro_MovCtaCo_Obter_HistoricoTit:

    MovCtaCo_Obter_HistoricoTit = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213852)

    End Select

    Exit Function

End Function

Public Function RelProdRanking_Obter_ValorFrete(dValorFrete As Double, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal lNumIntRel As Long) As Long
'Para obter o valor frete de uma filialempresa em um periodo

Dim lErro As Long, objRelRankingProd As ClassRelRankingProd

On Error GoTo Erro_RelProdRanking_Obter_ValorFrete

    Set objRelRankingProd = ObtemObj("RelRankingProd")
    If objRelRankingProd Is Nothing Then

        Set objRelRankingProd = New ClassRelRankingProd
        lErro = GuardaObj("RelRankingProd", objRelRankingProd)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    RelProdRanking_Obter_ValorFrete = objRelRankingProd.RelProdRanking_Obter_ValorFrete(dValorFrete, iFilialEmpresa, dtDataDe, dtDataAte, lNumIntRel)

    Exit Function

Erro_RelProdRanking_Obter_ValorFrete:

    RelProdRanking_Obter_ValorFrete = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152929)

    End Select

    Exit Function

End Function

Public Function CF_Obter_ValorFrete(dValorFrete As Double, ByVal iFilialEmpresa As Integer, ByVal iCaixa As Integer, ByVal dtData As Date) As Long
'Para obter o valor frete de uma filialempresa em um periodo

Dim lErro As Long, objRelRankingProd As ClassRelRankingProd

On Error GoTo Erro_CF_Obter_ValorFrete

    Set objRelRankingProd = ObtemObj("RelRankingProd")
    If objRelRankingProd Is Nothing Then

        Set objRelRankingProd = New ClassRelRankingProd
        lErro = GuardaObj("RelRankingProd", objRelRankingProd)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    CF_Obter_ValorFrete = objRelRankingProd.CF_Obter_ValorFrete(dValorFrete, iFilialEmpresa, iCaixa, dtData)

    Exit Function

Erro_CF_Obter_ValorFrete:

    CF_Obter_ValorFrete = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152929)

    End Select

    Exit Function

End Function

