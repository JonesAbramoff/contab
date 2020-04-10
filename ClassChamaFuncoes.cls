VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'Function EstoqueMes_Le(objEstoqueMes As ClassEstoqueMes) As Long
''Le as informacoes sobre o EstoqueMes através do parametro Filial/Ano/Mes passado em objEstoqueMes
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_EstoqueMes_Le
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE", X)
'    If lErro <> SUCESSO Then Error 36514
'
'    EstoqueMes_Le = X.EstoqueMes_Le(objEstoqueMes)
'
'    Exit Function
'
'Erro_EstoqueMes_Le:
'
'    EstoqueMes_Le = Err
'
'    Select Case Err
'
'        Case 36514
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144669)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Conta_Critica(ByVal sConta As String, sContaFormatada As String, objPlanoConta As ClassPlanoConta, sSiglaModulo As String) As Long
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Critica
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_CRITICA", X)
'    If lErro <> SUCESSO Then Error 7001
'
'    Conta_Critica = X.Conta_Critica(sConta, sContaFormatada, objPlanoConta, sSiglaModulo)
'
'    Exit Function
'
'Erro_Conta_Critica:
'
'    Conta_Critica = Err
'
'    Select Case Err
'
'        Case 7001
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144670)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Critica(ByVal iExercicio As Integer, iExercicioFechado As Integer) As Long
''Verifica se o exercicio está fechado
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Exercicio_Critica
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_CRITICA", X)
'    If lErro <> SUCESSO Then Error 7003
'
'
'    Exercicio_Critica = X.Exercicio_Critica(iExercicio, iExercicioFechado)
'    Exit Function
'
'Erro_Exercicio_Critica:
'
'    Exercicio_Critica = Err
'
'    Select Case Err
'
'        Case 7003
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144671)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lote_Critica_Atualizado(tLote As ClassLote, iLoteAtualizado As Integer) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Lote_Critica_Atualizado
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_CRITICA_ATUALIZADO", X)
'    If lErro <> SUCESSO Then Error 7004
'
'
'    Lote_Critica_Atualizado = X.Lote_Critica_Atualizado(tLote, iLoteAtualizado)
'    Exit Function
'
'Erro_Lote_Critica_Atualizado:
'    Lote_Critica_Atualizado = Err
'    Select Case Err
'
'        Case 7004
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144672)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function ContaAnalitica_Critica_Exclusao(sConta As String) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_ContaAnalitica_Critica_Exclusao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTAANALITICA_CRITICA_EXCLUSAO", X)
'    If lErro <> SUCESSO Then Error 7005
'
'
'    ContaAnalitica_Critica_Exclusao = X.ContaAnalitica_Critica_Exclusao(sConta)
'
'    Exit Function
'
'Erro_ContaAnalitica_Critica_Exclusao:
'    ContaAnalitica_Critica_Exclusao = Err
'    Select Case Err
'
'        Case 7005
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144673)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Conta_Critica_Possui_Movimento(sConta As String, iTemMovimento As Integer) As Long
''verifica se a conta possui movimentação nos exercicios existentes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Critica_Possui_Movimento
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_CRITICA_POSSUI_MOVIMENTO", X)
'    If lErro <> SUCESSO Then Error 5656
'
'
'    Conta_Critica_Possui_Movimento = X.Conta_Critica_Possui_Movimento(sConta, iTemMovimento)
'
'    Exit Function
'
'Erro_Conta_Critica_Possui_Movimento:
'
'    Conta_Critica_Possui_Movimento = Err
'
'    Select Case Err
'
'        Case 5656
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144674)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Critica_Exclusao(sCcl As String) As Long
''verifica se o Ccl possui rateios associados, etc.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica_Exclusao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA_EXCLUSAO", X)
'    If lErro <> SUCESSO Then Error 5679
'
'
'    Ccl_Critica_Exclusao = X.Ccl_Critica_Exclusao(sCcl)
'
'    Exit Function
'
'Erro_Ccl_Critica_Exclusao:
'
'    Ccl_Critica_Exclusao = Err
'
'    Select Case Err
'
'        Case 5679
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144675)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Critica_Possui_Movimento(sCcl As String, iTemMovimento As Integer) As Long
''verifica se o Ccl possui movimentação nos exercicios existentes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica_Possui_Movimento
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA_POSSUI_MOVIMENTO", X)
'    If lErro <> SUCESSO Then Error 5680
'
'
'    Ccl_Critica_Possui_Movimento = X.Ccl_Critica_Possui_Movimento(sCcl, iTemMovimento)
'
'    Exit Function
'
'Erro_Ccl_Critica_Possui_Movimento:
'
'    Ccl_Critica_Possui_Movimento = Err
'
'    Select Case Err
'
'        Case 5680
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144676)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Critica_Possui_LanPendente(sCcl As String, iTemLanPendente As Integer) As Long
''Verifica se existem lançamentos pendentes para o centro de custo em questão. Se existir retorna iTemLanPendente = 1, senao iTemLanPendente = 0
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica_Possui_LanPendente
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA_POSSUI_LANPENDENTE", X)
'    If lErro <> SUCESSO Then Error 5681
'
'
'    Ccl_Critica_Possui_LanPendente = X.Ccl_Critica_Possui_LanPendente(sCcl, iTemLanPendente)
'
'    Exit Function
'
'Erro_Ccl_Critica_Possui_LanPendente:
'
'    Ccl_Critica_Possui_LanPendente = Err
'
'    Select Case Err
'
'        Case 5681
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144677)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Conta_Critica_Possui_LanPendente(sConta As String, iTemLanPendente As Integer) As Long
''Verifica se existem lançamentos pendentes para a conta em questão. Se existir retorna iTemLanPendente = 1, senao iTemLanPendente = 0
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Critica_Possui_LanPendente
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_CRITICA_POSSUI_LANPENDENTE", X)
'    If lErro <> SUCESSO Then Error 5682
'
'
'    Conta_Critica_Possui_LanPendente = X.Conta_Critica_Possui_LanPendente(sConta, iTemLanPendente)
'
'    Exit Function
'
'Erro_Conta_Critica_Possui_LanPendente:
'
'    Conta_Critica_Possui_LanPendente = Err
'
'    Select Case Err
'
'        Case 5682
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144678)
'
'
'    End Select
'
'    Exit Function
'
'
'End Function
'
'Function Ccl_Critica(ByVal sCcl As String, sCclFormatada As String, objCcl As ClassCcl) As Long
''le o centro de custo/lucro do banco de dados se estiver cadastrado.
'
'    Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA", X)
'    If lErro <> SUCESSO Then Error 5704
'
'
'    Ccl_Critica = X.Ccl_Critica(sCcl, sCclFormatada, objCcl)
'
'    Exit Function
'
'Erro_Ccl_Critica:
'
'    Ccl_Critica = Err
'
'    Select Case Err
'
'        Case 5704
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144679)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function LanPendente_Critica_TotaisLote(objLote As ClassLote, iTotaisIguais As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Critica_TotaisLote
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_CRITICA_TOTAISLOTE", X)
'    If lErro <> SUCESSO Then Error 14003
'
'
'    LanPendente_Critica_TotaisLote = X.LanPendente_Critica_TotaisLote(objLote, iTotaisIguais)
'
'    Exit Function
'
'Erro_LanPendente_Critica_TotaisLote:
'
'    LanPendente_Critica_TotaisLote = Err
'
'    Select Case Err
'
'        Case 14003
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144680)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCcl_Critica_Possui_Movimento(sConta As String, sCcl As String, iTemMovimento As Integer) As Long
''verifica se a assoc possui movimentação nos exercicios existentes.
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Critica_Possui_Movimento
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_CRITICA_POSSUI_MOVIMENTO", X)
'    If lErro <> SUCESSO Then Error 14009
'
'
'    ContaCcl_Critica_Possui_Movimento = X.ContaCcl_Critica_Possui_Movimento(sConta, sCcl, iTemMovimento)
'
'    Exit Function
'
'Erro_ContaCcl_Critica_Possui_Movimento:
'
'    ContaCcl_Critica_Possui_Movimento = Err
'
'    Select Case Err
'
'        Case 14009
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144681)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Critica_Colecao_ExtraContab(ByVal sConta As String, colCcl As Collection) As Long
''Verifica se os Ccl's associados a sConta estao marcados na lisbox
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Critica_Colecao_ExtraContab
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_CRITICA_COLECAO_EXTRACONTAB", X)
'    If lErro <> SUCESSO Then Error 14011
'
'
'    ContaCcl_Critica_Colecao_ExtraContab = X.ContaCcl_Critica_Colecao_ExtraContab(sConta, colCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Critica_Colecao_ExtraContab:
'
'    ContaCcl_Critica_Colecao_ExtraContab = Err
'
'    Select Case Err
'
'        Case 14011
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144682)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Critica_Colecao_Contabil(ByVal sConta As String, colCcl As Collection) As Long
''Verifica se os Ccl's associados a sConta estao marcados na lisbox
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Critica_Colecao_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_CRITICA_COLECAO_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14032
'
'
'    ContaCcl_Critica_Colecao_Contabil = X.ContaCcl_Critica_Colecao_Contabil(sConta, colCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Critica_Colecao_Contabil:
'
'    ContaCcl_Critica_Colecao_Contabil = Err
'
'    Select Case Err
'
'        Case 14032
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144683)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Saldo_Inicial_Critica(iDisponivel As Integer) As Long
''verifica se o saldo inicial pode ser alterado
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Saldo_Inicial_Critica
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SALDO_INICIAL_CRITICA", X)
'    If lErro <> SUCESSO Then Error 5971
'
'
'    Saldo_Inicial_Critica = X.Saldo_Inicial_Critica(iDisponivel)
'
'    Exit Function
'
'Erro_Saldo_Inicial_Critica:
'
'    Saldo_Inicial_Critica = Err
'
'    Select Case Err
'
'        Case 5971
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144684)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Conta_Critica_ContaPai(sConta As String, sSiglaModulo As String) As Long
''verifica se a conta tem um pai sintetico
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Critica_ContaPai
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_CRITICA_CONTAPAI", X)
'    If lErro <> SUCESSO Then Error 9620
'
'
'    Conta_Critica_ContaPai = X.Conta_Critica_ContaPai(sConta, sSiglaModulo)
'
'    Exit Function
'
'Erro_Conta_Critica_ContaPai:
'
'    Conta_Critica_ContaPai = Err
'
'    Select Case Err
'
'        Case 9620
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144685)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Conta_Critica_Categoria(sConta As String, iCategoria As Integer) As Long
''verifica se a conta é de nivel 1 e tem a categoria preenchida.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Critica_Categoria
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_CRITICA_CATEGORIA", X)
'    If lErro <> SUCESSO Then Error 9719
'
'
'    Conta_Critica_Categoria = X.Conta_Critica_Categoria(sConta, iCategoria)
'
'    Exit Function
'
'Erro_Conta_Critica_Categoria:
'
'    Conta_Critica_Categoria = Err
'
'    Select Case Err
'
'        Case 9719
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144686)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Formata(sCcl As String, sCclFormata As String, iCclPreenchida As Integer) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Ccl_Formata
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_FORMATA", X)
'    If lErro <> SUCESSO Then Error 6123
'
'
'    Ccl_Formata = X.Ccl_Formata(sCcl, sCclFormata, iCclPreenchida)
'    Exit Function
'
'Erro_Ccl_Formata:
'
'    Ccl_Formata = Err
'
'    Select Case Err
'
'        Case 6123
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144687)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Conta_Formata(sConta As String, sContaFormata As String, iContaPreenchida As Integer) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Conta_Formata
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_FORMATA", X)
'    If lErro <> SUCESSO Then Error 7006
'
'
'    Conta_Formata = X.Conta_Formata(sConta, sContaFormata, iContaPreenchida)
'    Exit Function
'
'Erro_Conta_Formata:
'    Conta_Formata = Err
'    Select Case Err
'
'        Case 7006
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144688)
'
'
'    End Select
'    Exit Function
'End Function
'
''Devolve a conta no Formato do BD sem centro de custo: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
''A String de Entrada vem do GRID com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'Function Conta_Formata_SemCcl(sConta As String, sContaFormatada As String, iContaPreenchida As Integer) As Long
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Conta_Formata_SemCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_FORMATA_SEMCCL", X)
'    If lErro <> SUCESSO Then Error 5859
'
'
'    Conta_Formata_SemCcl = X.Conta_Formata_SemCcl(sConta, sContaFormatada, iContaPreenchida)
'    Exit Function
'
'Erro_Conta_Formata_SemCcl:
'    Conta_Formata_SemCcl = Err
'    Select Case Err
'
'        Case 5859
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144689)
'
'
'    End Select
'    Exit Function
'
'
'
'End Function
'
'Function Str_Limpa(sString As String, sChar As String) As String
'End Function
'
'Function Str_Subst(sString As String, sChar1 As String, sChar2 As String) As String
'End Function
'
'Function LotePendente_Grava(objLote As ClassLote) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_LotePendente_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_GRAVA", X)
'    If lErro <> SUCESSO Then Error 7007
'
'
'    LotePendente_Grava = X.LotePendente_Grava(objLote)
'    Exit Function
'
'Erro_LotePendente_Grava:
'    LotePendente_Grava = Err
'    Select Case Err
'
'        Case 7007
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144690)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function LotePendente_Grava_Totais(objLote As ClassLote) As Long
'
'Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_LotePendente_Grava_Totais
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_GRAVA_TOTAIS", X)
'    If lErro <> SUCESSO Then Error 14002
'
'
'    LotePendente_Grava_Totais = X.LotePendente_Grava_Totais(objLote)
'    Exit Function
'
'Erro_LotePendente_Grava_Totais:
'
'    LotePendente_Grava_Totais = Err
'    Select Case Err
'
'        Case 14002
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144691)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function LotePendente_Exclui(tLote As ClassLote) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_LotePendente_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 7008
'
'
'    LotePendente_Exclui = X.LotePendente_Exclui(tLote)
'    Exit Function
'
'Erro_LotePendente_Exclui:
'    LotePendente_Exclui = Err
'    Select Case Err
'
'        Case 7008
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144692)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Inclui_Associacao_ExtraCont(ByVal sConta As String, ByVal sCcl As String) As Long
''Inclui no BD as associacoes de sConta com cada Ccl marcado
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
' Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_ContaCcl_Inclui_Associacao_ExtraCont
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_INCLUI_ASSOCIACAO_EXTRACONT", X)
'    If lErro <> SUCESSO Then Error 14012
'
'
'    ContaCcl_Inclui_Associacao_ExtraCont = X.ContaCcl_Inclui_Associacao_ExtraCont(sConta, sCcl)
'    Exit Function
'
'Erro_ContaCcl_Inclui_Associacao_ExtraCont:
'
'    ContaCcl_Inclui_Associacao_ExtraCont = Err
'
'    Select Case Err
'
'        Case 14012
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144693)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Inclui_Associacao_Contabil(ByVal sConta As String, ByVal sCcl As String) As Long
''Inclui no BD as associacoes de sConta com cada Ccl marcado
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
' Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_ContaCcl_Inclui_Associacao_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_INCLUI_ASSOCIACAO_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14031
'
'
'    ContaCcl_Inclui_Associacao_Contabil = X.ContaCcl_Inclui_Associacao_Contabil(sConta, sCcl)
'    Exit Function
'
'Erro_ContaCcl_Inclui_Associacao_Contabil:
'
'    ContaCcl_Inclui_Associacao_Contabil = Err
'
'    Select Case Err
'
'        Case 14031
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144694)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'Function PlanoConta_Insere(objPlanoConta As ClassPlanoConta) As Long
''insere uma conta no banco de dados
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_PlanoConta_Insere
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_INSERE", X)
'    If lErro <> SUCESSO Then Error 7009
'
'
'    PlanoConta_Insere = X.PlanoConta_Insere(objPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Insere:
'    PlanoConta_Insere = Err
'    Select Case Err
'
'        Case 7009
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144695)
'
'
'    End Select
'    Exit Function
'
'End Function
'
'Function PlanoConta_Altera(objPlanoConta As ClassPlanoConta) As Long
''altera uma conta no banco de dados
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_PlanoConta_Altera
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_ALTERA", X)
'    If lErro <> SUCESSO Then Error 7010
'
'
'    PlanoConta_Altera = X.PlanoConta_Altera(objPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Altera:
'    PlanoConta_Altera = Err
'    Select Case Err
'
'        Case 7010
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144696)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Conta_Exclui(sConta As String) As Long
''exclui uma conta e seus descendentes do banco de dados
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Conta_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 7011
'
'
'    Conta_Exclui = X.Conta_Exclui(sConta)
'
'    Exit Function
'
'Erro_Conta_Exclui:
'    Conta_Exclui = Err
'    Select Case Err
'
'        Case 7011
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144697)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function DocAuto_Exclui_Conta(sConta As String) As Long
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_DocAuto_Exclui_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_EXCLUI_CONTA", X)
'    If lErro <> SUCESSO Then Error 7012
'
'
'    DocAuto_Exclui_Conta = X.DocAuto_Exclui_Conta(sConta)
'
'    Exit Function
'
'Erro_DocAuto_Exclui_Conta:
'    DocAuto_Exclui_Conta = Err
'    Select Case Err
'
'        Case 7012
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144698)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Lancamento_Exclui(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
''Exclui Documento Pendente do BD
'
'Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Lancamento_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 7013
'
'
'    Lancamento_Exclui = X.Lancamento_Exclui(objLancamento_Cabecalho)
'
'    Exit Function
'
'Erro_Lancamento_Exclui:
'
'    Lancamento_Exclui = Err
'
'    Select Case Err
'
'        Case 7013
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144699)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Exclui_BD(objLancamento_Cabecalho As ClassLancamento_Cabecalho, dCredito As Double, dDebito As Double, iNumLancAtual As Integer, iNumDocAtual As Integer)
''exclui da tabela de lançamentos pendentes os lançamentos do voucher em questão
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Exclui_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_EXCLUI_BD", X)
'    If lErro <> SUCESSO Then Error 14310
'
'
'    Lancamento_Exclui_BD = X.Lancamento_Exclui_BD(objLancamento_Cabecalho, dCredito, dDebito, iNumLancAtual, iNumDocAtual)
'
'    Exit Function
'
'Erro_Lancamento_Exclui_BD:
'
'    Lancamento_Exclui_BD = Err
'
'    Select Case Err
'
'        Case 14310
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144700)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Grava Lançamemtos Pendentes no BD
'Function Lancamento_Grava(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_GRAVA", X)
'    If lErro <> SUCESSO Then Error 7014
'
'
'    Lancamento_Grava = X.Lancamento_Grava(objLancamento_Cabecalho, colLancamento_Detalhe)
'
'    Exit Function
'
'Erro_Lancamento_Grava:
'
'    Lancamento_Grava = Err
'
'    Select Case Err
'
'        Case 7014
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144701)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Grava_BD(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection, dCredito As Double, dDebito As Double, iNumLancAtual As Integer) As Long
''grava na tabela de lançamentos pendentes os lançamentos do voucher em questão
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Grava_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_GRAVA_BD", X)
'    If lErro <> SUCESSO Then Error 14311
'
'
'    Lancamento_Grava_BD = X.Lancamento_Grava_BD(objLancamento_Cabecalho, colLancamento_Detalhe, dCredito, dDebito, iNumLancAtual)
'
'    Exit Function
'
'Erro_Lancamento_Grava_BD:
'
'    Lancamento_Grava_BD = Err
'
'    Select Case Err
'
'        Case 14311
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144702)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Atualiza(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Atualiza
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_ATUALIZA", X)
'    If lErro <> SUCESSO Then Error 14308
'
'
'    Lancamento_Atualiza = X.Lancamento_Atualiza(objLancamento_Cabecalho, colLancamento_Detalhe)
'
'    Exit Function
'
'Erro_Lancamento_Atualiza:
'
'    Lancamento_Atualiza = Err
'
'    Select Case Err
'
'        Case 14308
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144703)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Atualiza_BD(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection, dCredito As Double, dDebito As Double, iNumLancAtual As Integer) As Long
''grava na tabela de lançamentos pendentes os lançamentos do voucher em questão
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Atualiza_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_ATUALIZA_BD", X)
'    If lErro <> SUCESSO Then Error 14309
'
'
'    Lancamento_Atualiza_BD = X.Lancamento_Atualiza_BD(objLancamento_Cabecalho, colLancamento_Detalhe, dCredito, dDebito, iNumLancAtual)
'
'    Exit Function
'
'Erro_Lancamento_Atualiza_BD:
'
'    Lancamento_Atualiza_BD = Err
'
'    Select Case Err
'
'        Case 14309
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144704)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Exclui(sCcl As String) As Long
''exclui o ccl (rotina de mais alto nível)
'
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Ccl_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 7015
'
'
'    Ccl_Exclui = X.Ccl_Exclui(sCcl)
'
'    Exit Function
'
'Erro_Ccl_Exclui:
'    Ccl_Exclui = Err
'    Select Case Err
'
'        Case 7015
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144705)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Ccl_Grava(objCcl As ClassCcl) As Long
''grava o centro de custo/lucro no banco de dados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_GRAVA", X)
'    If lErro <> SUCESSO Then Error 7016
'
'
'    Ccl_Grava = X.Ccl_Grava(objCcl)
'
'    Exit Function
'
'Erro_Ccl_Grava:
'
'    Ccl_Grava = Err
'
'    Select Case Err
'
'        Case 7016
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144706)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Altera(objCcl As ClassCcl) As Long
''altera o centro de custo/lucro no banco de dados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Altera
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_ALTERA", X)
'    If lErro <> SUCESSO Then Error 9609
'
'
'    Ccl_Altera = X.Ccl_Altera(objCcl)
'
'    Exit Function
'
'Erro_Ccl_Altera:
'
'    Ccl_Altera = Err
'
'    Select Case Err
'
'        Case 9609
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144707)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Insere(objCcl As ClassCcl) As Long
''altera o centro de custo/lucro no banco de dados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Insere
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_INSERE", X)
'    If lErro <> SUCESSO Then Error 9610
'
'
'    Ccl_Insere = X.Ccl_Insere(objCcl)
'
'    Exit Function
'
'Erro_Ccl_Insere:
'
'    Ccl_Insere = Err
'
'    Select Case Err
'
'        Case 9610
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144708)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCta_Exclui_Conta(sConta As String) As Long
''exclui os saldos de conta associados a conta em questão
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCta_Exclui_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCTA_EXCLUI_CONTA", X)
'    If lErro <> SUCESSO Then Error 9167
'
'
'    MvPerCta_Exclui_Conta = X.MvPerCta_Exclui_Conta(sConta)
'
'    Exit Function
'
'Erro_MvPerCta_Exclui_Conta:
'
'    MvPerCta_Exclui_Conta = Err
'
'    Select Case Err
'
'        Case 9167
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144709)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCcl_Exclui_Conta(sConta As String) As Long
''exclui os saldos de ccl associados a conta em questão
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCcl_Exclui_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCCL_EXCLUI_CONTA", X)
'    If lErro <> SUCESSO Then Error 9166
'
'
'    MvPerCcl_Exclui_Conta = X.MvPerCcl_Exclui_Conta(sConta)
'
'    Exit Function
'
'Erro_MvPerCcl_Exclui_Conta:
'
'    MvPerCcl_Exclui_Conta = Err
'
'    Select Case Err
'
'        Case 9166
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144710)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Exclui_Ccl(sCcl As String) As Long
''exclui os lançamentos de documento automático associados o Ccl em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Exclui_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_EXCLUI_CCL", X)
'    If lErro <> SUCESSO Then Error 5676
'
'
'    DocAuto_Exclui_Ccl = X.DocAuto_Exclui_Ccl(sCcl)
'
'    Exit Function
'
'Erro_DocAuto_Exclui_Ccl:
'
'    DocAuto_Exclui_Ccl = Err
'
'    Select Case Err
'
'        Case 5676
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144711)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCcl_Exclui_Ccl(sCcl As String) As Long
''exclui os saldos de ccl associados
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCcl_Exclui_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCCL_EXCLUI_CCL", X)
'    If lErro <> SUCESSO Then Error 5677
'
'
'    MvPerCcl_Exclui_Ccl = X.MvPerCcl_Exclui_Ccl(sCcl)
'
'    Exit Function
'
'Erro_MvPerCcl_Exclui_Ccl:
'
'    MvPerCcl_Exclui_Ccl = Err
'
'    Select Case Err
'
'        Case 5677
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144712)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Exclui_Conta(sCcl As String) As Long
''Exclui as contas associadas ao centro de custo e os respectivos relacionamentos.
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Exclui_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_EXCLUI_CONTA", X)
'    If lErro <> SUCESSO Then Error 5678
'
'
'    Ccl_Exclui_Conta = X.Ccl_Exclui_Conta(sCcl)
'
'    Exit Function
'
'Erro_Ccl_Exclui_Conta:
'
'    Ccl_Exclui_Conta = Err
'
'    Select Case Err
'
'        Case 5678
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144713)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Ccl_Exclui_Sem_Assoc_ExtraContabil(ByVal sConta As String, ByVal sCcl As String) As Long
''Exclui do BD os Ccl's associados a sConta que nao estejam marcados na listbox
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Exclui_Sem_Assoc_ExtraContabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_EXCLUI_SEM_ASSOC_EXTRACONTABIL", X)
'    If lErro <> SUCESSO Then Error 14010
'
'
'    Ccl_Exclui_Sem_Assoc_ExtraContabil = X.Ccl_Exclui_Sem_Assoc_ExtraContabil(sConta, sCcl)
'
'    Exit Function
'
'Erro_Ccl_Exclui_Sem_Assoc_ExtraContabil:
'
'    Ccl_Exclui_Sem_Assoc_ExtraContabil = Err
'
'    Select Case Err
'
'        Case 14010
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144714)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Ccl_Exclui_Sem_Assoc_Contabil(ByVal sConta As String, ByVal sCcl As String) As Long
''Exclui do BD os Ccl's associados a sConta que nao estejam marcados na listbox
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Exclui_Sem_Assoc_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_EXCLUI_SEM_ASSOC_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14028
'
'
'    Ccl_Exclui_Sem_Assoc_Contabil = X.Ccl_Exclui_Sem_Assoc_Contabil(sConta, sCcl)
'
'    Exit Function
'
'Erro_Ccl_Exclui_Sem_Assoc_Contabil:
'
'    Ccl_Exclui_Sem_Assoc_Contabil = Err
'
'    Select Case Err
'
'        Case 14028
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144715)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'Public Function ContaCcl_Exclui_Assoc(sConta As String, sCcl As String) As Long
''Exclui do BD as associacoes de sConta com sCcl
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_Assoc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_ASSOC", X)
'    If lErro <> SUCESSO Then Error 14019
'
'
'    ContaCcl_Exclui_Assoc = X.ContaCcl_Exclui_Assoc(sConta, sCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_Assoc:
'
'    ContaCcl_Exclui_Assoc = Err
'
'    Select Case Err
'
'        Case 14019
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144716)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function MvPerCcl_Exclui_ContaCcl(sConta As String, sCcl As String) As Long
''Exclui do BD os saldos associados a sConta+sCcl
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCcl_Exclui_ContaCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCCL_EXCLUI_CONTACCL", X)
'    If lErro <> SUCESSO Then Error 14020
'
'
'    MvPerCcl_Exclui_ContaCcl = X.MvPerCcl_Exclui_ContaCcl(sConta, sCcl)
'
'    Exit Function
'
'Erro_MvPerCcl_Exclui_ContaCcl:
'
'    MvPerCcl_Exclui_ContaCcl = Err
'
'    Select Case Err
'
'        Case 14020
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144717)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function DocAuto_Exclui_ContaCcl(sConta As String, sCcl As String) As Long
''Exclui do BD os documentos automaticos associados a sConta+sCcl
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Exclui_ContaCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_EXCLUI_CONTACCL", X)
'    If lErro <> SUCESSO Then Error 14021
'
'
'    DocAuto_Exclui_ContaCcl = X.DocAuto_Exclui_ContaCcl(sConta, sCcl)
'
'    Exit Function
'
'Erro_DocAuto_Exclui_ContaCcl:
'
'    DocAuto_Exclui_ContaCcl = Err
'
'    Select Case Err
'
'        Case 14021
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144718)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Public Function Exclui_RateioCab_ContaCcl(sConta As String, sCcl As String) As Long
''Exclui do BD os rateios cab associados a sConta+sCcl
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
''Dim X As Object, lErro As Long
'
''On Error GoTo Erro_Exclui_RateioCab_ContaCcl
'
''
''
'
''    lErro = Rotina_ObterObjetoFuncao("CTB_EXCLUI_RATEIOCAB_CONTACCL",X)
''    If lErro <> SUCESSO Then Error 14022
'
''
''    Exclui_RateioCab_ContaCcl = X.Exclui_RateioCab_ContaCcl(sConta, sCcl)
'
''    Exit Function
'
''Erro_Exclui_RateioCab_ContaCcl:
'
''   Exclui_RateioCab_ContaCcl = Err
'
''    Select Case Err
'
''        Case 14022
''
'
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144719)
'
''    End Select
'
''    Exit Function
'
''End Function
'
'Public Function ContaCcl_Exclui_RateioOn_ExtraContab(sConta As String, sCcl As String) As Long
''Exclui do BD os rateios on associados a sConta+sCcl
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_RateioOn_ExtraContab
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_RATEIOON_EXTRACONTAB", X)
'    If lErro <> SUCESSO Then Error 14023
'
'
'    ContaCcl_Exclui_RateioOn_ExtraContab = X.ContaCcl_Exclui_RateioOn_ExtraContab(sConta, sCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_RateioOn_ExtraContab:
'
'    ContaCcl_Exclui_RateioOn_ExtraContab = Err
'
'    Select Case Err
'
'        Case 14023
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144720)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Exclui_RateioOn_Contabil(sConta As String) As Long
''Exclui do BD os rateios on associados a sConta
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_RateioOn_Contabil
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_RATEIOON_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14030
'
'
'    ContaCcl_Exclui_RateioOn_Contabil = X.ContaCcl_Exclui_RateioOn_Contabil(sConta)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_RateioOn_Contabil:
'
'    ContaCcl_Exclui_RateioOn_Contabil = Err
'
'    Select Case Err
'
'        Case 14030
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144721)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Exclui_RateioOff_ExtraConta(sConta As String, sCcl As String) As Long
''Exclui do BD os rateios off associados a sConta+sCcl
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_RateioOff_ExtraConta
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_RATEIOOFF_EXTRACONTA", X)
'    If lErro <> SUCESSO Then Error 14024
'
'    ContaCcl_Exclui_RateioOff_ExtraConta = X.ContaCcl_Exclui_RateioOff_ExtraConta(sConta, sCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_RateioOff_ExtraConta:
'
'    ContaCcl_Exclui_RateioOff_ExtraConta = Err
'
'    Select Case Err
'
'        Case 14024
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144722)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Exclui_RateioOff_Contabil(sConta As String) As Long
''Exclui do BD os rateios off associados a sConta
''PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_RateioOff_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_RATEIOOFF_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14029
'
'
'    ContaCcl_Exclui_RateioOff_Contabil = X.ContaCcl_Exclui_RateioOff_Contabil(sConta)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_RateioOff_Contabil:
'
'    ContaCcl_Exclui_RateioOff_Contabil = Err
'
'    Select Case Err
'
'        Case 14029
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144723)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function HistPadrao_Exclui(iHistPadrao As Integer) As Long
''exclui o Histórico Padrão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_HistPadrao_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_HISTPADRAO_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 14007
'
'
'    HistPadrao_Exclui = X.HistPadrao_Exclui(iHistPadrao)
'
'    Exit Function
'
'Erro_HistPadrao_Exclui:
'
'    HistPadrao_Exclui = Err
'
'    Select Case Err
'
'        Case 14007
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144724)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'Function HistPadrao_Grava(objHistPadrao As ClassHistPadrao, iOperacao As Integer) As Long
''grava o Histórico Padrão no banco de dados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_HistPadrao_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_HISTPADRAO_GRAVA", X)
'    If lErro <> SUCESSO Then Error 14000
'
'
'    HistPadrao_Grava = X.HistPadrao_Grava(objHistPadrao, iOperacao)
'
'    Exit Function
'
'Erro_HistPadrao_Grava:
'
'    HistPadrao_Grava = Err
'
'    Select Case Err
'
'        Case 14000
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144725)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'Function HistPadrao_Automatico(iHistPadrao As Integer) As Long
''retorna o número do proximo HistPadrao disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_HistPadrao_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_HISTPADRAO_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 14008
'
'
'    HistPadrao_Automatico = X.HistPadrao_Automatico(iHistPadrao)
'
'    Exit Function
'
'Erro_HistPadrao_Automatico:
'
'    HistPadrao_Automatico = Err
'
'    Select Case Err
'
'        Case 14008
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144726)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Function Contas_Seleciona(colNodes As Nodes)
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_Contas_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_CONTAS_SELECIONA",X)
''    If lErro <> SUCESSO Then Error 7017
''
''
''    Contas_Seleciona = X.Contas_Seleciona(colNodes)
''    Exit Function
''
''Erro_Contas_Seleciona:
''    Contas_Seleciona = Err
''    Select Case Err
''
''        Case 7017
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144727)
''
''    End Select
''    Exit Function
''End Function
'
'Function Conta_SelecionaUma(sConta As String, objPlanoConta As ClassPlanoConta, sSiglaModulo As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_SelecionaUma
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_SELECIONA_UMA", X)
'    If lErro <> SUCESSO Then Error 7018
'
'
'    Conta_SelecionaUma = X.Conta_SelecionaUma(sConta, objPlanoConta, sSiglaModulo)
'    Exit Function
'
'Erro_Conta_SelecionaUma:
'    Conta_SelecionaUma = Err
'    Select Case Err
'
'        Case 7018
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144728)
'
'
'    End Select
'    Exit Function
'End Function
'
'Public Function Conta_Perde_Foco(Conta As Object, Desc As Object) As Long
''recebe MaskEdBox da conta e o label da descrição
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Perde_Foco
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_PERDE_FOCO", X)
'    If lErro <> SUCESSO Then Error 14066
'
'
'    Conta_Perde_Foco = X.Conta_Perde_Foco(Conta, Desc)
'    Exit Function
'
'Erro_Conta_Perde_Foco:
'
'    Conta_Perde_Foco = Err
'
'    Select Case Err
'
'        Case 14066
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144729)
'
'
'    End Select
'    Exit Function
'End Function
'
'Public Function Traz_Conta_MaskEd(sConta As String, MaskEdConta As Object, DescConta As Object) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Traz_Conta_MaskEd
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_TRAZ_CONTA_MASKED", X)
'    If lErro <> SUCESSO Then Error 14067
'
'
'    Traz_Conta_MaskEd = X.Traz_Conta_MaskEd(sConta, MaskEdConta, DescConta)
'    Exit Function
'
'Erro_Traz_Conta_MaskEd:
'
'    Traz_Conta_MaskEd = Err
'
'    Select Case Err
'
'        Case 14067
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144730)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Carga_Arvore_Conta(colNodes As Nodes) As Long
''copiado de PlanoConta
''move os dados do plano de contas do banco de dados para a arvore colNodes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Carga_Arvore_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CARGA_ARVORE_CONTA", X)
'    If lErro <> SUCESSO Then Error 14068
'
'
'    Carga_Arvore_Conta = X.Carga_Arvore_Conta(colNodes)
'    Exit Function
'
'Erro_Carga_Arvore_Conta:
'
'    Carga_Arvore_Conta = Err
'
'    Select Case Err
'
'        Case 14068
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144731)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Carga_Arvore_Conta1(objNodeAvo As Node, colNodes As Nodes) As Long
''copiado de PlanoConta
''move os dados do plano de contas do banco de dados para a arvore colNodes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Carga_Arvore_Conta1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CARGA_ARVORE_CONTA1", X)
'    If lErro <> SUCESSO Then Error 20595
'
'
'    Carga_Arvore_Conta1 = X.Carga_Arvore_Conta1(objNodeAvo, colNodes)
'    Exit Function
'
'Erro_Carga_Arvore_Conta1:
'
'    Carga_Arvore_Conta1 = Err
'
'    Select Case Err
'
'        Case 20595
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144732)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Inicializa_Mascara_Conta_MaskEd(MaskEdConta As Object) As Long
''inicializa a mascara de conta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Inicializa_Mascara_Conta_MaskEd
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_INICIALIZA_MASCARA_CONTA_MASKED", X)
'    If lErro <> SUCESSO Then Error 14069
'
'    Inicializa_Mascara_Conta_MaskEd = X.Inicializa_Mascara_Conta_MaskEd(MaskEdConta)
'    Exit Function
'
'Erro_Inicializa_Mascara_Conta_MaskEd:
'
'    Inicializa_Mascara_Conta_MaskEd = Err
'
'    Select Case Err
'
'        Case 14069
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144733)
'
'
'    End Select
'    Exit Function
'End Function
'
''Function Ccls_Seleciona(colNodes As Nodes)
''
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_Ccls_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_CCLS_SELECIONA",X)
''    If lErro <> SUCESSO Then Error 7019
''
''
''    Ccls_Seleciona = X.Ccls_Seleciona(colNodes)
''
''    Exit Function
''
''Erro_Ccls_Seleciona:
''    Ccls_Seleciona = Err
''    Select Case Err
''
''        Case 7019
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144734)
''
''    End Select
''    Exit Function
''End Function
'
''Function Lote_Seleciona(iExercicio As Integer, tLote() As ClassLote, iNum_Lotes_Abertos As Integer) As Long
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_Lote_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_SELECIONA",X)
''    If lErro <> SUCESSO Then Error 7020
''
''
''    Lote_Seleciona = X.Lote_Seleciona(iExercicio, tLote, iNum_Lotes_Abertos)
''    Exit Function
''
''Erro_Lote_Seleciona:
''    Lote_Seleciona = Err
''    Select Case Err
''
''        Case 7020
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144735)
''
''    End Select
''    Exit Function
''End Function
'
'Function LotePendente_Le(objLote As ClassLote) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE", X)
'    If lErro <> SUCESSO Then Error 7021
'
'
'    LotePendente_Le = X.LotePendente_Le(objLote)
'
'    Exit Function
'
'Erro_LotePendente_Le:
'
'    LotePendente_Le = Err
'
'    Select Case Err
'
'        Case 7021
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144736)
'
'    End Select
'    Exit Function
'End Function
'
''Function Lote_Periodo(sOrigem As String, iExercicio As Integer, iLote As Integer, iPeriodo As Integer) As Long
''
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_Lote_Periodo
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_PERIODO",X)
''    If lErro <> SUCESSO Then Error 7022
''
''
''    Lote_Periodo = X.Lote_Periodo(sOrigem, iExercicio, iLote, iPeriodo)
''
''    Exit Function
''
''Erro_Lote_Periodo:
''    Lote_Periodo = Err
''    Select Case Err
''
''        Case 7022
''
''
''        Case Else
''
''
''    End Select
''    Exit Function
''End Function
'
'Function HistPadrao_Le(objHistPadrao As ClassHistPadrao) As Long
''Le o historico padrao a partir da chave fornecida em objHistPadrao. Retorna os dados em objHistPadrao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_HistPadrao_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_HISTPADRAO_LE", X)
'    If lErro <> SUCESSO Then Error 7023
'
'
'    HistPadrao_Le = X.HistPadrao_Le(objHistPadrao)
'
'    Exit Function
'
'Erro_HistPadrao_Le:
'    HistPadrao_Le = Err
'    Select Case Err
'
'        Case 7023
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144737)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function HistPadrao_Le_Todos_Historicos(colHistPadrao As Collection) As Long
''Le os históricos padrões do BD e preenche a ListBox com
''concatenações do número do histórico + descrição do histórico
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_HistPadrao_Le_Todos_Historicos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_HISTPADRAO_LE_TODOS_HISTORICOS", X)
'    If lErro <> SUCESSO Then Error 14005
'
'
'    HistPadrao_Le_Todos_Historicos = X.HistPadrao_Le_Todos_Historicos(colHistPadrao)
'
'    Exit Function
'
'Erro_HistPadrao_Le_Todos_Historicos:
'
'    HistPadrao_Le_Todos_Historicos = Err
'
'    Select Case Err
'
'        Case 14005
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144738)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'Function PlanoConta_Le_ContaSimples(ByVal lContaSimples As Long, objPlanoConta As ClassPlanoConta) As Long
''le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_ContaSimples
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_CONTASIMPLES", X)
'    If lErro <> SUCESSO Then Error 7024
'
'
'    PlanoConta_Le_ContaSimples = X.PlanoConta_Le_ContaSimples(lContaSimples, objPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_ContaSimples:
'    PlanoConta_Le_ContaSimples = Err
'    Select Case Err
'
'        Case 7024
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144739)
'
'
'    End Select
'    Exit Function
'End Function
'Function PlanoContas_Le_HistPadrao(iHistPadrao As Integer) As Long
''Le o historico padrao no plano de contas.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoContas_Le_HistPadrao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTAS_LE_HISTPADRAO", X)
'    If lErro <> SUCESSO Then Error 14006
'
'
'    PlanoContas_Le_HistPadrao = X.PlanoContas_Le_HistPadrao(iHistPadrao)
'
'    Exit Function
'
'Erro_PlanoContas_Le_HistPadrao:
'    PlanoContas_Le_HistPadrao = Err
'    Select Case Err
'
'        Case 14006
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144740)
'
'
'    End Select
'    Exit Function
'End Function
'Function Lancamento_Le_Conta(sConta As String) As Long
''verifica se existe no BD pelo menos um lancamento para a conta fornecida
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 7025
'
'
'    Lancamento_Le_Conta = X.Lancamento_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_Lancamento_Le_Conta:
'    Lancamento_Le_Conta = Err
'    Select Case Err
'
'        Case 7025
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144741)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function LanPendente_Le_Conta(sConta As String) As Long
''verifica se existe no BD pelo menos um lancamento pendente para a conta fornecida
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 7026
'
'
'    LanPendente_Le_Conta = X.LanPendente_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_LanPendente_Le_Conta:
'    LanPendente_Le_Conta = Err
'    Select Case Err
'
'        Case 7026
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144742)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function ContaCcl_Le_Conta(sConta As String) As Long
''verifica se existe no BD pelo menos uma associacao da conta fornecida com centros de custo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 7027
'
'
'    ContaCcl_Le_Conta = X.ContaCcl_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_ContaCcl_Le_Conta:
'
'    ContaCcl_Le_Conta = Err
'
'    Select Case Err
'
'        Case 7027
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144743)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCcl_Exclui_Conta_ExtraContabil(sConta As String) As Long
''exclui as associações da conta com centros de custo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_Conta_ExtraContabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_CONTA_EXTRACONTABIL", X)
'    If lErro <> SUCESSO Then Error 14185
'
'
'    ContaCcl_Exclui_Conta_ExtraContabil = X.ContaCcl_Exclui_Conta_ExtraContabil(sConta)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_Conta_ExtraContabil:
'
'    ContaCcl_Exclui_Conta_ExtraContabil = Err
'
'    Select Case Err
'
'        Case 14185
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144744)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCcl_Exclui_Conta_Contabil(ByVal sConta As String) As Long
''Exclui a conta de PlanoContas. Tranferir para ClassGrava.
''O PROGRAMA DEVE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Exclui_Conta_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_EXCLUI_CONTA_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14186
'
'
'    ContaCcl_Exclui_Conta_Contabil = X.ContaCcl_Exclui_Conta_Contabil(sConta)
'
'    Exit Function
'
'Erro_ContaCcl_Exclui_Conta_Contabil:
'
'    ContaCcl_Exclui_Conta_Contabil = Err
'
'    Select Case Err
'
'        Case 14186
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144745)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_ContaMaior(sConta As String, objPlanoConta As ClassPlanoConta) As Long
''le uma conta imediatamente maior que a conta passada como parametro. Devolve os dados em objPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_ContaMaior
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_CONTAMAIOR", X)
'    If lErro <> SUCESSO Then Error 7028
'
'
'    PlanoConta_Le_ContaMaior = X.PlanoConta_Le_ContaMaior(sConta, objPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_ContaMaior:
'    PlanoConta_Le_ContaMaior = Err
'    Select Case Err
'
'        Case 7028
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144746)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Lancamento_Le_Lote(objLote As ClassLote) As Long
''verifica se existe no BD pelo menos um lancamento para o lote fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Lote
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_LOTE", X)
'    If lErro <> SUCESSO Then Error 7029
'
'
'    Lancamento_Le_Lote = X.Lancamento_Le_Lote(objLote)
'
'    Exit Function
'
'Erro_Lancamento_Le_Lote:
'    Lancamento_Le_Lote = Err
'    Select Case Err
'
'        Case 7029
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144747)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Lancamento_Le_Doc(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
''Tenta ler o documento em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 14025
'
'
'    Lancamento_Le_Doc = X.Lancamento_Le_Doc(objLancamento_Cabecalho)
'
'    Exit Function
'
'Erro_Lancamento_Le_Doc:
'
'    Lancamento_Le_Doc = Err
'
'    Select Case Err
'
'        Case 14025
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144748)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_Lote(objLote As ClassLote) As Long
''verifica se existe no BD pelo menos um lancamento pendente para o lote fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Lote
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_LOTE", X)
'    If lErro <> SUCESSO Then Error 7030
'
'
'    LanPendente_Le_Lote = X.LanPendente_Le_Lote(objLote)
'
'    Exit Function
'
'Erro_LanPendente_Le_Lote:
'    LanPendente_Le_Lote = Err
'    Select Case Err
'
'        Case 7030
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144749)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function RateioOn_Le_Conta(sConta As String) As Long
''verifica se existe na tabela RateioOn pelo menos um rateio com a conta em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 7031
'
'
'    RateioOn_Le_Conta = X.RateioOn_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_RateioOn_Le_Conta:
'    RateioOn_Le_Conta = Err
'    Select Case Err
'
'        Case 7031
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144750)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOn_Le_Doc(objRateioOn As ClassRateioOn, colRateioOns As Collection) As Long
''le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOns
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 14176
'
'
'    RateioOn_Le_Doc = X.RateioOn_Le_Doc(objRateioOn, colRateioOns)
'
'    Exit Function
'
'Erro_RateioOn_Le_Doc:
'
'    RateioOn_Le_Doc = Err
'
'    Select Case Err
'
'        Case 14176
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144751)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOn_Exclui(iCodigo As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 14177
'
'
'    RateioOn_Exclui = X.RateioOn_Exclui(iCodigo)
'
'    Exit Function
'
'Erro_RateioOn_Exclui:
'
'    RateioOn_Exclui = Err
'
'    Select Case Err
'
'        Case 14177
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144752)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOn_Grava(iCodigo As Integer, sDescricao As String, colRateioOns As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_GRAVA", X)
'    If lErro <> SUCESSO Then Error 14178
'
'
'    RateioOn_Grava = X.RateioOn_Grava(iCodigo, sDescricao, colRateioOns)
'
'    Exit Function
'
'Erro_RateioOn_Grava:
'
'    RateioOn_Grava = Err
'
'    Select Case Err
'
'        Case 14178
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144753)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOn_Grava_BD(iCodigo As Integer, sDescricao As String, colRateioOns As Collection) As Long
''grava na tabela de Rateios Automaticos (RateioOn)lançamentos do voucher em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Grava_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_GRAVA_BD", X)
'    If lErro <> SUCESSO Then Error 14179
'
'
'    RateioOn_Grava_BD = X.RateioOn_Grava_BD(iCodigo, sDescricao, colRateioOns)
'
'    Exit Function
'
'Erro_RateioOn_Grava_BD:
'
'    RateioOn_Grava_BD = Err
'
'    Select Case Err
'
'        Case 14179
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144754)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOn_Automatico(iCodigo As Integer) As Long
''retorna o número do proximo código de Rateio disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 14181
'
'
'    RateioOn_Automatico = X.RateioOn_Automatico(iCodigo)
'
'    Exit Function
'
'Erro_RateioOn_Automatico:
'
'    RateioOn_Automatico = Err
'
'    Select Case Err
'
'        Case 14181
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144755)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Function RateioOffCab_Le_Conta(sConta As String) As Long
'''verifica se existe na tabela RateioOffCab pelo menos um rateio com a conta em questão.
''
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_RateioOffCab_Le_Conta
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFFCAB_LE_CONTA",X)
''    If lErro <> SUCESSO Then Error 7032
''
''
''    RateioOffCab_Le_Conta = X.RateioOffCab_Le_Conta(sConta)
''
''
''    Exit Function
''
''Erro_RateioOffCab_Le_Conta:
''    RateioOffCab_Le_Conta = Err
''    Select Case Err
''
''        Case 7032
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144756)
''
''    End Select
''    Exit Function
''End Function
'
'Function RateioOff_Le_Conta(sConta As String) As Long
''verifica se existe na tabela RateioOff pelo menos um rateio com a conta em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 7033
'
'
'    RateioOff_Le_Conta = X.RateioOff_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_RateioOff_Le_Conta:
'
'    RateioOff_Le_Conta = Err
'
'    Select Case Err
'
'        Case 7033
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144757)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Le_Doc(objRateioOff As ClassRateioOff, colRateioOff As Collection) As Long
''le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOff
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 14260
'
'
'    RateioOff_Le_Doc = X.RateioOff_Le_Doc(objRateioOff, colRateioOff)
'
'    Exit Function
'
'Erro_RateioOff_Le_Doc:
'
'    RateioOff_Le_Doc = Err
'
'    Select Case Err
'
'        Case 14260
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144758)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Exclui(lCodigo As Long) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 14261
'
'
'    RateioOff_Exclui = X.RateioOff_Exclui(lCodigo)
'
'    Exit Function
'
'Erro_RateioOff_Exclui:
'
'    RateioOff_Exclui = Err
'
'    Select Case Err
'
'        Case 14261
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144759)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Grava(objRateioOff As ClassRateioOff, colRateioOff As Collection, colContas As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_GRAVA", X)
'    If lErro <> SUCESSO Then Error 14262
'
'
'    RateioOff_Grava = X.RateioOff_Grava(objRateioOff, colRateioOff, colContas)
'
'    Exit Function
'
'Erro_RateioOff_Grava:
'
'    RateioOff_Grava = Err
'
'    Select Case Err
'
'        Case 14262
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144760)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Grava_BD(objRateioOff1 As ClassRateioOff, colRateioOff As Collection) As Long
''grava na tabela de Rateios Automaticos (RateioOff)lançamentos do voucher em questão
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Grava_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_GRAVA_BD", X)
'    If lErro <> SUCESSO Then Error 14263
'
'
'    RateioOff_Grava_BD = X.RateioOff_Grava_BD(objRateioOff1, colRateioOff)
'
'    Exit Function
'
'Erro_RateioOff_Grava_BD:
'
'    RateioOff_Grava_BD = Err
'
'    Select Case Err
'
'        Case 14263
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144761)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Exclui_BD(lCodigo As Long) As Long
''Exclui da tabela RateioOff os Rateios com CclOrigem= sCclOrigem e ContaOrigem= sContaOrigem
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Exclui_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_EXCLUI_BD", X)
'    If lErro <> SUCESSO Then Error 14264
'
'
'    RateioOff_Exclui_BD = X.RateioOff_Exclui_BD(lCodigo)
'
'    Exit Function
'
'Erro_RateioOff_Exclui_BD:
'
'    RateioOff_Exclui_BD = Err
'
'    Select Case Err
'
'        Case 14264
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144762)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Le_Conta(sConta As String) As Long
''verifica se existe na tabela DocAuto pelo menos um documento automático com a conta em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 7034
'
'
'    DocAuto_Le_Conta = X.DocAuto_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_DocAuto_Le_Conta:
'    DocAuto_Le_Conta = Err
'    Select Case Err
'
'        Case 7034
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144763)
'
'
'    End Select
'    Exit Function
'
'End Function
'
'Function DocAuto_Le_Doc(objDocAuto As ClassDocAuto, colDocAutos As Collection) As Long
''le os lançamentos relativos ao documento passado como parametro e coloca-os em colDocAutos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 14070
'
'
'    DocAuto_Le_Doc = X.DocAuto_Le_Doc(objDocAuto, colDocAutos)
'
'    Exit Function
'
'Erro_DocAuto_Le_Doc:
'
'    DocAuto_Le_Doc = Err
'
'    Select Case Err
'
'        Case 14070
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144764)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Le Todos os dados dos Documentos Para fazer o carregamento da Arvore de Documento Automatico
'Function DocAuto_Le_Todos(colDocAuto As Collection)
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 14307
'
'
'    DocAuto_Le_Todos = X.DocAuto_Le_Todos(colDocAuto)
'
'    Exit Function
'
'Erro_DocAuto_Le_Todos:
'
'    DocAuto_Le_Todos = Err
'
'    Select Case Err
'
'        Case 14307
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144765)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Tem_Filho(sConta As String, iTemFilho As Integer) As Long
''verifica se a conta em questão tem contas filha
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Tem_Filho
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_TEM_FILHO", X)
'    If lErro <> SUCESSO Then Error 7035
'
'
'    PlanoConta_Tem_Filho = X.PlanoConta_Tem_Filho(sConta, iTemFilho)
'
'    Exit Function
'
'Erro_PlanoConta_Tem_Filho:
'    PlanoConta_Tem_Filho = Err
'    Select Case Err
'
'        Case 7035
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144766)
'
'
'    End Select
'    Exit Function
'
'End Function
'
'Function Ccl_Tem_Filho(sCcl As String, iTemFilho As Integer) As Long
''testa se o centro de custo em questão tem centros de custo/lucro filhos.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Tem_Filho
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_TEM_FILHO", X)
'    If lErro <> SUCESSO Then Error 10376
'
'
'
'    Ccl_Tem_Filho = X.Ccl_Tem_Filho(sCcl, iTemFilho)
'
'    Exit Function
'
'Erro_Ccl_Tem_Filho:
'
'    Ccl_Tem_Filho = Err
'
'    Select Case Err
'
'        Case 10376
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144767)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_ExisteConta(iExiste As Integer) As Long
''Verifica se existe pelo menos uma conta cadastrada em PlanoContano BD
''se existir iExiste = 1 , caso contrario iExiste = 0
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_ExisteConta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_EXISTECONTA", X)
'    If lErro <> SUCESSO Then Error 14184
'
'
'    PlanoConta_ExisteConta = X.PlanoConta_ExisteConta(iExiste)
'
'    Exit Function
'
'Erro_PlanoConta_ExisteConta:
'
'    PlanoConta_ExisteConta = Err
'
'    Select Case Err
'
'        Case 14184
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144768)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Function DocumentoLancado_NaoExiste(lDoc As Long, sOrigem As String, iExercicio As Integer, iPeriodoLan As Integer) As Long
''
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_DocumentoLancado_NaoExiste
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_DOCUMENTOLANCADO_NAOEXISTE",X)
''    If lErro <> SUCESSO Then Error 7037
''
''
''    DocumentoLancado_NaoExiste = X.DocumentoLancado_NaoExiste(lDoc, sOrigem, iExercicio, iPeriodoLan)
''
''    Exit Function
''
''Erro_DocumentoLancado_NaoExiste:
''    DocumentoLancado_NaoExiste = Err
''    Select Case Err
''
''        Case 7037
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144769)
''
''    End Select
''    Exit Function
''End Function
'
'Function Periodo_Le(ByVal dtData As Date, objPeriodo As ClassPeriodo) As Long
''le o registro referente ao periodo em questão e coloca os dados em objPeriodo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE", X)
'    If lErro <> SUCESSO Then Error 7036
'
'
'    Periodo_Le = X.Periodo_Le(dtData, objPeriodo)
'
'    Exit Function
'
'Erro_Periodo_Le:
'    Periodo_Le = Err
'    Select Case Err
'
'        Case 7036
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144770)
'
'
'    End Select
'    Exit Function
'End Function
'
''Function DocumentoPendente_NaoExiste(lDoc As Long, sOrigem As String, iExercicio As Integer, iPeriodoLan As Integer) As Long
''
''    Dim X As Object, lErro As Long
''
''    On Error GoTo Erro_DocumentoPendente_NaoExiste
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_DOCUMENTOPENDENTE_NAOEXISTE",X)
''    If lErro <> SUCESSO Then Error 7038
''
''
''    DocumentoPendente_NaoExiste = X.DocumentoPendente_NaoExiste(lDoc, sOrigem, iExercicio, iPeriodoLan)
''
''    Exit Function
''
''Erro_DocumentoPendente_NaoExiste:
''    DocumentoPendente_NaoExiste = Err
''    Select Case Err
''
''        Case 7038
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144771)
''
''    End Select
''    Exit Function
''End Function
'
'Function Ccl_Le(objCcl As ClassCcl) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_LE", X)
'    If lErro <> SUCESSO Then Error 7039
'
'
'    Ccl_Le = X.Ccl_Le(objCcl)
'
'    Exit Function
'
'Erro_Ccl_Le:
'    Ccl_Le = Err
'    Select Case Err
'
'        Case 7039
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144772)
'
'
'    End Select
'    Exit Function
'
'End Function
'
'Function ContaCcl_Le_Ccl(sCcl As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE_CCL", X)
'    If lErro <> SUCESSO Then Error 7040
'
'
'    ContaCcl_Le_Ccl = X.ContaCcl_Le_Ccl(sCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Le_Ccl:
'    ContaCcl_Le_Ccl = Err
'    Select Case Err
'
'        Case 7040
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144773)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function Lancamento_Le_Ccl(sCcl As String) As Long
''verifica se existe no BD pelo menos um lancamento para o centro de custo/lucro fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_CCL", X)
'    If lErro <> SUCESSO Then Error 7041
'
'
'    Lancamento_Le_Ccl = X.Lancamento_Le_Ccl(sCcl)
'
'    Exit Function
'
'Erro_Lancamento_Le_Ccl:
'    Lancamento_Le_Ccl = Err
'    Select Case Err
'
'        Case 7041
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144774)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function LanPendente_Le_Ccl(sCcl As String) As Long
''verifica se existe no BD pelo menos um lancamento pendente para o centro deu custo/lucro fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_CCL", X)
'    If lErro <> SUCESSO Then Error 7042
'
'
'    LanPendente_Le_Ccl = X.LanPendente_Le_Ccl(sCcl)
'
'    Exit Function
'
'Erro_LanPendente_Le_Ccl:
'    LanPendente_Le_Ccl = Err
'    Select Case Err
'
'        Case 7042
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144775)
'
'
'    End Select
'    Exit Function
'End Function
'
'Function DocAuto_Le_Ccl(sCcl As String) As Long
''verifica se existe na tabela DocAuto pelo menos um documento automático com o Ccl em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Le_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_LE_CCL", X)
'    If lErro <> SUCESSO Then Error 5672
'
'
'    DocAuto_Le_Ccl = X.DocAuto_Le_Ccl(sCcl)
'
'    Exit Function
'
'Erro_DocAuto_Le_Ccl:
'
'    DocAuto_Le_Ccl = Err
'
'    Select Case Err
'
'        Case 5672
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144776)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function RateioOn_Le_Ccl(sCcl As String) As Long
''verifica se existe na tabela RateioOn pelo menos um rateio com o Ccl em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Le_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_LE_CCL", X)
'    If lErro <> SUCESSO Then Error 5673
'
'
'    RateioOn_Le_Ccl = X.RateioOn_Le_Ccl(sCcl)
'
'    Exit Function
'
'Erro_RateioOn_Le_Ccl:
'
'    RateioOn_Le_Ccl = Err
'
'    Select Case Err
'
'        Case 5673
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144777)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Function RateioOffCab_Le_Ccl(sCcl As String) As Long
'''verifica se existe na tabela RateioOffCab pelo menos um rateio com o Ccl em questão.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_RateioOffCab_Le_Ccl
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFFCAB_LE_CCL",X)
''    If lErro <> SUCESSO Then Error 5674
''
''
''    RateioOffCab_Le_Ccl = X.RateioOffCab_Le_Ccl(sCcl)
''
''    Exit Function
''
''Erro_RateioOffCab_Le_Ccl:
''
''    RateioOffCab_Le_Ccl = Err
''
''    Select Case Err
''
''        Case 5674
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144778)
''
''    End Select
''
''    Exit Function
''
''End Function
'
'Function RateioOff_Le_Ccl(sCcl As String) As Long
''verifica se existe na tabela RateioOff pelo menos um rateio com o Ccl em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Le_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_LE_CCL", X)
'    If lErro <> SUCESSO Then Error 5675
'
'
'    RateioOff_Le_Ccl = X.RateioOff_Le_Ccl(sCcl)
'
'    Exit Function
'
'Erro_RateioOff_Le_Ccl:
'
'    RateioOff_Le_Ccl = Err
'
'    Select Case Err
'
'        Case 5675
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144779)
'
'
'    End Select
'
'    Exit Function
'
'End Function
''Function Carga_Controle_Periodos(ByVal iExercicio As Integer, objListCombo As Object) As Long
'''carrega o controle passado como parametro com os periodos do exercicio passadao como parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Carga_Controle_Periodos
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("CTB_CARGA_CONTROLE_PERIODOS",X)
''    If lErro <> SUCESSO Then Error 14004
''
''
''    Carga_Controle_Periodos = X.Carga_Controle_Periodos(iExercicio, objListCombo)
''
''    Exit Function
''
''Erro_Carga_Controle_Periodos:
''
''    Carga_Controle_Periodos = Err
''
''    Select Case Err
''
''        Case 14004
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144780)
''
''    End Select
''
''    Exit Function
''
''End Function
'Function Lote_Automatico(objLote As ClassLote) As Long
''retorna o número do proximo lote disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lote_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 14001
'
'
'    Lote_Automatico = X.Lote_Automatico(objLote)
'
'    Exit Function
'
'Erro_Lote_Automatico:
'
'    Lote_Automatico = Err
'
'    Select Case Err
'
'        Case 14001
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144781)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lote_Automatico1(objLote As ClassLote) As Long
''retorna o número do proximo lote disponivel
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lote_Automatico1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_AUTOMATICO1", X)
'    If lErro <> SUCESSO Then Error 39703
'
'
'    Lote_Automatico1 = X.Lote_Automatico1(objLote)
'
'    Exit Function
'
'Erro_Lote_Automatico1:
'
'    Lote_Automatico1 = Err
'
'    Select Case Err
'
'        Case 39703
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144782)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Voucher_Automatico(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, sOrigem As String, lDoc As Long) As Long
''retorna o número do proximo voucher (documento) disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Voucher_Automatico
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VOUCHER_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 14026
'
'    Voucher_Automatico = X.Voucher_Automatico(iFilialEmpresa, iExercicio, iPeriodo, sOrigem, lDoc)
'
'    Exit Function
'
'Erro_Voucher_Automatico:
'
'    Voucher_Automatico = Err
'
'    Select Case Err
'
'        Case 14026
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144783)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Automatico(lDoc As Long) As Long
''retorna o número do proximo voucher (documento) disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 14071
'
'
'    DocAuto_Automatico = X.DocAuto_Automatico(lDoc)
'
'    Exit Function
'
'Erro_DocAuto_Automatico:
'
'    DocAuto_Automatico = Err
'
'    Select Case Err
'
'        Case 14071
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144784)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Exclui(lDoc As Long) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 14072
'
'
'    DocAuto_Exclui = X.DocAuto_Exclui(lDoc)
'
'    Exit Function
'
'Erro_DocAuto_Exclui:
'
'    DocAuto_Exclui = Err
'
'    Select Case Err
'
'        Case 14072
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144785)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Grava(lDoc As Long, sOrigem As String, sDescricao As String, colDocAutos As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_GRAVA", X)
'    If lErro <> SUCESSO Then Error 14073
'
'
'    DocAuto_Grava = X.DocAuto_Grava(lDoc, sOrigem, sDescricao, colDocAutos)
'
'    Exit Function
'
'Erro_DocAuto_Grava:
'
'    DocAuto_Grava = Err
'
'    Select Case Err
'
'        Case 14073
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144786)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Grava_BD(lDoc As Long, sOrigem As String, sDescricao As String, colDocAutos As Collection) As Long
''grava na tabela de Documentos Automaticos (DocAuto)lançamentos do voucher em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Grava_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_GRAVA_BD", X)
'    If lErro <> SUCESSO Then Error 14074
'
'
'    DocAuto_Grava_BD = X.DocAuto_Grava_BD(lDoc, sOrigem, sDescricao, colDocAutos)
'
'    Exit Function
'
'Erro_DocAuto_Grava_BD:
'
'    DocAuto_Grava_BD = Err
'
'    Select Case Err
'
'        Case 14074
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144787)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DocAuto_Exclui_BD(lDoc As Long)
''exclui da tabela DocAuto os Documentos Automáticos do voucher em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DocAuto_Exclui_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DOCAUTO_EXCLUI_BD", X)
'    If lErro <> SUCESSO Then Error 14075
'
'
'    DocAuto_Exclui_BD = X.DocAuto_Exclui_BD(lDoc)
'
'    Exit Function
'
'Erro_DocAuto_Exclui_BD:
'
'    DocAuto_Exclui_BD = Err
'
'    Select Case Err
'
'        Case 14075
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144788)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Atualizacao_ExtraContabil(colContas As Collection, colCcl As Collection) As Long
''Inicia as atualizacoes correspondentes a marcacao da listbox
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Atualizacao_ExtraContabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_ATUALIZACAO_EXTRACONTABIL", X)
'    If lErro <> SUCESSO Then Error 14013
'
'
'    ContaCcl_Atualizacao_ExtraContabil = X.ContaCcl_Atualizacao_ExtraContabil(colContas, colCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Atualizacao_ExtraContabil:
'
'    ContaCcl_Atualizacao_ExtraContabil = Err
'
'    Select Case Err
'
'        Case 14013
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144789)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Atualizacao_Contabil(colContas As Collection, colCcl As Collection) As Long
''Inicia as atualizacoes correspondentes a marcacao da listbox - Ccl Contábil
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Atualizacao_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_ATUALIZACAO_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 14027
'
'
'    ContaCcl_Atualizacao_Contabil = X.ContaCcl_Atualizacao_Contabil(colContas, colCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Atualizacao_Contabil:
'
'    ContaCcl_Atualizacao_Contabil = Err
'
'    Select Case Err
'
'        Case 14027
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144790)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Associa_Conta(colContas As Collection, ByVal sCcl As String) As Long
''Preenche colecao para marcacao de listbox Contas
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Associa_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_ASSOCIA_CONTA", X)
'    If lErro <> SUCESSO Then Error 14014
'
'
'    ContaCcl_Associa_Conta = X.ContaCcl_Associa_Conta(colContas, sCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Associa_Conta:
'
'    ContaCcl_Associa_Conta = Err
'
'    Select Case Err
'
'        Case 14014
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144791)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Associa_Ccl(colCcl As Collection, ByVal sConta As String) As Long
''Preenche colecao para marcacao de listbox Ccl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Associa_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_ASSOCIA_CCL", X)
'    If lErro <> SUCESSO Then Error 14016
'
'
'    ContaCcl_Associa_Ccl = X.ContaCcl_Associa_Ccl(colCcl, sConta)
'
'    Exit Function
'
'Erro_ContaCcl_Associa_Ccl:
'
'    ContaCcl_Associa_Ccl = Err
'
'    Select Case Err
'
'        Case 14016
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144792)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function PlanoConta_Le_Conta(ByVal sConta As String) As Long
''Verifica no BD a existencia de sConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_CONTA", X)
'    If lErro <> SUCESSO Then Error 14015
'
'
'    PlanoConta_Le_Conta = X.PlanoConta_Le_Conta(sConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Conta:
'
'    PlanoConta_Le_Conta = Err
'
'    Select Case Err
'
'        Case 14015
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144793)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Le_Todos_Ccl(colCcl As Collection) As Long
''Preenche colecao para preencher listbox de Ccl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le_Todos_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE_TODOS_CCL", X)
'    If lErro <> SUCESSO Then Error 14017
'
'
'    ContaCcl_Le_Todos_Ccl = X.ContaCcl_Le_Todos_Ccl(colCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Le_Todos_Ccl:
'
'    ContaCcl_Le_Todos_Ccl = Err
'
'    Select Case Err
'
'        Case 14017
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144794)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Le_Todos_Conta(colContas As Collection, ByVal iTipoConta As Integer) As Long
''Preenche a colecao colContas com o conteudo de PlanoContas, de acordo com iTipoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le_Todos_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE_TODOS_CONTA", X)
'    If lErro <> SUCESSO Then Error 14018
'
'
'    ContaCcl_Le_Todos_Conta = X.ContaCcl_Le_Todos_Conta(colContas, iTipoConta)
'
'    Exit Function
'
'Erro_ContaCcl_Le_Todos_Conta:
'
'    ContaCcl_Le_Todos_Conta = Err
'
'    Select Case Err
'
'        Case 14018
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144795)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Exercicio_Le(ByVal iExercicio As Integer, objExercicio As ClassExercicio) As Long
''le o registro referente ao exercicio em questão e coloca os dados em objExercicio
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE", X)
'    If lErro <> SUCESSO Then Error 5825
'
'
'    Exercicio_Le = X.Exercicio_Le(iExercicio, objExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Le:
'    Exercicio_Le = Err
'    Select Case Err
'
'        Case 5825
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144796)
'
'    End Select
'    Exit Function
'End Function
'
'Function Periodo_Le_ExercicioPeriodo(ByVal iExercicio As Integer, ByVal iPeriodo As Integer, objPeriodo As ClassPeriodo) As Long
''le o registro referente ao periodo em questão e coloca os dados em objPeriodo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le_ExercicioPeriodo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE_EXERCICIOPERIODO", X)
'    If lErro <> SUCESSO Then Error 5832
'
'
'    Periodo_Le_ExercicioPeriodo = X.Periodo_Le_ExercicioPeriodo(iExercicio, iPeriodo, objPeriodo)
'
'    Exit Function
'
'Erro_Periodo_Le_ExercicioPeriodo:
'
'    Periodo_Le_ExercicioPeriodo = Err
'
'    Select Case Err
'
'        Case 5832
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144797)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_Doc(objDoc As ClassLancamento_Cabecalho, colLancamentos As Collection) As Long
''le os lançamentos relativos ao documento passado como parametro e coloca-os em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 5840
'
'
'    LanPendente_Le_Doc = X.LanPendente_Le_Doc(objDoc, colLancamentos)
'
'    Exit Function
'
'Erro_LanPendente_Le_Doc:
'
'    LanPendente_Le_Doc = Err
'
'    Select Case Err
'
'        Case 5840
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144798)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_Doc1(objDoc As ClassLancamento_Cabecalho) As Long
''Tenta ler o documento em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Doc1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_DOC1", X)
'    If lErro <> SUCESSO Then Error 9594
'
'
'    LanPendente_Le_Doc1 = X.LanPendente_Le_Doc1(objDoc)
'
'    Exit Function
'
'Erro_LanPendente_Le_Doc1:
'
'    LanPendente_Le_Doc1 = Err
'
'    Select Case Err
'
'        Case 9594
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144799)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCcl_Le(objContaCcl As ClassContaCcl) As Long
''le a contaccl a partir da chave fornecida em objContaCcl. Devolve os dados em objContaCcl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE", X)
'    If lErro <> SUCESSO Then Error 5872
'
'
'    ContaCcl_Le = X.ContaCcl_Le(objContaCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Le:
'
'    ContaCcl_Le = Err
'
'    Select Case Err
'
'        Case 5872
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144800)
'
'
'    End Select
'
'    Exit Function
'
'
'End Function
'
'Function HistPadrao_Le_Todos(colHistPadrao As Collection)
''le todos os historicos contidos na tabela de historico padrão e coloca-os em colHistPadrao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_HistPadrao_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_HISTPADRAO_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 5891
'
'
'    HistPadrao_Le_Todos = X.HistPadrao_Le_Todos(colHistPadrao)
'
'    Exit Function
'
'Erro_HistPadrao_Le_Todos:
'
'    HistPadrao_Le_Todos = Err
'
'    Select Case Err
'
'        Case 5891
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144801)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Todos(colPlanoConta As Collection)
''le todos as contas contidas na tabela de Plano de Contas e coloca-as em colPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 5914
'
'
'    PlanoConta_Le_Todos = X.PlanoConta_Le_Todos(colPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Todos:
'
'    PlanoConta_Le_Todos = Err
'
'    Select Case Err
'
'        Case 5914
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144802)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Le_Todos(colCcl As Collection)
''le todos os centros de custo/lucro contidos na tabela Ccl e coloca-os em colCcl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 5919
'
'
'    Ccl_Le_Todos = X.Ccl_Le_Todos(colCcl)
'
'    Exit Function
'
'Erro_Ccl_Le_Todos:
'
'    Ccl_Le_Todos = Err
'
'    Select Case Err
'
'        Case 5919
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144803)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Configuracao_Le(objConfiguracao As ClassConfiguracao) As Long
''le o registro da tabela de Configuracao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Configuracao_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAO_LE", X)
'    If lErro <> SUCESSO Then Error 5960
'
'
'    Configuracao_Le = X.Configuracao_Le(objConfiguracao)
'
'    Exit Function
'
'Erro_Configuracao_Le:
'
'    Configuracao_Le = Err
'
'    Select Case Err
'
'        Case 5960
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144804)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_Data(ByVal dtData As Date, objExercicio As ClassExercicio) As Long
''retorna os dados do exercicio para a data em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_Data
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_DATA", X)
'    If lErro <> SUCESSO Then Error 5969
'
'
'    Exercicio_Le_Data = X.Exercicio_Le_Data(dtData, objExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Le_Data:
'
'    Exercicio_Le_Data = Err
'
'    Select Case Err
'
'        Case 5969
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144805)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicios_Abertos_Le_Todos(colExerciciosAbertos As Collection) As Long
''le todos os exercicios abertos e coloca-os em colExerciciosAbertos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicios_Abertos_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIOS_ABERTOS_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 5988
'
'
'    Exercicios_Abertos_Le_Todos = X.Exercicios_Abertos_Le_Todos(colExerciciosAbertos)
'
'    Exit Function
'
'Erro_Exercicios_Abertos_Le_Todos:
'
'    Exercicios_Abertos_Le_Todos = Err
'
'    Select Case Err
'
'        Case 5988
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144806)
'
'
'    End Select
'
'    Exit Function
'
'
'End Function
'
'Function Periodo_Le_Todos_Exercicio(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, colPeriodos As Collection) As Long
''le os periodos referentes ao exercicio iExercicio e coloca-os em colPeriodos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le_Todos_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE_TODOS_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 5989
'
'
'    Periodo_Le_Todos_Exercicio = X.Periodo_Le_Todos_Exercicio(iFilialEmpresa, iExercicio, colPeriodos)
'
'    Exit Function
'
'Erro_Periodo_Le_Todos_Exercicio:
'
'    Periodo_Le_Todos_Exercicio = Err
'
'    Select Case Err
'
'        Case 5989
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144807)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le_NumRegistros(lNumReg As Long, lComando As Long) As Long
''retorna o número de registros contidos na tabela de lotes pendentes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le_NumRegistros
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE_NUMREGISTROS", X)
'    If lErro <> SUCESSO Then Error 9011
'
'
'    LotePendente_Le_NumRegistros = X.LotePendente_Le_NumRegistros(lNumReg, lComando)
'
'    Exit Function
'
'Erro_LotePendente_Le_NumRegistros:
'
'    LotePendente_Le_NumRegistros = Err
'
'    Select Case Err
'
'        Case 9011
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144808)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le_Posicionado(ByVal colLote As Collection, ByVal lPosicao As Long, ByVal lNumReg As Long, lComando As Long, sOrigem As String, iExercicio As Integer, iPeriodo As Integer, iLote As Integer) As Long
''le os lotes lNumReg a partir da posicao lPosicao e coloca-os e colLote.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le_Posicionado
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE_POSICIONADO", X)
'    If lErro <> SUCESSO Then Error 9021
'
'
'    LotePendente_Le_Posicionado = X.LotePendente_Le_Posicionado(colLote, lPosicao, lNumReg, lComando, sOrigem, iExercicio, iPeriodo, iLote)
'
'    Exit Function
'
'Erro_LotePendente_Le_Posicionado:
'
'    LotePendente_Le_Posicionado = Err
'
'    Select Case Err
'
'        Case 9021
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144809)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_NomeExterno(ByVal lComando As Long, iExercicio As Integer, sNomeExterno As String) As Long
''le o nome externo do exercicio passado como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_NomeExterno
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_NOMEEXTERNO", X)
'    If lErro <> SUCESSO Then Error 9267
'
'
'    Exercicio_Le_NomeExterno = X.Exercicio_Le_NomeExterno(lComando, iExercicio, sNomeExterno)
'
'    Exit Function
'
'Erro_Exercicio_Le_NomeExterno:
'
'    Exercicio_Le_NomeExterno = Err
'
'    Select Case Err
'
'        Case 9267
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144810)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Periodo_Le_NomeExterno(ByVal lComando As Long, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long
''le o nome externo referente ao periodo em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le_NomeExterno
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE_NOMEEXTERNO", X)
'    If lErro <> SUCESSO Then Error 9268
'
'
'    Periodo_Le_NomeExterno = X.Periodo_Le_NomeExterno(lComando, iExercicio, iPeriodo, sNomeExterno)
'
'    Exit Function
'
'Erro_Periodo_Le_NomeExterno:
'
'    Periodo_Le_NomeExterno = Err
'
'    Select Case Err
'
'        Case 9268
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144811)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Origem_Le_Todos(colOrigem As Collection) As Long
''le todos as origens contidas na tabela Origem e coloca-as em colOrigem
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Origem_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORIGEM_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 9274
'
'
'    Origem_Le_Todos = X.Origem_Le_Todos(colOrigem)
'
'    Exit Function
'
'Erro_Origem_Le_Todos:
'
'    Origem_Le_Todos = Err
'
'    Select Case Err
'
'        Case 9274
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144812)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function PlanoConta_Le_Maior_Menor_Conta(sMaior As String, sMenor As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Maior_Menor_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELOPCOES_LE_MAIOR_MENOR_CONTA", X)
'    If lErro <> SUCESSO Then Error 14078
'
'
'    PlanoConta_Le_Maior_Menor_Conta = X.PlanoConta_Le_Maior_Menor_Conta(sMaior, sMenor)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Maior_Menor_Conta:
'
'    PlanoConta_Le_Maior_Menor_Conta = Err
'
'    Select Case Err
'
'        Case 14078
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144813)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lote_Le(objLote As ClassLote) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lote_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_LE", X)
'    If lErro <> SUCESSO Then Error 9294
'
'
'    Lote_Le = X.Lote_Le(objLote)
'
'    Exit Function
'
'Erro_Lote_Le:
'
'    Lote_Le = Err
'
'    Select Case Err
'
'        Case 9294
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144814)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Segmento_Grava_Conta(colSegmentos As Collection) As Long
''grava os registros da colecao na tabela Segmento
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Segmento_Grava_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SEGMENTO_GRAVA_CONTA", X)
'    If lErro <> SUCESSO Then Error 14182
'
'
'    Segmento_Grava_Conta = X.Segmento_Grava_Conta(colSegmentos)
'
'    Exit Function
'
'Erro_Segmento_Grava_Conta:
'
'    Segmento_Grava_Conta = Err
'
'    Select Case Err
'
'        Case 14182
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144815)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Segmento_Le_Codigo(objSegmento As ClassSegmento, colSegmentos As Collection) As Long
''le os campos (nivel , tipo , tamanho , delimitador , preenchimento) da tabela de Segmentos relativos ao codigo passado por parametro e coloca-os em colSegmentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Segmento_Le_Codigo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SEGMENTO_LE_CODIGO", X)
'    If lErro <> SUCESSO Then Error 14183
'
'
'    Segmento_Le_Codigo = X.Segmento_Le_Codigo(objSegmento, colSegmentos)
'
'    Exit Function
'
'Erro_Segmento_Le_Codigo:
'
'    Segmento_Le_Codigo = Err
'
'    Select Case Err
'
'        Case 14183
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144816)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicios_Le_Todos(colExercicios As Collection) As Long
''le  todos os exercicios  e coloca-os em colExercicios
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicios_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIOS_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 9389
'
'
'    Exercicios_Le_Todos = X.Exercicios_Le_Todos(colExercicios)
'
'    Exit Function
'
'Erro_Exercicios_Le_Todos:
'
'    Exercicios_Le_Todos = Err
'
'    Select Case Err
'
'        Case 9389
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144817)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Periodo_Le_Todos(colPeriodos As Collection) As Long
''le todos os periodos de todos os exercicios e coloca-os em colPeriodos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 9390
'
'
'    Periodo_Le_Todos = X.Periodo_Le_Todos(colPeriodos)
'
'    Exit Function
'
'Erro_Periodo_Le_Todos:
'
'    Periodo_Le_Todos = Err
'
'    Select Case Err
'
'        Case 9390
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144818)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Configuracao_Altera(objConfiguracao As ClassConfiguracao) As Long
''grava os registros do objeto na tabela Configuracao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Configuracao_Altera
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAO_ALTERA", X)
'    If lErro <> SUCESSO Then Error 14211
'
'
'    Configuracao_Altera = X.Configuracao_Altera(objConfiguracao)
'
'    Exit Function
'
'Erro_Configuracao_Altera:
'
'    Configuracao_Altera = Err
'
'    Select Case Err
'
'        Case 14211
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144819)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Configuracao_Altera_Trans(objConfiguracao As ClassConfiguracao) As Long
''grava os registros do objeto na tabela Configuracao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Configuracao_Altera_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAO_ALTERA_TRANS", X)
'    If lErro <> SUCESSO Then Error 20733
'
'
'    Configuracao_Altera_Trans = X.Configuracao_Altera_Trans(objConfiguracao)
'
'    Exit Function
'
'Erro_Configuracao_Altera_Trans:
'
'    Configuracao_Altera_Trans = Err
'
'    Select Case Err
'
'        Case 20733
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144820)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ConfiguracaoSetup_Grava(objConfiguracao As ClassConfiguracao) As Long
''Atualiza os seguintes campos da tabela Configuracao : LotePorPeriodo, DocPorPeriodo, UsoCcl, TipoContaDefault , OrigemDefault e NaturezaDefault
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ConfiguracaoSetup_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAOSETUP_GRAVA", X)
'    If lErro <> SUCESSO Then Error 14212
'
'
'    ConfiguracaoSetup_Grava = X.ConfiguracaoSetup_Grava(objConfiguracao)
'
'    Exit Function
'
'Erro_ConfiguracaoSetup_Grava:
'
'    ConfiguracaoSetup_Grava = Err
'
'    Select Case Err
'
'        Case 14212
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144821)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Atualizacao(sNomeArqParam As String, iID_Atualizacao As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Atualizacao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_ATUALIZACAO", X)
'    If lErro <> SUCESSO Then Error 20367
'
'
'    Rotina_Atualizacao = X.Rotina_Atualizacao(sNomeArqParam, iID_Atualizacao)
'
'    Exit Function
'
'Erro_Rotina_Atualizacao:
'
'    Rotina_Atualizacao = Err
'
'    Select Case Err
'
'        Case 20367
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144822)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Inicializacao(ByVal objExercicio As ClassExercicio, ByVal iPeriodo_Inicial As Integer, adtDataInicio() As Date, adtDataFim() As Date, asNomeExterno() As String, aiFechado() As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Inicializacao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_INICIALIZACAO", X)
'    If lErro <> SUCESSO Then Error 11402
'
'
'
'    Exercicio_Inicializacao = X.Exercicio_Inicializacao(objExercicio, iPeriodo_Inicial, adtDataInicio(), adtDataFim(), asNomeExterno(), aiFechado())
'
'    Exit Function
'
'Erro_Exercicio_Inicializacao:
'
'    Exercicio_Inicializacao = Err
'
'    Select Case Err
'
'        Case 11402
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144823)
'
'    End Select
'    Exit Function
'End Function
'
'Function Exercicio_Reinicializacao(ByVal objExercicio As ClassExercicio, ByVal iPeriodo_Inicial As Integer, adtDataInicio() As Date, adtDataFim() As Date, asNomeExterno() As String, aiFechado() As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Reinicializacao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_REINICIALIZACAO", X)
'    If lErro <> SUCESSO Then Error 11403
'
'
'
'    Exercicio_Reinicializacao = X.Exercicio_Reinicializacao(objExercicio, iPeriodo_Inicial, adtDataInicio(), adtDataFim(), asNomeExterno(), aiFechado())
'
'    Exit Function
'
'Erro_Exercicio_Reinicializacao:
'
'    Exercicio_Reinicializacao = Err
'
'    Select Case Err
'
'        Case 11403
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144824)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Orcamento_Le(objOrcamento As ClassOrcamento, colOrcamento As Collection) As Long
''le os registros da tabela Orcamento que satisfazem as condicoes passadas como parametro e coloca-os em colOrcamento
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Orcamento_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORCAMENTO_LE", X)
'    If lErro <> SUCESSO Then Error 14255
'
'
'
'    Orcamento_Le = X.Orcamento_Le(objOrcamento, colOrcamento)
'
'    Exit Function
'
'Erro_Orcamento_Le:
'
'    Orcamento_Le = Err
'
'    Select Case Err
'
'        Case 14255
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144825)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Orcamento_Exclui_BD(objOrcamento As ClassOrcamento) As Long
''exclui os registros da tabela de Orcamento para a conta x exercicio em questão
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Orcamento_Exclui_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORCAMENTO_EXCLUI_BD", X)
'    If lErro <> SUCESSO Then Error 14256
'
'
'
'    Orcamento_Exclui_BD = X.Orcamento_Exclui_BD(objOrcamento)
'
'    Exit Function
'
'Erro_Orcamento_Exclui_BD:
'
'    Orcamento_Exclui_BD = Err
'
'    Select Case Err
'
'        Case 14256
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144826)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Orcamento_Insere_BD(colOrcamento As Collection) As Long
''grava na tabela de Orcamento os novos registros do orcamento em questão
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Orcamento_Insere_BD
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORCAMENTO_INSERE_BD", X)
'    If lErro <> SUCESSO Then Error 14257
'
'
'
'    Orcamento_Insere_BD = X.Orcamento_Insere_BD(colOrcamento)
'
'    Exit Function
'
'Erro_Orcamento_Insere_BD:
'
'    Orcamento_Insere_BD = Err
'
'    Select Case Err
'
'        Case 14257
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144827)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Orcamento_Exclui(objOrcamento As ClassOrcamento) As Long
''Chama a rotina de exclusao de registros da tabela Orcamento, após obter a conta e o exercicio da tela
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Orcamento_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORCAMENTO_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 14258
'
'
'
'    Orcamento_Exclui = X.Orcamento_Exclui(objOrcamento)
'
'    Exit Function
'
'Erro_Orcamento_Exclui:
'
'    Orcamento_Exclui = Err
'
'    Select Case Err
'
'        Case 14258
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144828)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Orcamento_Grava(colOrcamento As Collection) As Long
''Grava os novos registros da tabela Orcamento, obtidos na tela de Orcamento
'
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Orcamento_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORCAMENTO_GRAVA", X)
'    If lErro <> SUCESSO Then Error 14259
'
'
'
'    Orcamento_Grava = X.Orcamento_Grava(colOrcamento)
'
'    Exit Function
'
'Erro_Orcamento_Grava:
'
'    Orcamento_Grava = Err
'
'    Select Case Err
'
'        Case 14259
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144829)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Apura_Exercicio(sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, sHistorico As String, sConta_Resultado As String, colContasApuracao As Collection) As Long
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Apura_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_APURA_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 20363
'
'
'    Rotina_Apura_Exercicio = X.Rotina_Apura_Exercicio(sNomeArqParam, iFilialEmpresa, iExercicio, iLote, sHistorico, sConta_Resultado, colContasApuracao)
'
'    Exit Function
'
'Erro_Rotina_Apura_Exercicio:
'
'    Rotina_Apura_Exercicio = Err
'
'    Select Case Err
'
'        Case 20363
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144830)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Apura_Periodos(sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo_Inicial As Integer, ByVal iPeriodo_Final As Integer, sContaResultado As String, sConta_Ponte As String, colContasApuracao As Collection, sHistorico As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Apura_Periodos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_APURA_PERIODOS", X)
'    If lErro <> SUCESSO Then Error 20364
'
'
'    Rotina_Apura_Periodos = X.Rotina_Apura_Periodos(sNomeArqParam, iFilialEmpresa, iExercicio, iPeriodo_Inicial, iPeriodo_Final, sContaResultado, sConta_Ponte, colContasApuracao, sHistorico)
'
'    Exit Function
'
'Erro_Rotina_Apura_Periodos:
'
'    Rotina_Apura_Periodos = Err
'
'    Select Case Err
'
'        Case 20364
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144831)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Rotina_Fechamento_Exercicio(sNomeArqParam As String, ByVal iExercicio As Integer, sConta_Ativo_Inicial As String, sConta_Ativo_Final As String, sConta_Passivo_Inicial As String, sConta_Passivo_Final As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Fechamento_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_FECHAMENTO_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 20365
'
'
'    Rotina_Fechamento_Exercicio = X.Rotina_Fechamento_Exercicio(sNomeArqParam, iExercicio, sConta_Ativo_Inicial, sConta_Ativo_Final, sConta_Passivo_Inicial, sConta_Passivo_Final)
'
'    Exit Function
'
'Erro_Rotina_Fechamento_Exercicio:
'
'    Rotina_Fechamento_Exercicio = Err
'
'    Select Case Err
'
'        Case 20365
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144832)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Rotina_Reabertura_Exercicio(sNomeArqParam As String, ByVal iExercicio As Integer, sConta_Ativo_Inicial As String, sConta_Ativo_Final As String, sConta_Passivo_Inicial As String, sConta_Passivo_Final As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Reabertura_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_REABERTURA_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 20366
'
'
'    Rotina_Reabertura_Exercicio = X.Rotina_Reabertura_Exercicio(sNomeArqParam, iExercicio, sConta_Ativo_Inicial, sConta_Ativo_Final, sConta_Passivo_Inicial, sConta_Passivo_Final)
'
'    Exit Function
'
'Erro_Rotina_Reabertura_Exercicio:
'
'    Rotina_Reabertura_Exercicio = Err
'
'    Select Case Err
'
'        Case 20366
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144833)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Rotina_Reprocessamento(sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Reprocessamento
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_REPROCESSAMENTO", X)
'    If lErro <> SUCESSO Then Error 20367
'
'
'    Rotina_Reprocessamento = X.Rotina_Reprocessamento(sNomeArqParam, iFilialEmpresa, iExercicio, iPeriodo)
'
'    Exit Function
'
'Erro_Rotina_Reprocessamento:
'
'    Rotina_Reprocessamento = Err
'
'    Select Case Err
'
'        Case 20367
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144834)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Atualiza_Lancamentos(objLancamento As ClassLancamento_Cabecalho) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Atualiza_Lancamentos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_ATUALIZA_LANCAMENTOS", X)
'    If lErro <> SUCESSO Then Error 9548
'
'
'    Rotina_Atualiza_Lancamentos = X.Rotina_Atualiza_Lancamentos(objLancamento)
'
'    Exit Function
'
'Erro_Rotina_Atualiza_Lancamentos:
'
'    Rotina_Atualiza_Lancamentos = Err
'
'    Select Case Err
'
'        Case 9548
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144835)
'
'
'    End Select
'
'    Exit Function
'
'
'End Function
'
'Function LoteApuracao_Automatico(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, iLote As Integer) As Long
''retorna o número do proximo lote disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LoteApuracao_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEAPURACAO_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 14305
'
'
'    LoteApuracao_Automatico = X.LoteApuracao_Automatico(iFilialEmpresa, iExercicio, iLote)
'
'    Exit Function
'
'Erro_LoteApuracao_Automatico:
'
'    LoteApuracao_Automatico = Err
'
'    Select Case Err
'
'        Case 14305
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144836)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_Prim_Aberto(objExercicio As ClassExercicio) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_Prim_Aberto
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_PRIM_ABERTO", X)
'    If lErro <> SUCESSO Then Error 14306
'
'
'    Exercicio_Le_Prim_Aberto = X.Exercicio_Le_Prim_Aberto(objExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Le_Prim_Aberto:
'
'    Exercicio_Le_Prim_Aberto = Err
'
'    Select Case Err
'
'        Case 14306
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144837)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_Ultimo_Fechado(objExercicio As ClassExercicio) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_Ultimo_Fechado
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_ULTIMO_FECHADO", X)
'    If lErro <> SUCESSO Then Error 14315
'
'
'    Exercicio_Le_Ultimo_Fechado = X.Exercicio_Le_Ultimo_Fechado(objExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Le_Ultimo_Fechado:
'
'    Exercicio_Le_Ultimo_Fechado = Err
'
'    Select Case Err
'
'        Case 14315
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144838)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le_Desatualizados(iFilialEmpresa As Integer, colLote As Collection) As Long
''Le todos os lotes com status = LOTE_DESATUALIZADO e IdAtualizacao = 0
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le_Desatualizados
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE_DESATUALIZADOS", X)
'    If lErro <> SUCESSO Then Error 14312
'
'
'    LotePendente_Le_Desatualizados = X.LotePendente_Le_Desatualizados(iFilialEmpresa, colLote)
'
'    Exit Function
'
'Erro_LotePendente_Le_Desatualizados:
'
'    LotePendente_Le_Desatualizados = Err
'
'    Select Case Err
'
'        Case 14312
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144839)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le_DesatualizadoII(iFilialEmpresa As Integer, colLote As Collection) As Long
''Le todos os lotes com status = LOTE_DESATUALIZADO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le_DesatualizadoII
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE_DESATUALIZADOII", X)
'    If lErro <> SUCESSO Then Error 14313
'
'
'    LotePendente_Le_DesatualizadoII = X.LotePendente_Le_DesatualizadoII(iFilialEmpresa, colLote)
'
'    Exit Function
'
'Erro_LotePendente_Le_DesatualizadoII:
'
'    LotePendente_Le_DesatualizadoII = Err
'
'    Select Case Err
'
'        Case 14313
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144840)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Atualiza(colLote As Collection, iIDAtualizacao As Integer) As Long
''atualiza o IdAtualizacao das tabelas Configuracao e LotePendente
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Atualiza
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_ATUALIZA", X)
'    If lErro <> SUCESSO Then Error 14314
'
'
'    LotePendente_Atualiza = X.LotePendente_Atualiza(colLote, iIDAtualizacao)
'
'    Exit Function
'
'Erro_LotePendente_Atualiza:
'
'    LotePendente_Atualiza = Err
'
'    Select Case Err
'
'        Case 14314
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144841)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCategoria_Le(objContaCategoria As ClassContaCategoria) As Long
''Le a categoria da conta a partir da chave fornecida em objContaCategoria. Retorna os dados em objContaCategoria
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCategoria_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACATEGORIA_LE", X)
'    If lErro <> SUCESSO Then Error 9652
'
'
'    ContaCategoria_Le = X.ContaCategoria_Le(objContaCategoria)
'
'    Exit Function
'
'Erro_ContaCategoria_Le:
'
'    ContaCategoria_Le = Err
'
'    Select Case Err
'
'        Case 9652
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144842)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoContas_Le_Categoria(iCodigo As Integer) As Long
''Le a categoria no plano de contas.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoContas_Le_Categoria
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTAS_LE_CATEGORIA", X)
'    If lErro <> SUCESSO Then Error 9686
'
'
'    PlanoContas_Le_Categoria = X.PlanoContas_Le_Categoria(iCodigo)
'
'    Exit Function
'
'Erro_PlanoContas_Le_Categoria:
'
'    PlanoContas_Le_Categoria = Err
'
'    Select Case Err
'
'        Case 9686
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144843)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCategoria_Le_Todos(colContaCategoria As Collection) As Long
''le todos as categorias contidas na tabela de categorias e coloca-as em colContaCategoria
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCategoria_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACATEGORIA_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 9687
'
'
'    ContaCategoria_Le_Todos = X.ContaCategoria_Le_Todos(colContaCategoria)
'
'    Exit Function
'
'Erro_ContaCategoria_Le_Todos:
'
'    ContaCategoria_Le_Todos = Err
'
'    Select Case Err
'
'        Case 9687
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144844)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCategoria_Grava(objContaCategoria As ClassContaCategoria) As Long
''grava a Categoria no banco de dados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCategoria_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACATEGORIA_GRAVA", X)
'    If lErro <> SUCESSO Then Error 9696
'
'
'    ContaCategoria_Grava = X.ContaCategoria_Grava(objContaCategoria)
'
'    Exit Function
'
'Erro_ContaCategoria_Grava:
'
'    ContaCategoria_Grava = Err
'
'    Select Case Err
'
'        Case 9696
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144845)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCategoria_Exclui(iCodigo As Integer) As Long
''exclui a Categoria
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCategoria_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACATEGORIA_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 9697
'
'
'    ContaCategoria_Exclui = X.ContaCategoria_Exclui(iCodigo)
'
'    Exit Function
'
'Erro_ContaCategoria_Exclui:
'
'    ContaCategoria_Exclui = Err
'
'    Select Case Err
'
'        Case 9697
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144846)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCategoria_Automatico(iCodigo As Integer) As Long
''retorna o proximo codigo de Categoria disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCategoria_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACATEGORIA_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 9720
'
'
'    ContaCategoria_Automatico = X.ContaCategoria_Automatico(iCodigo)
'
'    Exit Function
'
'Erro_ContaCategoria_Automatico:
'
'    ContaCategoria_Automatico = Err
'
'    Select Case Err
'
'        Case 9720
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144847)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Categoria_Nivel(iCategoria As Integer, iNivelConta As Integer, objPlanoConta As ClassPlanoConta) As Long
''Le a conta na categoria e no nivel especificados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Categoria_Nivel
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_CATEGORIA_NIVEL", X)
'    If lErro <> SUCESSO Then Error 9728
'
'
'    PlanoConta_Le_Categoria_Nivel = X.PlanoConta_Le_Categoria_Nivel(iCategoria, iNivelConta, objPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Categoria_Nivel:
'
'    PlanoConta_Le_Categoria_Nivel = Err
'
'    Select Case Err
'
'        Case 9728
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144848)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCategoria_Le_Nome(sNome As String, objContaCategoria As ClassContaCategoria) As Long
''Le a categoria da conta que possui o sNome especificado. Retorna os dados em objContaCategoria
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCategoria_Le_Nome
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACATEGORIA_LE_NOME", X)
'    If lErro <> SUCESSO Then Error 9729
'
'
'    ContaCategoria_Le_Nome = X.ContaCategoria_Le_Nome(sNome, objContaCategoria)
'
'    Exit Function
'
'Erro_ContaCategoria_Le_Nome:
'
'    ContaCategoria_Le_Nome = Err
'
'    Select Case Err
'
'        Case 9729
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144849)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Todas_Categorias(colPlanoConta As Collection) As Long
''le todas as contas que possuam categoria no plano de contas
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Todas_Categorias
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_TODAS_CATEGORIAS", X)
'    If lErro <> SUCESSO Then Error 9749
'
'
'    PlanoConta_Le_Todas_Categorias = X.PlanoConta_Le_Todas_Categorias(colPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Todas_Categorias:
'
'    PlanoConta_Le_Todas_Categorias = Err
'
'    Select Case Err
'
'        Case 9749
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144850)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CTBConfig_Le(objCTBConfig As ClassCTBConfig) As Long
''Le o registro de configuração especificado no codigo de objCTBConfig e retorna os demais dados em objCTBConfig
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CTBConfig_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CTBCONFIG_LE", X)
'    If lErro <> SUCESSO Then Error 9756
'
'
'    CTBConfig_Le = X.CTBConfig_Le(objCTBConfig)
'
'    Exit Function
'
'Erro_CTBConfig_Le:
'
'    CTBConfig_Le = Err
'
'    Select Case Err
'
'        Case 9756
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144851)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function DRE_MontaArqTemp(i_Exercicio As Integer, i_NivelMax As Integer, lCodIdentExec As Long) As Long
''monta arquivo temporario auxiliar a execucao do relatorio "Demonstrativo do Resultado do Exercicio"
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DRE_MontaArqTemp
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DRE_MONTA_ARQ_TEMP", X)
'    If lErro <> SUCESSO Then Error 7119
'
'
'    DRE_MontaArqTemp = X.MontaArqTemp(i_Exercicio, i_NivelMax, lCodIdentExec)
'
'    Exit Function
'
'Erro_DRE_MontaArqTemp:
'
'    DRE_MontaArqTemp = Err
'
'    Select Case Err
'
'        Case 7119
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144852)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function DRP_MontaArqTemp(i_Exercicio As Integer, i_Periodo As Integer, i_NivelMax As Integer, lCodIdentExec As Long) As Long
''monta arquivo temporario auxiliar a execucao do relatorio "Demonstrativo do Resultado do Periodo"
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_DRP_MontaArqTemp
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_DRP_MONTA_ARQ_TEMP", X)
'    If lErro <> SUCESSO Then Error 7118
'
'
'    DRP_MontaArqTemp = X.MontaArqTemp(i_Exercicio, i_Periodo, i_NivelMax, lCodIdentExec)
'
'    Exit Function
'
'Erro_DRP_MontaArqTemp:
'
'    DRP_MontaArqTemp = Err
'
'    Select Case Err
'
'        Case 7118
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144853)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CTBConfig_ObterProxLong(sChave As String, lValor As Long) As Long
''obtem o valor correspondente a sChave (do tipo long), retorna-o em lValor e incrementa o contador na tabela CTBconfig
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CTBConfig_ObterProxLong
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIG_OBTER_PROX_LONG", X)
'    If lErro <> SUCESSO Then Error 7117
'
'
'    CTBConfig_ObterProxLong = X.CTBConfig_ObterProxLong(sChave, lValor)
'
'    Exit Function
'
'Erro_CTBConfig_ObterProxLong:
'
'    CTBConfig_ObterProxLong = Err
'
'    Select Case Err
'
'        Case 7117
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144854)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_Primeiro(objExercicio As ClassExercicio) As Long
''retorna o primeiro exercicio do sistema
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_Primeiro
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_PRIMEIRO", X)
'    If lErro <> SUCESSO Then Error 9848
'
'
'    Exercicio_Le_Primeiro = X.Exercicio_Le_Primeiro(objExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Le_Primeiro:
'
'    Exercicio_Le_Primeiro = Err
'
'    Select Case Err
'
'        Case 9848
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144855)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_NumRegistros(lNumReg As Long) As Long
''retorna o número de registros contidos na tabela de Exercicio
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_NumRegistros
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_NUMREGISTROS", X)
'    If lErro <> SUCESSO Then Error 9861
'
'
'    Exercicio_Le_NumRegistros = X.Exercicio_Le_NumRegistros(lNumReg)
'
'    Exit Function
'
'Erro_Exercicio_Le_NumRegistros:
'
'    Exercicio_Le_NumRegistros = Err
'
'    Select Case Err
'
'        Case 9861
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144856)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_ContaCcl(sConta As String, sCcl As String) As Long
''verifica se existe no BD pelo menos um lancamento pendente para a conta e centro de custo/lucro fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_ContaCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_CONTACCL", X)
'    If lErro <> SUCESSO Then Error 9895
'
'
'    LanPendente_Le_ContaCcl = X.LanPendente_Le_ContaCcl(sConta, sCcl)
'
'    Exit Function
'
'Erro_LanPendente_Le_ContaCcl:
'
'    LanPendente_Le_ContaCcl = Err
'
'    Select Case Err
'
'        Case 9895
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144857)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCcl_Critica_Possui_LanPendente(sConta As String, sCcl As String, iTemLanPendente As Integer) As Long
''Verifica se existem lançamentos pendentes para a conta e centro de custo extra contabil em questão. Se existir retorna iTemLanPendente = 1, senao iTemLanPendente = 0
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Critica_Possui_LanPendente
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_CRITICA_POSSUI_LANPENDENTE", X)
'    If lErro <> SUCESSO Then Error 9897
'
'
'    ContaCcl_Critica_Possui_LanPendente = X.ContaCcl_Critica_Possui_LanPendente(sConta, sCcl, iTemLanPendente)
'
'    Exit Function
'
'Erro_ContaCcl_Critica_Possui_LanPendente:
'
'    ContaCcl_Critica_Possui_LanPendente = Err
'
'    Select Case Err
'
'        Case 9897
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144858)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOn_Le_ContaCcl(sConta As String, sCcl As String) As Long
''verifica se existe na tabela RateioOn pelo menos um rateio com a conta e centro de custo em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOn_Le_ContaCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOON_LE_CONTACCL", X)
'    If lErro <> SUCESSO Then Error 9904
'
'
'    RateioOn_Le_ContaCcl = X.RateioOn_Le_ContaCcl(sConta, sCcl)
'
'    Exit Function
'
'Erro_RateioOn_Le_ContaCcl:
'
'    RateioOn_Le_ContaCcl = Err
'
'    Select Case Err
'
'        Case 9904
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144859)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaCcl_Critica_Exclusao(sConta As String, sCcl As String) As Long
''verifica se a conta+ccl possui movimento, lancamentos pendentes e rateioon associado.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Critica_Exclusao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_CRITICA_EXCLUSAO", X)
'    If lErro <> SUCESSO Then Error 9910
'
'
'    ContaCcl_Critica_Exclusao = X.ContaCcl_Critica_Exclusao(sConta, sCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Critica_Exclusao:
'
'    ContaCcl_Critica_Exclusao = Err
'
'    Select Case Err
'
'        Case 9910
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144860)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicios_Nao_Fechados_Le_Todos(colExercicios As Collection) As Long
''le todos os exercicios não fechados e coloca-os em colExercicios
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicios_Nao_Fechados_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIOS_NAO_FECHADOS_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 9931
'
'
'    Exercicios_Nao_Fechados_Le_Todos = X.Exercicios_Nao_Fechados_Le_Todos(colExercicios)
'
'    Exit Function
'
'Erro_Exercicios_Nao_Fechados_Le_Todos:
'
'    Exercicios_Nao_Fechados_Le_Todos = Err
'
'    Select Case Err
'
'        Case 9931
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144861)
'
'
'    End Select
'
'    Exit Function
'
'
'End Function
'
'Public Function ContaCcl_Le_Todos_Conta1(sConta As String, colContaCcl As Collection) As Long
''Le todas as associacoes da conta em questão com centros de custo/lucro e carrega-os em colContaCcl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le_Todos_Conta1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE_TODOS_CONTA1", X)
'    If lErro <> SUCESSO Then Error 9947
'
'
'    ContaCcl_Le_Todos_Conta1 = X.ContaCcl_Le_Todos_Conta1(sConta, colContaCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Le_Todos_Conta1:
'
'    ContaCcl_Le_Todos_Conta1 = Err
'
'    Select Case Err
'
'        Case 9947
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144862)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Le_Todas_Contas_Distintas(colConta As Collection) As Long
''Le todas as contas que possuem associação com centro de custo e coloca-as em colConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Le_Todas_Contas_Distintas
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_LE_TODAS_CONTAS_DISTINTAS", X)
'    If lErro <> SUCESSO Then Error 9954
'
'
'    ContaCcl_Le_Todas_Contas_Distintas = X.ContaCcl_Le_Todas_Contas_Distintas(colConta)
'
'    Exit Function
'
'Erro_ContaCcl_Le_Todas_Contas_Distintas:
'
'    ContaCcl_Le_Todas_Contas_Distintas = Err
'
'    Select Case Err
'
'        Case 9954
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144863)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Atualiza_Saldo_Contabil(colSaldoInicialContaCcl As Collection) As Long
''atualiza o saldo inicial de uma coleção de contas com centro de custo contábil
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Atualiza_Saldo_Contabil
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_ATUALIZA_SALDO_CONTABIL", X)
'    If lErro <> SUCESSO Then Error 10032
'
'
'    ContaCcl_Atualiza_Saldo_Contabil = X.ContaCcl_Atualiza_Saldo_Contabil(colSaldoInicialContaCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Atualiza_Saldo_Contabil:
'
'    ContaCcl_Atualiza_Saldo_Contabil = Err
'
'    Select Case Err
'
'        Case 10032
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144864)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function ContaCcl_Atualiza_Saldo_Extra(colContaCcl As Collection) As Long
''atualiza o saldo inicial de uma coleção de contas com centro de custo extra contábil
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaCcl_Atualiza_Saldo_Extra
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTACCL_ATUALIZA_SALDO_EXTRA", X)
'    If lErro <> SUCESSO Then Error 10033
'
'
'    ContaCcl_Atualiza_Saldo_Extra = X.ContaCcl_Atualiza_Saldo_Extra(colContaCcl)
'
'    Exit Function
'
'Erro_ContaCcl_Atualiza_Saldo_Extra:
'
'    ContaCcl_Atualiza_Saldo_Extra = Err
'
'    Select Case Err
'
'        Case 10033
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144865)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Conta1(sConta As String, objPlanoConta As ClassPlanoConta) As Long
''O argumento sConta deve ser o número da conta FORMATADO BD
''le a conta passada como parametro e retorna os dados em objPlanoConta.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Conta1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_CONTA1", X)
'    If lErro <> SUCESSO Then Error 10048
'
'
'    PlanoConta_Le_Conta1 = X.PlanoConta_Le_Conta1(sConta, objPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Conta1:
'
'    PlanoConta_Le_Conta1 = Err
'
'    Select Case Err
'
'        Case 10048
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144866)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Todos_Modulo(colAnaliticas As Collection, colSinteticas As Collection, sSiglaModulo As String) As Long
''le todas as contas analíticas, visiveis para o modulo em questão, e as conta sinteticas contidas na tabela de Plano de Contas e coloca-as em colAnaliticas (contas analiticas) e colSinteticas (contas sinteticas)
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Todos_Modulo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_TODOS_MODULO", X)
'    If lErro <> SUCESSO Then Error 10060
'
'
'    PlanoConta_Le_Todos_Modulo = X.PlanoConta_Le_Todos_Modulo(colAnaliticas, colSinteticas, sSiglaModulo)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Todos_Modulo:
'
'    PlanoConta_Le_Todos_Modulo = Err
'
'    Select Case Err
'
'        Case 10060
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144867)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carga_Arvore_Conta_Modulo(colNodes As Nodes, sSiglaModulo As String) As Long
''move os dados do plano de contas, relativos ao modulo em questão,  do banco de dados para a arvore colNodes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Carga_Arvore_Conta_Modulo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CARGA_ARVORE_CONTA_MODULO", X)
'    If lErro <> SUCESSO Then Error 10072
'
'
'    Carga_Arvore_Conta_Modulo = X.Carga_Arvore_Conta_Modulo(colNodes, sSiglaModulo)
'
'    Exit Function
'
'Erro_Carga_Arvore_Conta_Modulo:
'
'    Carga_Arvore_Conta_Modulo = Err
'
'    Select Case Err
'
'        Case 10072
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144868)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContasPatrimoniaisNivelUm_Le(colXContas As Collection) As Long
''preenche uma colecao com os codigos das contas patrimoniais de nivel 1
''p/evitar problemas o tipo de dados Variant vs cta que comeca por numererico vou
''colocar todas as ctas precedidas pela letra 'X'. A colecao deverá conter portanto
''elementos como "X1000000" e "X2000000"
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContasPatrimoniaisNivelUm_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTASPATRIMONIAISNIVELUM_LE", X)
'    If lErro <> SUCESSO Then Error 14504
'
'
'    ContasPatrimoniaisNivelUm_Le = X.ContasPatrimoniaisNivelUm_Le(colXContas)
'
'    Exit Function
'
'Erro_ContasPatrimoniaisNivelUm_Le:
'
'    ContasPatrimoniaisNivelUm_Le = Err
'
'    Select Case Err
'
'        Case 14504
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144869)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Monta_Filtro_Ctas_Patrimoniais(objRelOpcoes As AdmRelOpcoes, sExprCtasPatr As String) As Long
''monta expressao do tipo "GP(@TC1,@TC2) E GP(@TC3,@TC4)"
''onde GP é uima macroa com a expressao: Conta >= $T0 E Conta <= $T1
''esta funcao vai incluir tambem os valores dos @T? que criar
''Os valores de T1, T2, T3, T4 serao algo como: "1000000", "19999999", "20000000", "29999999"
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Monta_Filtro_Ctas_Patrimoniais
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MONTA_FILTRO_CTAS_PATRIMONIAIS", X)
'    If lErro <> SUCESSO Then Error 14505
'
'    Monta_Filtro_Ctas_Patrimoniais = X.Monta_Filtro_Ctas_Patrimoniais(objRelOpcoes, sExprCtasPatr)
'
'    Exit Function
'
'Erro_Monta_Filtro_Ctas_Patrimoniais:
'
'    Monta_Filtro_Ctas_Patrimoniais = Err
'
'    Select Case Err
'
'        Case 14505
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144870)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Grava(objExercicio As ClassExercicio, colPeriodos As Collection) As Long
''grava os dados do exercicio objExercicio.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_GRAVA", X)
'    If lErro <> SUCESSO Then Error 10120
'
'
'    Exercicio_Grava = X.Exercicio_Grava(objExercicio, colPeriodos)
'
'    Exit Function
'
'Erro_Exercicio_Grava:
'
'    Exercicio_Grava = Err
'
'    Select Case Err
'
'        Case 10120
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144871)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Orcamento_Le_Exercicio(iExercicio As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Orcamento_Le_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ORCAMENTO_LE_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 10121
'
'
'
'    Orcamento_Le_Exercicio = X.Orcamento_Le_Exercicio(iExercicio)
'
'    Exit Function
'
'Erro_Orcamento_Le_Exercicio:
'
'    Orcamento_Le_Exercicio = Err
'
'    Select Case Err
'
'        Case 10121
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144872)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Movimento_Le_Exercicio(iExercicio As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Movimento_Le_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MOVIMENTO_LE_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 10122
'
'
'
'    Movimento_Le_Exercicio = X.Movimento_Le_Exercicio(iExercicio)
'
'    Exit Function
'
'Erro_Movimento_Le_Exercicio:
'
'    Movimento_Le_Exercicio = Err
'
'    Select Case Err
'
'        Case 10122
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144873)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_Exercicio(iExercicio As Integer) As Long
''verifica se existe no BD pelo menos um lancamento pendente para o exercício fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 10123
'
'
'
'    LanPendente_Le_Exercicio = X.LanPendente_Le_Exercicio(iExercicio)
'
'    Exit Function
'
'Erro_LanPendente_Le_Exercicio:
'
'    LanPendente_Le_Exercicio = Err
'
'    Select Case Err
'
'        Case 10123
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144874)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le_Exercicio(iExercicio As Integer) As Long
''verifica se existe no BD pelo menos um lote pendente para o exercício fornecido
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le_Exercicio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE_EXERCICIO", X)
'    If lErro <> SUCESSO Then Error 10124
'
'
'
'    LotePendente_Le_Exercicio = X.LotePendente_Le_Exercicio(iExercicio)
'
'    Exit Function
'
'Erro_LotePendente_Le_Exercicio:
'
'    LotePendente_Le_Exercicio = Err
'
'    Select Case Err
'
'        Case 10124
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144875)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Critica_Movimento(iExercicio As Integer) As Long
''Verifica se existe movimento, lote ou lançamento pendente no exercício iExercicio
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Critica_Movimento
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_CRITICA_MOVIMENTO", X)
'    If lErro <> SUCESSO Then Error 10126
'
'
'
'    Exercicio_Critica_Movimento = X.Exercicio_Critica_Movimento(iExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Critica_Movimento:
'
'    Exercicio_Critica_Movimento = Err
'
'    Select Case Err
'
'        Case 10126
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144876)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Exclui(iExercicio As Integer) As Long
''exclui o exercicio em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 10133
'
'
'
'    Exercicio_Exclui = X.Exercicio_Exclui(iExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Exclui:
'
'    Exercicio_Exclui = Err
'
'    Select Case Err
'
'        Case 10133
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144877)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PeriodosFilial_Le(objPeriodosFilial As ClassPeriodosFilial) As Long
''le o registro referente ao periodo em questão e coloca os dados em objPeriodosFilial
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PeriodosFilial_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODOSFILIAL_LE", X)
'    If lErro <> SUCESSO Then Error 10160
'
'
'
'    PeriodosFilial_Le = X.PeriodosFilial_Le(objPeriodosFilial)
'
'    Exit Function
'
'Erro_PeriodosFilial_Le:
'
'    PeriodosFilial_Le = Err
'
'    Select Case Err
'
'        Case 10160
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144878)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ExerciciosFilial_Grava(objExercicio As ClassExercicio, colPeriodos As Collection) As Long
''grava os dados do exercicio/filial contido em objExercicio
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ExerciciosFilial_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIOSFILIAL_GRAVA", X)
'    If lErro <> SUCESSO Then Error 10251
'
'
'
'    ExerciciosFilial_Grava = X.ExerciciosFilial_Grava(objExercicio, colPeriodos)
'
'    Exit Function
'
'Erro_ExerciciosFilial_Grava:
'
'    ExerciciosFilial_Grava = Err
'
'    Select Case Err
'
'        Case 10251
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144879)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function SaldoInicialConta_Le(objSaldoInicialConta As ClassSaldoInicialConta) As Long
''Le o saldo inicial da conta passada como parametro e devolve o resultado em objPlanoConta.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SaldoInicialConta_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SALDOINICIALCONTA_LE", X)
'    If lErro <> SUCESSO Then Error 10346
'
'
'
'    SaldoInicialConta_Le = X.SaldoInicialConta_Le(objSaldoInicialConta)
'
'    Exit Function
'
'Erro_SaldoInicialConta_Le:
'
'    SaldoInicialConta_Le = Err
'
'    Select Case Err
'
'        Case 10346
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144880)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function SaldoInicialContaCcl_Le_Todos_Conta(objSaldoInicialContaCcl As ClassSaldoInicialContaCcl, colSaldoInicialContaCcl As Collection) As Long
''Le todas as associacoes da conta em questão com centros de custo/lucro e carrega-os em colContaCcl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SaldoInicialContaCcl_Le_Todos_Conta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SALDOINICIALCONTACCL_LE_TODOS_CONTA", X)
'    If lErro <> SUCESSO Then Error 10351
'
'
'
'    SaldoInicialContaCcl_Le_Todos_Conta = X.SaldoInicialContaCcl_Le_Todos_Conta(objSaldoInicialContaCcl, colSaldoInicialContaCcl)
'
'    Exit Function
'
'Erro_SaldoInicialContaCcl_Le_Todos_Conta:
'
'    SaldoInicialContaCcl_Le_Todos_Conta = Err
'
'    Select Case Err
'
'        Case 10351
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144881)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CclContabil_Critica_Exclusao(sCcl As String) As Long
''verifica se o Ccl possui rateios associados, etc.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CclContabil_Critica_Exclusao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCLCONTABIL_CRITICA_EXCLUSAO", X)
'    If lErro <> SUCESSO Then Error 10402
'
'
'
'    CclContabil_Critica_Exclusao = X.CclContabil_Critica_Exclusao(sCcl)
'
'    Exit Function
'
'Erro_CclContabil_Critica_Exclusao:
'
'    CclContabil_Critica_Exclusao = Err
'
'    Select Case Err
'
'        Case 10402
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144882)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CclExtraContabil_Critica_Exclusao(sCcl As String) As Long
''verifica se o Ccl possui rateios associados, etc.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CclExtraContabil_Critica_Exclusao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCLEXTRACONTABIL_CRITICA_EXCLUSAO", X)
'    If lErro <> SUCESSO Then Error 10427
'
'
'
'    CclExtraContabil_Critica_Exclusao = X.CclExtraContabil_Critica_Exclusao(sCcl)
'
'    Exit Function
'
'Erro_CclExtraContabil_Critica_Exclusao:
'
'    CclExtraContabil_Critica_Exclusao = Err
'
'    Select Case Err
'
'        Case 10427
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144883)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Critica_CclPai(sCcl As String) As Long
''verifica se o centro de custo/lucro tem um pai sintetico
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica_CclPai
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA_CCLPAI", X)
'    If lErro <> SUCESSO Then Error 10462
'
'
'
'    Ccl_Critica_CclPai = X.Ccl_Critica_CclPai(sCcl)
'
'    Exit Function
'
'Erro_Ccl_Critica_CclPai:
'
'    Ccl_Critica_CclPai = Err
'
'    Select Case Err
'
'        Case 10462
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144884)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Le_Todos_Analiticos(colCcl As Collection) As Long
''le todos os centros de custo/lucro analiticos contidos na tabela Ccl e coloca-os em colCcl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Le_Todos_Analiticos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_LE_TODOS_ANALITICOS", X)
'    If lErro <> SUCESSO Then Error 10482
'
'
'
'    Ccl_Le_Todos_Analiticos = X.Ccl_Le_Todos_Analiticos(colCcl)
'
'    Exit Function
'
'Erro_Ccl_Le_Todos_Analiticos:
'
'    Ccl_Le_Todos_Analiticos = Err
'
'    Select Case Err
'
'        Case 10482
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144885)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Critica_Analitica(sCcl As String) As Long
''verifica se o centro de custo é analitico
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica_Analitica
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA_ANALITICA", X)
'    If lErro <> SUCESSO Then Error 10495
'
'
'
'    Ccl_Critica_Analitica = X.Ccl_Critica_Analitica(sCcl)
'
'    Exit Function
'
'Erro_Ccl_Critica_Analitica:
'
'    Ccl_Critica_Analitica = Err
'
'    Select Case Err
'
'        Case 10495
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144886)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Critica_Tipo(objCcl As ClassCcl) As Long
''critica o tipo do centro de custo/lucro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Critica_Tipo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_CRITICA_TIPO", X)
'    If lErro <> SUCESSO Then Error 10580
'
'
'
'    Ccl_Critica_Tipo = X.Ccl_Critica_Tipo(objCcl)
'
'    Exit Function
'
'Erro_Ccl_Critica_Tipo:
'
'    Ccl_Critica_Tipo = Err
'
'    Select Case Err
'
'        Case 10580
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144887)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Conta_Critica_Tipo(sConta As String, iTipoConta As Integer) As Long
''critica o tipo da conta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Critica_Tipo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_CRITICA_TIPO", X)
'    If lErro <> SUCESSO Then Error 10581
'
'
'
'    Conta_Critica_Tipo = X.Conta_Critica_Tipo(sConta, iTipoConta)
'
'    Exit Function
'
'Erro_Conta_Critica_Tipo:
'
'    Conta_Critica_Tipo = Err
'
'    Select Case Err
'
'        Case 10581
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144888)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Le_Codigo(objExercicio As ClassExercicio) As Long
''Lê o código do Exercício à partir do nome externoo passado como parâmetro. Retorna os dados em objExercício
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Le_Codigo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_LE_CODIGO", X)
'    If lErro <> SUCESSO Then Error 24812
'
'
'
'    Exercicio_Le_Codigo = X.Exercicio_Le_Codigo(objExercicio)
'
'    Exit Function
'
'Erro_Exercicio_Le_Codigo:
'
'    Exercicio_Le_Codigo = Err
'
'    Select Case Err
'
'        Case 24812
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144889)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Periodo_Le_Codigo(objPeriodo As ClassPeriodo) As Long
''Lê o código do Período à partir do nome externo e o exercício passados como parâmetro. Retorna os dados em objPeriodo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le_Codigo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE_CODIGO", X)
'    If lErro <> SUCESSO Then Error 24813
'
'
'
'    Periodo_Le_Codigo = X.Periodo_Le_Codigo(objPeriodo)
'
'    Exit Function
'
'Erro_Periodo_Le_Codigo:
'
'    Periodo_Le_Codigo = Err
'
'    Select Case Err
'
'        Case 24813
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144890)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ExerciciosFilial_Le(objExerciciosFilial As ClassExerciciosFilial) As Long
''retorna os dados do exerciciofilial em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ExerciciosFilial_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIOSFILIAL_LE", X)
'    If lErro <> SUCESSO Then Error 24916
'
'
'
'    ExerciciosFilial_Le = X.ExerciciosFilial_Le(objExerciciosFilial)
'
'    Exit Function
'
'Erro_ExerciciosFilial_Le:
'
'    ExerciciosFilial_Le = Err
'
'    Select Case Err
'
'        Case 24916
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144891)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContabItem_Le(objPadraoContab As ClassPadraoContab, colPadraoContabItem As ClassColPadraoContabItem) As Long
''retorna os lançamentos padrao de objPadraoContab. Os lançamentos são retornados em colPadraoContabItem
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContabItem_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTABITEM_LE", X)
'    If lErro <> SUCESSO Then Error 36055
'
'
'
'    PadraoContabItem_Le = X.PadraoContabItem_Le(objPadraoContab, colPadraoContabItem)
'
'    Exit Function
'
'Erro_PadraoContabItem_Le:
'
'    PadraoContabItem_Le = Err
'
'    Select Case Err
'
'        Case 36055
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144892)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MnemonicoCTB_Le(sModulo As String, iTransacao As Integer, colMnemonico As Collection) As Long
''retorna os mnemonicos associados ao modulo/transacao passados como parametro. Os mnemonicos são retornados em colMnemonico
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MnemonicoCTB_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MNEMONICOCTB_LE", X)
'    If lErro <> SUCESSO Then Error 36056
'
'
'
'    MnemonicoCTB_Le = X.MnemonicoCTB_Le(sModulo, iTransacao, colMnemonico)
'
'    Exit Function
'
'Erro_MnemonicoCTB_Le:
'
'    MnemonicoCTB_Le = Err
'
'    Select Case Err
'
'        Case 36056
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144893)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContab_Le(objPadraoContab As ClassPadraoContab) As Long
''retorna o modelo de geracao padrao deste Modulo/Transacao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContab_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_LE", X)
'    If lErro <> SUCESSO Then Error 36057
'
'
'
'    PadraoContab_Le = X.PadraoContab_Le(objPadraoContab)
'
'    Exit Function
'
'Erro_PadraoContab_Le:
'
'    PadraoContab_Le = Err
'
'    Select Case Err
'
'        Case 36057
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144894)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Gera_Lancamentos(sModulo As String, objTransacao As Object, sModelo As String, colLancamentos As ClassColLancamentos, iSubTipoTransacao As Integer) As Long
''retorna os lançamentos do Modulo/Transacao passados como parametro. Os lançamentos são retornados em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Gera_Lancamentos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_GERALANCAMENTOS", X)
'    If lErro <> SUCESSO Then Error 36058
'
'
'
'    Gera_Lancamentos = X.Gera_Lancamentos(sModulo, objTransacao, sModelo, colLancamentos, iSubTipoTransacao)
'
'    Exit Function
'
'Erro_Gera_Lancamentos:
'
'    Gera_Lancamentos = Err
'
'    Select Case Err
'
'        Case 36058
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144895)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Valida_Lancamentos_Padrao(sModulo As String, objTransacao As Object) As Long
''valida os lançamentos do Modulo/Transacao passados como parametro.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Valida_Lancamentos_Padrao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VALIDA_LANCAMENTOS_PADRAO", X)
'    If lErro <> SUCESSO Then Error 36059
'
'
'
'    Valida_Lancamentos_Padrao = X.Valida_Lancamentos_Padrao(sModulo, objTransacao)
'
'    Exit Function
'
'Erro_Valida_Lancamentos_Padrao:
'
'    Valida_Lancamentos_Padrao = Err
'
'    Select Case Err
'
'        Case 36059
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144896)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Valida_Formula(sModulo As String, iTransacao As Integer, sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer) As Long
''compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Valida_Formula
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VALIDA_FORMULA", X)
'    If lErro <> SUCESSO Then Error 36137
'
'
'
'    Valida_Formula = X.Valida_Formula(sModulo, iTransacao, sFormula, iTipo, iInicio, iTam)
'
'    Exit Function
'
'Erro_Valida_Formula:
'
'    Valida_Formula = Err
'
'    Select Case Err
'
'        Case 36137
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144897)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MnemonicoCTB_Le_Mnemonico(objMnemonico As ClassMnemonico) As Long
''retorna os dados do mnemonico passado como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MnemonicoCTB_Le_Mnemonico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MNEMONICOCTB_LE_MNEMONICO", X)
'    If lErro <> SUCESSO Then Error 36138
'
'
'
'    MnemonicoCTB_Le_Mnemonico = X.MnemonicoCTB_Le_Mnemonico(objMnemonico)
'
'    Exit Function
'
'Erro_MnemonicoCTB_Le_Mnemonico:
'
'    MnemonicoCTB_Le_Mnemonico = Err
'
'    Select Case Err
'
'        Case 36138
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144898)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function FormulaFuncao_Le_Todos(colFormulaFuncao As Collection) As Long
''Le todas as funções e carrega-as em colFormulaFuncao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FormulaFuncao_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_FORMULAFUNCAO_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 36139
'
'
'
'    FormulaFuncao_Le_Todos = X.FormulaFuncao_Le_Todos(colFormulaFuncao)
'
'    Exit Function
'
'Erro_FormulaFuncao_Le_Todos:
'
'    FormulaFuncao_Le_Todos = Err
'
'    Select Case Err
'
'        Case 36139
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144899)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function FormulaFuncao_Le(objFormulaFuncao As ClassFormulaFuncao) As Long
''Le os dados da função passada como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FormulaFuncao_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_FORMULAFUNCAO_LE", X)
'    If lErro <> SUCESSO Then Error 36140
'
'
'
'    FormulaFuncao_Le = X.FormulaFuncao_Le(objFormulaFuncao)
'
'    Exit Function
'
'Erro_FormulaFuncao_Le:
'
'    FormulaFuncao_Le = Err
'
'    Select Case Err
'
'        Case 36140
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144900)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function FormulaOperador_Le_Todos(colFormulaOperador As Collection) As Long
''Le todos os operadores e carrega-as em colFormulaOperador
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FormulaOperador_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_FORMULAOPERADOR_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 36141
'
'
'
'    FormulaOperador_Le_Todos = X.FormulaOperador_Le_Todos(colFormulaOperador)
'
'    Exit Function
'
'Erro_FormulaOperador_Le_Todos:
'
'    FormulaOperador_Le_Todos = Err
'
'    Select Case Err
'
'        Case 36141
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144901)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function FormulaOperador_Le(objFormulaOperador As ClassFormulaOperador) As Long
''Le os dados do operador passado como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FormulaOperador_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_FORMULAOPERADOR_LE", X)
'    If lErro <> SUCESSO Then Error 36142
'
'
'
'    FormulaOperador_Le = X.FormulaOperador_Le(objFormulaOperador)
'
'    Exit Function
'
'Erro_FormulaOperador_Le:
'
'    FormulaOperador_Le = Err
'
'    Select Case Err
'
'        Case 36142
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144902)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TransacaoCTB_Le_Todos(sSiglaModulo As String, colTransacao As Collection) As Long
''Le todas as transacoes relativas ao modulo em questão e carrega-as em colTransacao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_TransacaoCTB_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_TRANSACAOTCTB_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 36143
'
'
'
'    TransacaoCTB_Le_Todos = X.TransacaoCTB_Le_Todos(sSiglaModulo, colTransacao)
'
'    Exit Function
'
'Erro_TransacaoCTB_Le_Todos:
'
'    TransacaoCTB_Le_Todos = Err
'
'    Select Case Err
'
'        Case 36143
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144903)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContab_Le_Todos(sSiglaModulo As String, iTransacao As Integer, colPadraoContab As Collection, iSubtipo As Integer) As Long
''leitura dos modelos de contabilizacao referentes ao modulo/transacao em questao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContab_Le_Todos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_LE_TODOS", X)
'    If lErro <> SUCESSO Then Error 36144
'
'
'
'    PadraoContab_Le_Todos = X.PadraoContab_Le_Todos(sSiglaModulo, iTransacao, colPadraoContab, iSubtipo)
'
'    Exit Function
'
'Erro_PadraoContab_Le_Todos:
'
'    PadraoContab_Le_Todos = Err
'
'    Select Case Err
'
'        Case 36144
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144904)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContab_Le1(objPadraoContab As ClassPadraoContab) As Long
''retorna os dados do  modelo passado como parametro. Retorna seus dados em objPadraoContab.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContab_Le1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_LE1", X)
'    If lErro <> SUCESSO Then Error 36161
'
'
'
'    PadraoContab_Le1 = X.PadraoContab_Le1(objPadraoContab)
'
'    Exit Function
'
'Erro_PadraoContab_Le1:
'
'    PadraoContab_Le1 = Err
'
'    Select Case Err
'
'        Case 36161
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144905)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContab_Grava(objPadraoContab As ClassPadraoContab, colPadraoContabItem As ClassColPadraoContabItem) As Long
''grava o modelo padrão de contabilização em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContab_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_GRAVA", X)
'    If lErro <> SUCESSO Then Error 36197
'
'
'
'    PadraoContab_Grava = X.PadraoContab_Grava(objPadraoContab, colPadraoContabItem)
'
'    Exit Function
'
'Erro_PadraoContab_Grava:
'
'    PadraoContab_Grava = Err
'
'    Select Case Err
'
'        Case 36197
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144906)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContabItem_Exclui(objPadraoContab As ClassPadraoContab) As Long
''excluir os itens do padrao de contabilizacao
''ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContabItem_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTABITEM_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 36198
'
'
'
'    PadraoContabItem_Exclui = X.PadraoContabItem_Exclui(objPadraoContab)
'
'    Exit Function
'
'Erro_PadraoContabItem_Exclui:
'
'    PadraoContabItem_Exclui = Err
'
'    Select Case Err
'
'        Case 36198
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144907)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContab_Exclui(objPadraoContab As ClassPadraoContab) As Long
''exclui o modelo padrão de contabilização em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContab_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 36199
'
'
'
'    PadraoContab_Exclui = X.PadraoContab_Exclui(objPadraoContab)
'
'    Exit Function
'
'Erro_PadraoContab_Exclui:
'
'    PadraoContab_Exclui = Err
'
'    Select Case Err
'
'        Case 36199
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144908)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContabItem_Inserir(colPadraoContabItem As ClassColPadraoContabItem) As Long
''insere os lancamentos do padrao de contabilizacao
''ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContabItem_Inserir
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_INSERIR", X)
'    If lErro <> SUCESSO Then Error 36202
'
'
'
'    PadraoContabItem_Inserir = X.PadraoContabItem_Inserir(colPadraoContabItem)
'
'    Exit Function
'
'Erro_PadraoContabItem_Inserir:
'
'    PadraoContabItem_Inserir = Err
'
'    Select Case Err
'
'        Case 36202
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144909)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le1(iFilialEmpresa As Integer) As Long
''Lê todos os lotes de uma FilialEmpresa
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE1", X)
'    If lErro <> SUCESSO Then Error 16862
'
'
'
'    LotePendente_Le1 = X.LotePendente_Le1(iFilialEmpresa)
'
'    Exit Function
'
'Erro_LotePendente_Le1:
'
'    LotePendente_Le1 = Err
'
'    Select Case Err
'
'        Case 16862
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144910)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_Doc2(iFilialEmpresa As Integer, iTransacao As Integer, lNumIntDoc As Long, colLancamentos As Collection, sSiglaModulo As String) As Long
''le os lançamentos pendentes relativos ao documento passado como parametro e coloca-os em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_Doc2
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_DOC2", X)
'    If lErro <> SUCESSO Then Error 36297
'
'
'
'    LanPendente_Le_Doc2 = X.LanPendente_Le_Doc2(iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos, sSiglaModulo)
'
'    Exit Function
'
'Erro_LanPendente_Le_Doc2:
'
'    LanPendente_Le_Doc2 = Err
'
'    Select Case Err
'
'        Case 36297
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144911)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Le_Doc2(iFilialEmpresa As Integer, iTransacao As Integer, lNumIntDoc As Long, colLancamentos As Collection) As Long
''le os lançamentos relativos ao documento passado como parametro e coloca-os em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Doc2
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_DOC2", X)
'    If lErro <> SUCESSO Then Error 36298
'
'
'
'    Lancamento_Le_Doc2 = X.Lancamento_Le_Doc2(iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos)
'
'    Exit Function
'
'Erro_Lancamento_Le_Doc2:
'
'    Lancamento_Le_Doc2 = Err
'
'    Select Case Err
'
'        Case 36298
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144912)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TransacaoCTB_Le(objTransacaoCTB As ClassTransacaoCTB) As Long
''Le os dados da transacao passada como parametro e devolve os dados em objTransacaoCTB
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_TransacaoCTB_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_TRANSACAOTCTB_LE", X)
'    If lErro <> SUCESSO Then Error 36369
'
'
'
'    TransacaoCTB_Le = X.TransacaoCTB_Le(objTransacaoCTB)
'
'    Exit Function
'
'Erro_TransacaoCTB_Le:
'
'    TransacaoCTB_Le = Err
'
'    Select Case Err
'
'        Case 36369
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144913)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TransacaoCTB_Le1(objTransacaoCTB As ClassTransacaoCTB) As Long
''Le os dados da transacao passada como parametro e devolve os dados em objTransacaoCTB
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_TransacaoCTB_Le1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_TRANSACAOTCTB_LE1", X)
'    If lErro <> SUCESSO Then gError 83228
'
'
'
'    TransacaoCTB_Le1 = X.TransacaoCTB_Le1(objTransacaoCTB)
'
'    Exit Function
'
'Erro_TransacaoCTB_Le1:
'
'    TransacaoCTB_Le1 = gErr
'
'    Select Case gErr
'
'        Case 83228
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144914)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Grava0(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection) As Long
''Grava os dados do lançamento. Utilizado por Lancamento_Grava
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Grava0
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_GRAVA0", X)
'    If lErro <> SUCESSO Then Error 36383
'
'    Lancamento_Grava0 = X.Lancamento_Grava0(objLancamento_Cabecalho, colLancamento_Detalhe)
'
'    Exit Function
'
'Erro_Lancamento_Grava0:
'
'    Lancamento_Grava0 = Err
'
'    Select Case Err
'
'        Case 36383
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144915)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Atualiza0(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection) As Long
''Atualiza os lançamentos passados como parametro. Continuação de Lancamento_Atualiza.
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Atualiza0
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_ATUALIZA0", X)
'    If lErro <> SUCESSO Then Error 36385
'
'
'    Lancamento_Atualiza0 = X.Lancamento_Atualiza0(objLancamento_Cabecalho, colLancamento_Detalhe)
'
'    Exit Function
'
'Erro_Lancamento_Atualiza0:
'
'    Lancamento_Atualiza0 = Err
'
'    Select Case Err
'
'        Case 36385
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144916)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Exclui0(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
''Exclui os lançamentos passados como parametro. Continuação  da rotina Lancamento_Exclui
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_Lancamento_Exclui0
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_EXCLUI0", X)
'    If lErro <> SUCESSO Then Error 36392
'
'
'    Lancamento_Exclui0 = X.Lancamento_Exclui0(objLancamento_Cabecalho)
'
'    Exit Function
'
'Erro_Lancamento_Exclui0:
'
'    Lancamento_Exclui0 = Err
'
'    Select Case Err
'
'        Case 36392
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144917)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContabItem_Le_Custo(objPadraoContab As ClassPadraoContab, colPadraoContabItem As ClassColPadraoContabItem) As Long
''retorna os lançamentos de custo de objPadraoContab . Os lançamentos são retornados em colPadraoContabItem
''NÃO INCLUI OS LANÇAMENTOS NORMAIS
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContabItem_Le_Custo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTABITEM_LE_CUSTO", X)
'    If lErro <> SUCESSO Then Error 36408
'
'
'
'    PadraoContabItem_Le_Custo = X.PadraoContabItem_Le_Custo(objPadraoContab, colPadraoContabItem)
'
'    Exit Function
'
'Erro_PadraoContabItem_Le_Custo:
'
'    PadraoContabItem_Le_Custo = Err
'
'    Select Case Err
'
'        Case 36408
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144918)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le2(objLote As ClassLote) As Long
''Verifica se tem um lote pendente que não esteja em processo de atualização com a Filial/Origem/Exercicio/Periodo passados como parametro e se tiver devolve seus dados.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le2
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE2", X)
'    If lErro <> SUCESSO Then Error 36418
'
'
'
'    LotePendente_Le2 = X.LotePendente_Le2(objLote)
'
'    Exit Function
'
'Erro_LotePendente_Le2:
'
'    LotePendente_Le2 = Err
'
'    Select Case Err
'
'        Case 36418
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144919)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Insere_LoteCusto(objLote As ClassLote) As Long
''grava um lote de custo na tabela de lotes pendentes
''ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Insere_LoteCusto
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_INSERE_LOTECUSTO", X)
'    If lErro <> SUCESSO Then Error 36432
'
'
'
'    LotePendente_Insere_LoteCusto = X.LotePendente_Insere_LoteCusto(objLote)
'
'    Exit Function
'
'Erro_LotePendente_Insere_LoteCusto:
'
'    LotePendente_Insere_LoteCusto = Err
'
'    Select Case Err
'
'        Case 36432
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144920)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Voucher_Automatico1(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, sOrigem As String, lDoc As Long) As Long
''retorna o número do proximo voucher (documento) disponivel. Não verifica a tabela de Lancamentos ou Lancamentos Pendentes.
''ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Voucher_Automatico1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VOUCHER_AUTOMATICO1", X)
'    If lErro <> SUCESSO Then Error 36440
'
'
'
'    Voucher_Automatico1 = X.Voucher_Automatico1(iFilialEmpresa, iExercicio, iPeriodo, sOrigem, lDoc)
'
'    Exit Function
'
'Erro_Voucher_Automatico1:
'
'    Voucher_Automatico1 = Err
'
'    Select Case Err
'
'        Case 36440
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144921)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Gera_Lancamentos_Custo(sModulo As String, objTransacao As Object, sModelo As String, colLancamentos As ClassColLancamentos, iSubTipoTransacao As Integer) As Long
''retorna somente os lançamentos de custo do Modulo/Transacao passados como parametro. Os lançamentos são retornados em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Gera_Lancamentos_Custo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_GERALANCAMENTOS_CUSTO", X)
'    If lErro <> SUCESSO Then Error 36445
'
'
'
'    Gera_Lancamentos_Custo = X.Gera_Lancamentos_Custo(sModulo, objTransacao, sModelo, colLancamentos, iSubTipoTransacao)
'
'    Exit Function
'
'Erro_Gera_Lancamentos_Custo:
'
'    Gera_Lancamentos_Custo = Err
'
'    Select Case Err
'
'        Case 36445
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144922)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Produto_Le_Apropriacao(lComando As Long, sProduto As String, iApropriacao As Integer) As Long
''Lê o atributo Apropriação do produto, cujo codigo foi passado como parâmetro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Produto_Le_Apropriacao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PRODUTO_LE_APROPRIACAO", X)
'    If lErro <> SUCESSO Then Error 36530
'
'
'    Produto_Le_Apropriacao = X.Produto_Le_Apropriacao(lComando, sProduto, iApropriacao)
'
'    Exit Function
'
'Erro_Produto_Le_Apropriacao:
'
'    Produto_Le_Apropriacao = Err
'
'    Select Case Err
'
'        Case 36530
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144923)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst_Le_CustoMedio(ByVal lComando As Long, iFilialEmpresa As Integer, iAno As Integer, sProduto As String, iMes As Integer, dCustoMedio As Double) As Long
''le os dados de custo medio de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst_Le_CustoMedio
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SLDMESEST_LE_CUSTOMEDIO", X)
'    If lErro <> SUCESSO Then Error 36523
'
'
'    SldMesEst_Le_CustoMedio = X.SldMesEst_Le_CustoMedio(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst_Le_CustoMedio:
'
'    SldMesEst_Le_CustoMedio = Err
'
'    Select Case Err
'
'        Case 36523
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144924)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst_Le_CustoStandard(ByVal lComando As Long, iFilialEmpresa As Integer, iAno As Integer, sProduto As String, iMes As Integer, dCustoStandard As Double) As Long
''le os dados de custo standard de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst_Le_CustoStandard
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SLDMESEST_LE_CUSTOSTANDARD", X)
'    If lErro <> SUCESSO Then Error 36524
'
'
'    SldMesEst_Le_CustoStandard = X.SldMesEst_Le_CustoStandard(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoStandard)
'
'    Exit Function
'
'Erro_SldMesEst_Le_CustoStandard:
'
'    SldMesEst_Le_CustoStandard = Err
'
'    Select Case Err
'
'        Case 36524
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144925)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Netos(sConta As String, colPlanoConta As Collection) As Long
''le todos os netos de sConta e coloca-os em colPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Netos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_NETOS", X)
'    If lErro <> SUCESSO Then Error 36763
'
'
'    PlanoConta_Le_Netos = X.PlanoConta_Le_Netos(sConta, colPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Netos:
'
'    PlanoConta_Le_Netos = Err
'
'    Select Case Err
'
'        Case 36763
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144926)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Niveis0e1(colPlanoConta As Collection) As Long
''le todas as contas de nível 0 e 1 e coloca-as em colPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Niveis0e1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_NIVEIS0E1", X)
'    If lErro <> SUCESSO Then Error 36768
'
'
'    PlanoConta_Le_Niveis0e1 = X.PlanoConta_Le_Niveis0e1(colPlanoConta)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Niveis0e1:
'
'    PlanoConta_Le_Niveis0e1 = Err
'
'    Select Case Err
'
'        Case 36768
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144927)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Automatico(lCodigo As Long) As Long
''retorna o número do proximo código de Rateio Offline disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Automatico
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 36795
'
'
'    RateioOff_Automatico = X.RateioOff_Automatico(lCodigo)
'
'    Exit Function
'
'Erro_RateioOff_Automatico:
'
'    RateioOff_Automatico = Err
'
'    Select Case Err
'
'        Case 36795
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144928)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContabItem_Le_Modelo(objPadraoContab As ClassPadraoContab, colPadraoContabItem As ClassColPadraoContabItem) As Long
''retorna os lançamentos de objPadraoContab referentes ao modelo em questão. Os lançamentos são retornados em colPadraoContabItem
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContabItem_Le_Modelo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTABITEM_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 36809
'
'
'
'    PadraoContabItem_Le_Modelo = X.PadraoContabItem_Le_Modelo(objPadraoContab, colPadraoContabItem)
'
'    Exit Function
'
'Erro_PadraoContabItem_Le_Modelo:
'
'    PadraoContabItem_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 36809
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144929)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Le_Doc1(lComando As Long, objRateioOff As ClassRateioOff, colRateioOff As Collection) As Long
''le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOff
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Le_Doc1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_LE_DOC1", X)
'    If lErro <> SUCESSO Then Error 36825
'
'
'    RateioOff_Le_Doc1 = X.RateioOff_Le_Doc1(lComando, objRateioOff, colRateioOff)
'
'    Exit Function
'
'Erro_RateioOff_Le_Doc1:
'
'    RateioOff_Le_Doc1 = Err
'
'    Select Case Err
'
'        Case 36825
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144930)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Le_Lock(lComando As Long, objLote As ClassLote) As Long
''le e loca um lote a partir da chave fornecida através de objLote. Devolve os dados em objLote
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Le_Lock
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_LE_LOCK", X)
'    If lErro <> SUCESSO Then Error 36837
'
'
'    LotePendente_Le_Lock = X.LotePendente_Le_Lock(lComando, objLote)
'
'    Exit Function
'
'Erro_LotePendente_Le_Lock:
'
'    LotePendente_Le_Lock = Err
'
'    Select Case Err
'
'        Case 36837
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144931)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Grava_Estorno(objLancamento_Cabecalho As ClassLancamento_Cabecalho, objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
''objLancamento_Cabecalho contém algumas infos sobre os lançamentos de estorno a serem gravados. objLancamento_Cabecalho1 contem o lote ou doc a ser extornado
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Grava_Estorno
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_GRAVA_ESTORNO", X)
'    If lErro <> SUCESSO Then Error 36922
'
'
'    Lancamento_Grava_Estorno = X.Lancamento_Grava_Estorno(objLancamento_Cabecalho, objLancamento_Cabecalho1)
'
'    Exit Function
'
'Erro_Lancamento_Grava_Estorno:
'
'    Lancamento_Grava_Estorno = Err
'
'    Select Case Err
'
'        Case 36922
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144932)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_RateioOff(sNomeArqParam As String, ByVal iLote As Integer, ByVal dtData As Date, ByVal iPeriodoInicial As Integer, ByVal iPeriodoFinal As Integer, colRateios As Collection, ByVal iFilialEmpresa As Integer, sHistorico As String) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_RateioOff
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_RATEIOOFF", X)
'    If lErro <> SUCESSO Then Error 41535
'
'
'    Rotina_RateioOff = X.Rotina_RateioOff(sNomeArqParam, iLote, dtData, iPeriodoInicial, iPeriodoFinal, colRateios, iFilialEmpresa, sHistorico)
'
'    Exit Function
'
'Erro_Rotina_RateioOff:
'
'    Rotina_RateioOff = Err
'
'    Select Case Err
'
'        Case 41535
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144933)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Netos_Modulo(sConta As String, colPlanoConta As Collection, sSiglaModulo As String) As Long
''le todos os netos de sConta visiveis ao modulo sSiglaModulo e coloca-as em colPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Netos_Modulo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_NETOS_MODULO", X)
'    If lErro <> SUCESSO Then Error 44017
'
'
'    PlanoConta_Le_Netos_Modulo = X.PlanoConta_Le_Netos_Modulo(sConta, colPlanoConta, sSiglaModulo)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Netos_Modulo:
'
'    PlanoConta_Le_Netos_Modulo = Err
'
'    Select Case Err
'
'        Case 44017
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144934)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PlanoConta_Le_Niveis0e1_Modulo(colPlanoConta As Collection, sSiglaModulo As String) As Long
''le todas as contas de nível 0 e 1 visiveis ao Modulo (sSiglaModulo) e coloca-as em colPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PlanoConta_Le_Niveis0e1_Modulo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PLANOCONTA_LE_NIVEIS0E1_MODULO", X)
'    If lErro <> SUCESSO Then Error 44018
'
'
'    PlanoConta_Le_Niveis0e1_Modulo = X.PlanoConta_Le_Niveis0e1_Modulo(colPlanoConta, sSiglaModulo)
'
'    Exit Function
'
'Erro_PlanoConta_Le_Niveis0e1_Modulo:
'
'    PlanoConta_Le_Niveis0e1_Modulo = Err
'
'    Select Case Err
'
'        Case 44018
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144935)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carga_Arvore_Conta_Modulo1(objNodeAvo As Node, colNodes As Nodes, sSiglaModulo As String) As Long
''move os dados do plano de contas do banco de dados para a arvore colNodes.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Carga_Arvore_Conta_Modulo1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CARGA_ARVORE_CONTA_MODULO1", X)
'    If lErro <> SUCESSO Then Error 44032
'
'
'    Carga_Arvore_Conta_Modulo1 = X.Carga_Arvore_Conta_Modulo1(objNodeAvo, colNodes, sSiglaModulo)
'
'    Exit Function
'
'Erro_Carga_Arvore_Conta_Modulo1:
'
'    Carga_Arvore_Conta_Modulo1 = Err
'
'    Select Case Err
'
'        Case 44032
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144936)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaSimples_Critica(ByVal sContaSimplesText As String, ByVal sContaSimplesClipText As String, objPlanoConta As ClassPlanoConta) As Long
''verifica se a contasimples está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaSimples_Critica
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTASIMPLES_CRITICA", X)
'    If lErro <> SUCESSO Then Error 44036
'
'
'    ContaSimples_Critica = X.ContaSimples_Critica(sContaSimplesText, sContaSimplesClipText, objPlanoConta)
'
'    Exit Function
'
'Erro_ContaSimples_Critica:
'
'    ContaSimples_Critica = Err
'
'    Select Case Err
'
'        Case 44036
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144937)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Le_ContaCcl(sConta As String, sCcl As String) As Long
''verifica se existe na tabela RateioOff pelo menos um rateio com a conta e centro de custo em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Le_ContaCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_LE_CONTACCL", X)
'    If lErro <> SUCESSO Then Error 44050
'
'
'    RateioOff_Le_ContaCcl = X.RateioOff_Le_ContaCcl(sConta, sCcl)
'
'    Exit Function
'
'Erro_RateioOff_Le_ContaCcl:
'
'    RateioOff_Le_ContaCcl = Err
'
'    Select Case Err
'
'        Case 44050
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144938)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MnemonicoCTBValor_Le(objMnemonico As ClassMnemonicoCTBValor) As Long
''retorna os dados do mnemonico passado como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MnemonicoCTBValor_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MNEMONICOCTBVALOR_LE", X)
'    If lErro <> SUCESSO Then Error 44064
'
'
'    MnemonicoCTBValor_Le = X.MnemonicoCTBValor_Le(objMnemonico)
'
'    Exit Function
'
'Erro_MnemonicoCTBValor_Le:
'
'    MnemonicoCTBValor_Le = Err
'
'    Select Case Err
'
'        Case 44064
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144939)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ContaSimples_Critica_Modulo(ByVal sContaSimplesText As String, ByVal sContaSimplesClipText As String, objPlanoConta As ClassPlanoConta, sSiglaModulo As String) As Long
''verifica se a contasimples está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ContaSimples_Critica_Modulo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTASIMPLES_CRITICA_MODULO", X)
'    If lErro <> SUCESSO Then Error 44102
'
'
'    ContaSimples_Critica_Modulo = X.ContaSimples_Critica_Modulo(sContaSimplesText, sContaSimplesClipText, objPlanoConta, sSiglaModulo)
'
'    Exit Function
'
'Erro_ContaSimples_Critica_Modulo:
'
'    ContaSimples_Critica_Modulo = Err
'
'    Select Case Err
'
'        Case 44102
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144940)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
''??? excluido Jones 24/04/00
''''Function LanPrePendente_Grava_NumIntDocOrigem(iFilialEmpresa As Integer, lNumIntDoc As Long, iTransacao As Integer, Optional lNumIntDocOrigem As Long, Optional iTransacaoOrigem As Integer) As Long
'''''transforma os lançamentos pré-pendentes pertencentes ao documento lNumIntDoc da Transacao iTransacao em lançamentos pendentes ou atualizados
'''''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
''''
''''
''''Dim X As Object, lErro As Long
''''
''''On Error GoTo Erro_LanPrePendente_Grava_NumIntDocOrigem
''''
''''
''''
''''
''''    lErro = Rotina_ObterObjetoFuncao("CTB_LANPREPENDENTE_GRAVA_NUMINTDOCORIGEM", X)
''''    If lErro <> SUCESSO Then Error 44253
''''
''''
''''
''''    If Not IsMissing(lNumIntDocOrigem) Then
''''        If Not IsMissing(iTransacaoOrigem) Then
''''            LanPrePendente_Grava_NumIntDocOrigem = X.LanPrePendente_Grava_NumIntDocOrigem(iFilialEmpresa, lNumIntDoc, iTransacao, lNumIntDocOrigem, iTransacaoOrigem)
''''        Else
''''            LanPrePendente_Grava_NumIntDocOrigem = X.LanPrePendente_Grava_NumIntDocOrigem(iFilialEmpresa, lNumIntDoc, iTransacao, lNumIntDocOrigem)
''''        End If
''''    Else
''''        LanPrePendente_Grava_NumIntDocOrigem = X.LanPrePendente_Grava_NumIntDocOrigem(iFilialEmpresa, lNumIntDoc, iTransacao)
''''    End If
''''
''''    Exit Function
''''
''''Erro_LanPrePendente_Grava_NumIntDocOrigem:
''''
''''    LanPrePendente_Grava_NumIntDocOrigem = Err
''''
''''    Select Case Err
''''
''''        Case 44253
''''
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144941)
''''
''''    End Select
''''
''''    Exit Function
''''
''''End Function
'
'Function RelDRE_Le_Modelo(sRelatorio As String, sModelo As String, colRelDRE As Collection) As Long
''Le o modelo em questão e seu componentes
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Le_Modelo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 45445
'
'
'    RelDRE_Le_Modelo = X.RelDRE_Le_Modelo(sRelatorio, sModelo, colRelDRE)
'
'    Exit Function
'
'Erro_RelDRE_Le_Modelo:
'
'    RelDRE_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 45445
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144942)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDREConta_Le_Modelo(sRelatorio As String, sModelo As String, colRelDREConta As Collection) As Long
''Le os dados de Contas do modelo atual para a coleção
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDREConta_Le_Modelo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRECONTA_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 45446
'
'
'    RelDREConta_Le_Modelo = X.RelDREConta_Le_Modelo(sRelatorio, sModelo, colRelDREConta)
'
'    Exit Function
'
'Erro_RelDREConta_Le_Modelo:
'
'    RelDREConta_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 45446
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144943)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDREFormula_Le_Modelo(sRelatorio As String, sModelo As String, colRelDREFormula As Collection) As Long
''Le os dados de Fórmulas do modelo atual para a coleção
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDREFormula_Le_Modelo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDREFORMULA_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 45447
'
'
'    RelDREFormula_Le_Modelo = X.RelDREFormula_Le_Modelo(sRelatorio, sModelo, colRelDREFormula)
'
'    Exit Function
'
'Erro_RelDREFormula_Le_Modelo:
'
'    RelDREFormula_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 45447
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144944)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDRE_Grava(sRelatorio As String, sModelo As String, colRelDRE As Collection, colRelDREConta As Collection, colRelDREFormula As Collection) As Long
''Grava no BD o modelo em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_GRAVA", X)
'    If lErro <> SUCESSO Then Error 45448
'
'
'    RelDRE_Grava = X.RelDRE_Grava(sRelatorio, sModelo, colRelDRE, colRelDREConta, colRelDREFormula)
'
'    Exit Function
'
'Erro_RelDRE_Grava:
'
'    RelDRE_Grava = Err
'
'    Select Case Err
'
'        Case 45448
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144945)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDRE_Exclui(sRelatorio As String, sModelo As String) As Long
''exclui os elementos do modelo em questão da tabela RelDRE
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 45449
'
'
'    RelDRE_Exclui = X.RelDRE_Exclui(sRelatorio, sModelo)
'
'    Exit Function
'
'Erro_RelDRE_Exclui:
'
'    RelDRE_Exclui = Err
'
'    Select Case Err
'
'        Case 45449
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144946)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDREConta_Exclui(sRelatorio As String, sModelo As String) As Long
''exclui os elementos do modelo em questão da tabela RelDRE
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDREConta_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRECONTA_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 45450
'
'
'    RelDREConta_Exclui = X.RelDREConta_Exclui(sRelatorio, sModelo)
'
'    Exit Function
'
'Erro_RelDREConta_Exclui:
'
'    RelDREConta_Exclui = Err
'
'    Select Case Err
'
'        Case 45450
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144947)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDREFormula_Exclui(sRelatorio As String, sModelo As String) As Long
''exclui os elementos do modelo em questão da tabela RelDRE
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDREFormula_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDREFORMULA_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 45451
'
'
'    RelDREFormula_Exclui = X.RelDREFormula_Exclui(sRelatorio, sModelo)
'
'    Exit Function
'
'Erro_RelDREFormula_Exclui:
'
'    RelDREFormula_Exclui = Err
'
'    Select Case Err
'
'        Case 45451
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144948)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDRE_Insere(sRelatorio As String, sModelo As String, colRelDRE As Collection) As Long
''inclui os elementos de colRelDRE na tabela RelDRE
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Insere
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_INSERE", X)
'    If lErro <> SUCESSO Then Error 45452
'
'
'    RelDRE_Insere = X.RelDRE_Insere(sRelatorio, sModelo, colRelDRE)
'
'    Exit Function
'
'Erro_RelDRE_Insere:
'
'    RelDRE_Insere = Err
'
'    Select Case Err
'
'        Case 45452
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144949)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDREConta_Insere(sRelatorio As String, sModelo As String, colRelDREConta As Collection) As Long
''inclui os elementos de colRelDREConta na tabela RelDREConta
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDREConta_Insere
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRECONTA_INSERE", X)
'    If lErro <> SUCESSO Then Error 45453
'
'
'    RelDREConta_Insere = X.RelDREConta_Insere(sRelatorio, sModelo, colRelDREConta)
'
'    Exit Function
'
'Erro_RelDREConta_Insere:
'
'    RelDREConta_Insere = Err
'
'    Select Case Err
'
'        Case 45453
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144950)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDREFormula_Insere(sRelatorio As String, sModelo As String, colRelDREFormula As Collection) As Long
''inclui os elementos de colRelDREFormula na tabela RelDREFormula
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDREFormula_Insere
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDREFORMULA_INSERE", X)
'    If lErro <> SUCESSO Then Error 45454
'
'
'    RelDREFormula_Insere = X.RelDREFormula_Insere(sRelatorio, sModelo, colRelDREFormula)
'
'    Exit Function
'
'Erro_RelDREFormula_Insere:
'
'    RelDREFormula_Insere = Err
'
'    Select Case Err
'
'        Case 45454
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144951)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDRE_Exclui1(sRelatorio As String, sModelo As String) As Long
''Grava no BD o modelo em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Exclui1
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_EXCLUI1", X)
'    If lErro <> SUCESSO Then Error 45455
'
'    RelDRE_Exclui1 = X.RelDRE_Exclui1(sRelatorio, sModelo)
'
'    Exit Function
'
'Erro_RelDRE_Exclui1:
'
'    RelDRE_Exclui1 = Err
'
'    Select Case Err
'
'        Case 45455
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144952)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDRE_Le_Modelos_Distintos(sRelatorio As String, colModelos As Collection) As Long
''Le todos os modelos disponíveis no BD
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Le_Modelos_Distintos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_LE_MODELOS_DISTINTOS", X)
'    If lErro <> SUCESSO Then Error 45456
'
'
'    RelDRE_Le_Modelos_Distintos = X.RelDRE_Le_Modelos_Distintos(sRelatorio, colModelos)
'
'    Exit Function
'
'Erro_RelDRE_Le_Modelos_Distintos:
'
'    RelDRE_Le_Modelos_Distintos = Err
'
'    Select Case Err
'
'        Case 45456
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144953)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CTBConfig_Grava_Configs(colConfigs As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CTBConfig_Grava_Configs
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CTBCONFIG_GRAVA_CONFIGS", X)
'    If lErro <> SUCESSO Then Error 44663
'
'
'    CTBConfig_Grava_Configs = X.CTBConfig_Grava_Configs(colConfigs)
'
'    Exit Function
'
'Erro_CTBConfig_Grava_Configs:
'
'    CTBConfig_Grava_Configs = Err
'
'    Select Case Err
'
'        Case 44663
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144954)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Grava_Trans(objExercicio As ClassExercicio, colPeriodos As Collection) As Long
''grava os dados do exercicio objExercicio.
''IMPORTANTE: Tem que ser chamado em transação
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Grava_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_GRAVA_TRANS", X)
'    If lErro <> SUCESSO Then Error 20718
'
'
'    Exercicio_Grava_Trans = X.Exercicio_Grava_Trans(objExercicio, colPeriodos)
'
'    Exit Function
'
'Erro_Exercicio_Grava_Trans:
'
'    Exercicio_Grava_Trans = Err
'
'    Select Case Err
'
'        Case 20718
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144955)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Segmento_Grava_Conta_Trans(colSegmentos As Collection) As Long
''grava os registros da colecao na tabela Segmento
''IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Segmento_Grava_Conta_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SEGMENTO_GRAVA_CONTA_TRANS", X)
'    If lErro <> SUCESSO Then Error 20720
'
'
'    Segmento_Grava_Conta_Trans = X.Segmento_Grava_Conta_Trans(colSegmentos)
'
'    Exit Function
'
'Erro_Segmento_Grava_Conta_Trans:
'
'    Segmento_Grava_Conta_Trans = Err
'
'    Select Case Err
'
'        Case 20720
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144956)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CTBConfig_Grava_Configs_Trans(colConfigs As Collection) As Long
''Grava na tabela CRConfig uma coleção de registros de configuração
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CTBConfig_Grava_Configs_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CTBCONFIG_GRAVA_CONFIGS_TRANS", X)
'    If lErro <> SUCESSO Then Error 44671
'
'
'    CTBConfig_Grava_Configs_Trans = X.CTBConfig_Grava_Configs_Trans(colConfigs)
'
'    Exit Function
'
'Erro_CTBConfig_Grava_Configs_Trans:
'
'    CTBConfig_Grava_Configs_Trans = Err
'
'    Select Case Err
'
'        Case 44671
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144957)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ConfiguracaoSetup_Grava_Trans(objConfiguracao As ClassConfiguracao) As Long
''Insere os seguintes campos da tabela Configuracao : LotePorPeriodo, DocPorPeriodo, UsoCcl
''IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ConfiguracaoSetup_Grava_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAOSETUP_GRAVA_TRANS", X)
'    If lErro <> SUCESSO Then Error 20722
'
'
'    ConfiguracaoSetup_Grava_Trans = X.ConfiguracaoSetup_Grava_Trans(objConfiguracao)
'
'    Exit Function
'
'Erro_ConfiguracaoSetup_Grava_Trans:
'
'    ConfiguracaoSetup_Grava_Trans = Err
'
'    Select Case Err
'
'        Case 20722
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144958)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exercicio_Instalacao_Filial(iFilialEmpresa As Integer) As Long
''Insere os registros em MvPerCta, MvPerCcl, ExercicioOrigem, ExerciciosFilial e PeriodosFilial para a filial em questão em todos os exercícios existentes.
''AVISO: ESSA ROTINA SO PODE SER CHAMADA COM A TRANSACAO ABERTA
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Exercicio_Instalacao_Filial
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_EXERCICIO_INSTALACAO_FILIAL", X)
'    If lErro <> SUCESSO Then Error 44714
'
'
'    Exercicio_Instalacao_Filial = X.Exercicio_Instalacao_Filial(iFilialEmpresa)
'
'    Exit Function
'
'Erro_Exercicio_Instalacao_Filial:
'
'    Exercicio_Instalacao_Filial = Err
'
'    Select Case Err
'
'        Case 44714
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144959)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Segmento_Grava_Ccl(colSegmentos As Collection) As Long
''grava os registros da colecao na tabela Segmento
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Segmento_Grava_Ccl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SEGMENTO_GRAVA_CCL", X)
'    If lErro <> SUCESSO Then Error 44734
'
'
'    Segmento_Grava_Ccl = X.Segmento_Grava_Ccl(colSegmentos)
'
'    Exit Function
'
'Erro_Segmento_Grava_Ccl:
'
'    Segmento_Grava_Ccl = Err
'
'    Select Case Err
'
'        Case 44734
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144960)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Segmento_Grava_Ccl_Trans(colSegmentos As Collection) As Long
''grava os registros da colecao na tabela Segmento de Ccl
''IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Segmento_Grava_Ccl_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SEGMENTO_GRAVA_CCL_TRANS", X)
'    If lErro <> SUCESSO Then Error 44735
'
'
'    Segmento_Grava_Ccl_Trans = X.Segmento_Grava_Ccl_Trans(colSegmentos)
'
'    Exit Function
'
'Erro_Segmento_Grava_Ccl_Trans:
'
'    Segmento_Grava_Ccl_Trans = Err
'
'    Select Case Err
'
'        Case 44735
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144961)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_ExisteCcl(iExiste As Integer) As Long
''le o ccl a partir da chave fornecida em objCcl. Devolve os dados em objCcl
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_ExisteCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_EXISTECCL", X)
'    If lErro <> SUCESSO Then Error 44740
'
'
'    Ccl_ExisteCcl = X.Ccl_ExisteCcl(iExiste)
'
'    Exit Function
'
'Erro_Ccl_ExisteCcl:
'
'    Ccl_ExisteCcl = Err
'
'    Select Case Err
'
'        Case 44740
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144962)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''??? excluido Jones 24/04/00
''''Function LanPrePendente_Le_Doc1(objDoc As ClassLancamento_Cabecalho) As Long
'''''Tenta ler o documento pre-pendente em questão.
''''
''''Dim X As Object, lErro As Long
''''
''''On Error GoTo Erro_LanPrePendente_Le_Doc1
''''
''''
''''
''''
''''    lErro = Rotina_ObterObjetoFuncao("CTB_LANPREPENDENTE_LE_DOC1", X)
''''    If lErro <> SUCESSO Then Error 44914
''''
''''
''''    LanPrePendente_Le_Doc1 = X.LanPrePendente_Le_Doc1(objDoc)
''''
''''    Exit Function
''''
''''Erro_LanPrePendente_Le_Doc1:
''''
''''    LanPrePendente_Le_Doc1 = Err
''''
''''    Select Case Err
''''
''''        Case 44914
''''            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144963)
''''
''''    End Select
''''
''''    Exit Function
''''
''''End Function
'
'Function Voucher_Automatico_Obter_Batch(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, sOrigem As String, lDoc As Long, lComando As Long) As Long
''retorna em lDoc o proximo numero p/criacao de lancamentos pendentes
''Não verifica a tabela de Lancamentos ou Lancamentos Pendentes pois a origem tem que ser de lancamentos automaticos.
''ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Voucher_Automatico_Obter_Batch
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VOUCHER_AUTOMATICO_OBTER_BATCH", X)
'    If lErro <> SUCESSO Then Error 45911
'
'
'    Voucher_Automatico_Obter_Batch = X.Voucher_Automatico_Obter_Batch(iFilialEmpresa, iExercicio, iPeriodo, sOrigem, lDoc, lComando)
'
'    Exit Function
'
'Erro_Voucher_Automatico_Obter_Batch:
'
'    Voucher_Automatico_Obter_Batch = Err
'
'    Select Case Err
'
'        Case 45911
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144964)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Voucher_Automatico_Gravar_Batch(lDoc As Long, lComando As Long) As Long
''grava o proximo numero p/criacao de lancamentos pendentes com o valor de lDoc
''Não verifica a tabela de Lancamentos ou Lancamentos Pendentes pois a origem tem que ser de lancamentos automaticos.
''ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Voucher_Automatico_Gravar_Batch
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VOUCHER_AUTOMATICO_GRAVAR_BATCH", X)
'    If lErro <> SUCESSO Then Error 45912
'
'
'    Voucher_Automatico_Gravar_Batch = X.Voucher_Automatico_Gravar_Batch(lDoc, lComando)
'
'    Exit Function
'
'Erro_Voucher_Automatico_Gravar_Batch:
'
'    Voucher_Automatico_Gravar_Batch = Err
'
'    Select Case Err
'
'        Case 45912
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144965)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function PadraoContab_ObtemModeloPadrao(sModulo As String, sTransacao As String, sModelo As String, iSubtipo As Integer) As Long
''retorna o modelo padrao p/o modulo-transacao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_PadraoContab_ObtemModeloPadrao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PADRAOCONTAB_OBTEMMODELOPADRAO", X)
'    If lErro <> SUCESSO Then Error 45913
'
'
'    PadraoContab_ObtemModeloPadrao = X.PadraoContab_ObtemModeloPadrao(sModulo, sTransacao, sModelo, iSubtipo)
'
'    Exit Function
'
'Erro_PadraoContab_ObtemModeloPadrao:
'
'    PadraoContab_ObtemModeloPadrao = Err
'
'    Select Case Err
'
'        Case 45913
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144966)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Grava_Trans(objLote As ClassLote) As Long
'    Dim X As Object, lErro As Long
'
'    On Error GoTo Erro_LotePendente_Grava_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_GRAVA", X)
'    If lErro <> SUCESSO Then Error 32167
'
'
'    LotePendente_Grava_Trans = X.LotePendente_Grava_Trans(objLote)
'
'    Exit Function
'
'Erro_LotePendente_Grava_Trans:
'
'    LotePendente_Grava_Trans = Err
'
'    Select Case Err
'
'        Case 32167
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144967)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Grava_Totais_Trans(objLote As ClassLote) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Grava_Totais_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_GRAVA_TOTAIS", X)
'    If lErro <> SUCESSO Then Error 32168
'
'
'    LotePendente_Grava_Totais_Trans = X.LotePendente_Grava_Totais_Trans(objLote)
'
'    Exit Function
'
'Erro_LotePendente_Grava_Totais_Trans:
'
'    LotePendente_Grava_Totais_Trans = Err
'
'    Select Case Err
'
'        Case 32168
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144968)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LotePendente_Grava_Totais_Auto(objLote As ClassLote) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LotePendente_Grava_Totais_Auto
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTEPENDENTE_GRAVA_TOTAIS", X)
'    If lErro <> SUCESSO Then Error 32177
'
'
'    LotePendente_Grava_Totais_Auto = X.LotePendente_Grava_Totais_Auto(objLote)
'
'    Exit Function
'
'Erro_LotePendente_Grava_Totais_Auto:
'
'    LotePendente_Grava_Totais_Auto = Err
'
'    Select Case Err
'
'        Case 32177
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144969)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ConfiguracaoSetup_Altera_Trans(objConfiguracao As ClassConfiguracao) As Long
''grava os registros do objeto na tabela Configuracao
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ConfiguracaoSetup_Altera_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAOSETUP_ALTERA_TRANS", X)
'    If lErro <> SUCESSO Then Error 44983
'
'
'    ConfiguracaoSetup_Altera_Trans = X.ConfiguracaoSetup_Altera_Trans(objConfiguracao)
'
'    Exit Function
'
'Erro_ConfiguracaoSetup_Altera_Trans:
'
'    ConfiguracaoSetup_Altera_Trans = Err
'
'    Select Case Err
'
'        Case 44983
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144970)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lote_LotePendente_Le(objLote As ClassLote) As Long
''verifica se para o lote passado como parametro já tem um lote ou lote pendente cadastrado com o mesmo numero
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lote_LotePendente_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_LOTEPENDENTE_LE", X)
'    If lErro <> SUCESSO Then Error 55152
'
'
'    Lote_LotePendente_Le = X.Lote_LotePendente_Le(objLote)
'
'    Exit Function
'
'Erro_Lote_LotePendente_Le:
'
'    Lote_LotePendente_Le = Err
'
'    Select Case Err
'
'        Case 55152
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144971)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Sub Converte_TipoContaCcl(ByVal iTipo As Integer, sTipo As String)
''converte o tipo de conta/centro de custo do valor inteiro para o seu correspondente textual (sintetico/analitico)
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Converte_TipoContaCcl
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONVERTE_TIPOCONTACCL", X)
'    If lErro <> SUCESSO Then Error 55166
'
'
'    Call X.Converte_TipoContaCcl(iTipo, sTipo)
'
'    Exit Sub
'
'Erro_Converte_TipoContaCcl:
'
'    Select Case Err
'
'        Case 55166
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144972)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Sub Converte_NaturezaConta(ByVal iNatureza As Integer, sNatureza As String)
''converte a natureza de conta do valor inteiro para o seu correspondente textual (credito/debito)
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Converte_NaturezaConta
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONVERTE_NATUREZACONTA", X)
'    If lErro <> SUCESSO Then Error 55167
'
'
'    Call X.Converte_NaturezaConta(iNatureza, sNatureza)
'
'    Exit Sub
'
'Erro_Converte_NaturezaConta:
'
'    Select Case Err
'
'        Case 55167
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144973)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function MvPerCta_Calcula_Valor(iFilialEmpresa As Integer, iExercicio As Integer, sContaInicial As String, sContaFinal As String, dValorExercAtual As Double, dValorExercAnt As Double) As Long
''Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCta_Calcula_Valor
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCTA_CALCULA_VALOR", X)
'    If lErro <> SUCESSO Then Error 54685
'
'    MvPerCta_Calcula_Valor = X.MvPerCta_Calcula_Valor(iFilialEmpresa, iExercicio, sContaInicial, sContaFinal, dValorExercAtual, dValorExercAnt)
'
'    Exit Function
'
'Erro_MvPerCta_Calcula_Valor:
'
'    MvPerCta_Calcula_Valor = Err
'
'    Select Case Err
'
'        Case 54685
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144974)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCta_Calcula_Valor_Periodo(iFilialEmpresa As Integer, iExercicio As Integer, iPeriodo, sContaInicial As String, sContaFinal As String, dValorPerAnt As Double, dValorPerAtual As Double, dValorAcumPerAnt As Double, dValorAcumPerAtual As Double) As Long
''Calcula Valor pela soma de SldIni com os valores de crédito e débito do Exercicio/Periodo em questao
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCta_Calcula_Valor_Periodo
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCTA_CALCULA_VALOR_PERIODO", X)
'    If lErro <> SUCESSO Then Error 55901
'
'
'    MvPerCta_Calcula_Valor_Periodo = X.MvPerCta_Calcula_Valor_Periodo(iFilialEmpresa, iExercicio, iPeriodo, sContaInicial, sContaFinal, dValorPerAnt, dValorPerAtual, dValorAcumPerAnt, dValorAcumPerAtual)
'
'    Exit Function
'
'Erro_MvPerCta_Calcula_Valor_Periodo:
'
'    MvPerCta_Calcula_Valor_Periodo = Err
'
'    Select Case Err
'
'        Case 55901
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144975)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDRE_Grava_Valor(sRelatorio As String, colRelDRE As Collection) As Long
''Grava o Valor acumulado do Modelo na tabela RelDRE
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDRE_Grava_Valor
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDRE_GRAVA_VALOR", X)
'    If lErro <> SUCESSO Then Error 54686
'
'
'    RelDRE_Grava_Valor = X.RelDRE_Grava_Valor(sRelatorio, colRelDRE)
'
'    Exit Function
'
'Erro_RelDRE_Grava_Valor:
'
'    RelDRE_Grava_Valor = Err
'
'    Select Case Err
'
'        Case 54686
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144976)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamentos_Le_Doc(objDoc As ClassLancamento_Cabecalho, colLancamentos As Collection) As Long
''Lê os lançamentos relativos ao documento passado como parâmetro e coloca-os em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamentos_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTOS_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 54687
'
'
'    Lancamentos_Le_Doc = X.Lancamentos_Le_Doc(objDoc, colLancamentos)
'
'    Exit Function
'
'Erro_Lancamentos_Le_Doc:
'
'    Lancamentos_Le_Doc = Err
'
'    Select Case Err
'
'        Case 54687
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144977)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MnemonicoCTBValor_Exclui() As Long
''Exclui toda a tabela MnemonicoCTBValor
''IMPORTANTE: TEM QUE ESTAR EM TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MnemonicoCTBValor_Exclui
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MNEMONICOCTBVALOR_EXCLUI", X)
'    If lErro <> SUCESSO Then Error 54689
'
'
'    MnemonicoCTBValor_Exclui = X.MnemonicoCTBValor_Exclui()
'
'    Exit Function
'
'Erro_MnemonicoCTBValor_Exclui:
'
'    MnemonicoCTBValor_Exclui = Err
'
'    Select Case Err
'
'        Case 54689
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144978)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MnemonicoCTBValor_Grava(colMnemonico As Collection) As Long
''Grava os dados  passada como parâmetro, no BD
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MnemonicoCTBValor_Grava
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MNEMONICOCTBVALOR_GRAVA", X)
'    If lErro <> SUCESSO Then Error 54690
'
'
'    MnemonicoCTBValor_Grava = X.MnemonicoCTBValor_Grava(colMnemonico)
'
'    Exit Function
'
'Erro_MnemonicoCTBValor_Grava:
'
'    MnemonicoCTBValor_Grava = Err
'
'    Select Case Err
'
'        Case 54690
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144979)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MnemonicoCTBValor_Le_Globais(colMnemonicoGlobal As Collection) As Long
''le todos os Mnemônicos contidos na tabela MnemonicoCTBValor e coloca-os em colMnemonicoGlobal
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MnemonicoCTBValor_Le_Globais
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MNEMONICOCTBVALOR_LE_GLOBAIS", X)
'    If lErro <> SUCESSO Then Error 54691
'
'
'    MnemonicoCTBValor_Le_Globais = X.MnemonicoCTBValor_Le_Globais(colMnemonicoGlobal)
'
'    Exit Function
'
'Erro_MnemonicoCTBValor_Le_Globais:
'
'    MnemonicoCTBValor_Le_Globais = Err
'
'    Select Case Err
'
'        Case 54691
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144980)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Configuracao_Altera_DataInstalacao() As Long
''grava os registros do objeto na tabela Configuracao
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Configuracao_Altera_DataInstalacao
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONFIGURACAO_ALTERA_DATAINSTALACAO", X)
'    If lErro <> SUCESSO Then Error 55243
'
'
'    Configuracao_Altera_DataInstalacao = X.Configuracao_Altera_DataInstalacao()
'
'    Exit Function
'
'Erro_Configuracao_Altera_DataInstalacao:
'
'    Configuracao_Altera_DataInstalacao = Err
'
'    Select Case Err
'
'        Case 55243
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144981)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Conta_Analitica_Lock(sConta As String, sSiglaModulo As String) As Long
''O argumento sConta deve ser o número da conta FORMATADO BD
''le a conta passada como parametro e faz um lock shared
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Conta_Analitica_Lock
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CONTA_ANALITICA_LOCK", X)
'    If lErro <> SUCESSO Then Error 54722
'
'
'    Conta_Analitica_Lock = X.Conta_Analitica_Lock(sConta, sSiglaModulo)
'
'    Exit Function
'
'Erro_Conta_Analitica_Lock:
'
'    Conta_Analitica_Lock = Err
'
'    Select Case Err
'
'        Case 54722
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144982)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function UM_Conversao_Trans(iClasseUM As Integer, sUnidade1 As String, sUnidade2 As String, dFator As Double) As Long
''Calcula o fator de conversão entre 2 Unidades de Medida
''De sUnidade1 para sUnidade2
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_UM_Conversao_Trans
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_UM_CONVERSAO_TRANS", X)
'    If lErro <> SUCESSO Then Error 29927
'
'
'    UM_Conversao_Trans = X.UM_Conversao_Trans(iClasseUM, sUnidade1, sUnidade2, dFator)
'
'    Exit Function
'
'Erro_UM_Conversao_Trans:
'
'    UM_Conversao_Trans = Err
'
'    Select Case Err
'
'        Case 29927
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144983)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function UM_Le(objUnidadeMedida As ClassUnidadeDeMedida) As Long
''Lê a unidade de medida a partir da Classe e da SiglaUM -->
''e devolve dados em objUnidadeMedida
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_UM_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_UM_LE", X)
'    If lErro <> SUCESSO Then Error 25066
'
'
'    UM_Le = X.UM_Le(objUnidadeMedida)
'
'    Exit Function
'
'Erro_UM_Le:
'
'    UM_Le = Err
'
'    Select Case Err
'
'        Case 25066
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144984)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOff_Le_TodosDistintos(colRateioOff As Collection) As Long
''Le todos os rateios off distintos do BD
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOff_Le_TodosDistintos
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFF_LE_TODOSDISTINTOS", X)
'    If lErro <> SUCESSO Then Error 54730
'
'
'    RateioOff_Le_TodosDistintos = X.RateioOff_Le_TodosDistintos(colRateioOff)
'
'    Exit Function
'
'Erro_RateioOff_Le_TodosDistintos:
'
'    RateioOff_Le_TodosDistintos = Err
'
'    Select Case Err
'
'        Case 54730
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144985)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Exclui2(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
''Exclui os lançamentos passados como parametro. Continuação  da rotina Lancamento_Exclui
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Exclui2
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_EXCLUI2", X)
'    If lErro <> SUCESSO Then Error 54823
'
'
'    Lancamento_Exclui2 = X.Lancamento_Exclui2(objLancamento_Cabecalho)
'
'    Exit Function
'
'Erro_Lancamento_Exclui2:
'
'    Lancamento_Exclui2 = Err
'
'    Select Case Err
'
'        Case 54823
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144986)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Periodo_Le1(ByVal dtData As Date, objPeriodo As ClassPeriodo, iFilialEmpresa As Integer) As Long
''le o registro referente ao periodo em questão e coloca os dados em objPeriodo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Periodo_Le1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_PERIODO_LE1", X)
'    If lErro <> SUCESSO Then Error 54824
'
'
'    Periodo_Le1 = X.Periodo_Le1(dtData, objPeriodo, iFilialEmpresa)
'
'    Exit Function
'
'Erro_Periodo_Le1:
'
'    Periodo_Le1 = Err
'
'    Select Case Err
'
'        Case 54824
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144987)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Le_GrupoTransExc(iGrupoTransExc As Integer, lNumIntDocOrigem As Long, colLancamentos As Collection) As Long
''IMPORTANTE: está funcao é especifica da operacao de exclusao/estorno de contabilizacao
''Consulta a tabela de Lançamentos Pendentes com a chave NumIntDoc p/transacao do grupo identificado na tabela TransacaoCTB p/iGrupoTransExc.
''retorna lancamentos ordenados por filialempresa
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Le_GrupoTransExc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_LE_GRUPOTRANSEXC", X)
'    If lErro <> SUCESSO Then Error 54825
'
'
'    LanPendente_Le_GrupoTransExc = X.LanPendente_Le_GrupoTransExc(iGrupoTransExc, lNumIntDocOrigem, colLancamentos)
'
'    Exit Function
'
'Erro_LanPendente_Le_GrupoTransExc:
'
'    LanPendente_Le_GrupoTransExc = Err
'
'    Select Case Err
'
'        Case 54825
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144988)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Le_GrupoTransExc(iGrupoTransExc As Integer, lNumIntDocOrigem As Long, colLancamentos As Collection) As Long
''IMPORTANTE: está funcao é especifica da operacao de exclusao/estorno de contabilizacao
''Consulta a tabela de Lançamentos com a chave NumIntDoc p/transacao que na tabela TransacaoCTB tenha GrupoTransExc igual a iGrupoTransExc.
''retorna lancamentos ordenados por filialempresa
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_GrupoTransExc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_GRUPOTRANSEXC", X)
'    If lErro <> SUCESSO Then Error 54826
'
'
'    Lancamento_Le_GrupoTransExc = X.Lancamento_Le_GrupoTransExc(iGrupoTransExc, lNumIntDocOrigem, colLancamentos)
'
'    Exit Function
'
'Erro_Lancamento_Le_GrupoTransExc:
'
'    Lancamento_Le_GrupoTransExc = Err
'
'    Select Case Err
'
'        Case 54826
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144989)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOffContas_Le_Doc(objRateioOff As ClassRateioOff, colContas As Collection) As Long
''le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOff
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOffContas_Le_Doc
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFFCONTAS_LE_DOC", X)
'    If lErro <> SUCESSO Then Error 55815
'
'
'    RateioOffContas_Le_Doc = X.RateioOffContas_Le_Doc(objRateioOff, colContas)
'
'    Exit Function
'
'Erro_RateioOffContas_Le_Doc:
'
'    RateioOffContas_Le_Doc = Err
'
'    Select Case Err
'
'        Case 55815
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144990)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RateioOffContas_Le_Doc1(lComando As Long, objRateioOff As ClassRateioOff, colContas As Collection) As Long
''le os lançamentos relativos ao Rateio passado como parametro e coloca-os em colRateioOff
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RateioOffContas_Le_Doc1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RATEIOOFFCONTAS_LE_DOC1", X)
'    If lErro <> SUCESSO Then Error 55820
'
'
'    RateioOffContas_Le_Doc1 = X.RateioOffContas_Le_Doc1(lComando, objRateioOff, colContas)
'
'    Exit Function
'
'Erro_RateioOffContas_Le_Doc1:
'
'    RateioOffContas_Le_Doc1 = Err
'
'    Select Case Err
'
'        Case 55820
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144991)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPL_Exclui1(sRelatorio As String, sModelo As String) As Long
''Exclui do BD o modelo em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPL_Exclui1
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPL_EXCLUI1", X)
'    If lErro <> SUCESSO Then Error 60561
'
'    RelDMPL_Exclui1 = X.RelDMPL_Exclui1(sRelatorio, sModelo)
'
'    Exit Function
'
'Erro_RelDMPL_Exclui1:
'
'    RelDMPL_Exclui1 = Err
'
'    Select Case Err
'
'        Case 60561
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144992)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPL_Le_Modelo(sRelatorio As String, sModelo As String, colRel As Collection) As Long
''Le o modelo em questão e seu componentes
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPL_Le_Modelo
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPL_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 60590
'
'    RelDMPL_Le_Modelo = X.RelDMPL_Le_Modelo(sRelatorio, sModelo, colRel)
'
'    Exit Function
'
'Erro_RelDMPL_Le_Modelo:
'
'    RelDMPL_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 60590
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144993)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPLConta_Le_Modelo(sRelatorio As String, sModelo As String, colRelConta As Collection) As Long
''Le os dados de Contas do modelo atual para a coleção
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPLConta_Le_Modelo
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPLCONTA_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 60591
'
'    RelDMPLConta_Le_Modelo = X.RelDMPLConta_Le_Modelo(sRelatorio, sModelo, colRelConta)
'
'    Exit Function
'
'Erro_RelDMPLConta_Le_Modelo:
'
'    RelDMPLConta_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 60591
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144994)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPLFormula_Le_Modelo(sRelatorio As String, sModelo As String, colRelFormula As Collection) As Long
''Le os dados de Fórmulas do modelo atual para a coleção
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPLFormula_Le_Modelo
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPLFORMULA_LE_MODELO", X)
'    If lErro <> SUCESSO Then Error 60592
'
'    RelDMPLFormula_Le_Modelo = X.RelDMPLFormula_Le_Modelo(sRelatorio, sModelo, colRelFormula)
'
'    Exit Function
'
'Erro_RelDMPLFormula_Le_Modelo:
'
'    RelDMPLFormula_Le_Modelo = Err
'
'    Select Case Err
'
'        Case 60592
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144995)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPL_Le_Modelos_Distintos(sRelatorio As String, colModelos As Collection) As Long
''Le todos os modelos disponíveis no BD
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPL_Le_Modelos_Distintos
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPL_LE_MODELOS_DISTINTOS", X)
'    If lErro <> SUCESSO Then Error 60610
'
'    RelDMPL_Le_Modelos_Distintos = X.RelDMPL_Le_Modelos_Distintos(sRelatorio, colModelos)
'
'    Exit Function
'
'Erro_RelDMPL_Le_Modelos_Distintos:
'
'    RelDMPL_Le_Modelos_Distintos = Err
'
'    Select Case Err
'
'        Case 60610
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144996)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPL_Grava(sRelatorio As String, sModelo As String, colRel As Collection, colRelConta As Collection, colRelFormula As Collection) As Long
''Grava no BD o modelo em questão
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPL_Grava
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPL_GRAVA", X)
'    If lErro <> SUCESSO Then Error 60646
'
'    RelDMPL_Grava = X.RelDMPL_Grava(sRelatorio, sModelo, colRel, colRelConta, colRelFormula)
'
'    Exit Function
'
'Erro_RelDMPL_Grava:
'
'    RelDMPL_Grava = Err
'
'    Select Case Err
'
'        Case 60646
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144997)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCta_Calcula_Valor1(iFilialEmpresa As Integer, iExercicio As Integer, sContaInicial As String, sContaFinal As String, dValorExercAtual As Double) As Long
''Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCta_Calcula_Valor1
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCTA_CALCULA_VALOR1", X)
'    If lErro <> SUCESSO Then Error 60680
'
'    MvPerCta_Calcula_Valor1 = X.MvPerCta_Calcula_Valor1(iFilialEmpresa, iExercicio, sContaInicial, sContaFinal, dValorExercAtual)
'
'    Exit Function
'
'Erro_MvPerCta_Calcula_Valor1:
'
'    MvPerCta_Calcula_Valor1 = Err
'
'    Select Case Err
'
'        Case 60680
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144998)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelDMPL_Grava_Valor(sRelatorio As String, colRel As Collection) As Long
''Grava o Valor acumulado do Modelo na tabela RelDMPL
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_RelDMPL_Grava_Valor
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_RELDMPL_GRAVA_VALOR", X)
'    If lErro <> SUCESSO Then Error 60690
'
'    RelDMPL_Grava_Valor = X.RelDMPL_Grava_Valor(sRelatorio, colRel)
'
'    Exit Function
'
'Erro_RelDMPL_Grava_Valor:
'
'    RelDMPL_Grava_Valor = Err
'
'    Select Case Err
'
'        Case 60690
'            Call Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_DISPONIVEL", Err, Error$)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144999)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Grava_Estorno_Doc_Basico(objLancamento_Cabecalho As ClassLancamento_Cabecalho, objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
''Faz Atualiza os lancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Grava_Estorno_Doc_Basico
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_GRAVA_ESTORNO_DOC_BASICO", X)
'    If lErro <> SUCESSO Then Error 65086
'
'    Lancamento_Grava_Estorno_Doc_Basico = X.Lancamento_Grava_Estorno_Doc_Basico(objLancamento_Cabecalho, objLancamento_Cabecalho1)
'
'    Exit Function
'
'Erro_Lancamento_Grava_Estorno_Doc_Basico:
'
'    Lancamento_Grava_Estorno_Doc_Basico = Err
'
'    Select Case Err
'
'        Case 65086
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145000)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lote_Gera_Automatico(iFilialEmpresa As Integer, sCodUsuario As String, sOrigem As String, iExercicio As Integer, iPeriodo As Integer, iLote As Integer) As Long
''retorna o lote automatico para a contabilidade caso o usuário tenha optado por geração automatica do numero do lote
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lote_Gera_Automatico
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LOTE_GERA_AUTOMATICO", X)
'    If lErro <> SUCESSO Then Error 60891
'
'    Lote_Gera_Automatico = X.Lote_Gera_Automatico(iFilialEmpresa, sCodUsuario, sOrigem, iExercicio, iPeriodo, iLote)
'
'    Exit Function
'
'Erro_Lote_Gera_Automatico:
'
'    Lote_Gera_Automatico = Err
'
'    Select Case Err
'
'        Case 60891
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145001)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function UsuarioLote_Grava_Lote(sCodUsuario As String, sOrigem As String, ByVal iLote As Integer) As Long
''se o lote padrão para o Usuario e Origem em questão for diferente do que está gravado, atualiza. Se não existir, insere.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_UsuarioLote_Grava_Lote
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_USUARIOLOTE_GRAVA_LOTE", X)
'    If lErro <> SUCESSO Then Error 60903
'
'    UsuarioLote_Grava_Lote = X.UsuarioLote_Grava_Lote(sCodUsuario, sOrigem, iLote)
'
'    Exit Function
'
'Erro_UsuarioLote_Grava_Lote:
'
'    UsuarioLote_Grava_Lote = Err
'
'    Select Case Err
'
'        Case 60903
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145002)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Voucher_Automatico_Trans(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, sOrigem As String, lDoc As Long) As Long
''retorna o número do proximo voucher (documento) disponivel
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Voucher_Automatico_Trans
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_VOUCHER_AUTOMATICO_TRANS", X)
'    If lErro <> SUCESSO Then gError 64343
'
'    Voucher_Automatico_Trans = X.Voucher_Automatico_Trans(iFilialEmpresa, iExercicio, iPeriodo, sOrigem, lDoc)
'
'    Exit Function
'
'Erro_Voucher_Automatico_Trans:
'
'    Voucher_Automatico_Trans = gErr
'
'    Select Case gErr
'
'        Case 64343
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145003)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CustoMedioProducaoApurado_Le_Mes(sProduto As String, iMes As Integer, iAno As Integer, dCMPAtual As Double) As Long
''Calcula Custo Médio de Produção Apurado do Produto passado
''Se o mês  passado ainda não foi apurado retorna ZERO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CustoMedioProducaoApurado_Le_Mes
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CUSTOMEDIOPRODUCAOAPURADO_LE_MES", X)
'    If lErro <> SUCESSO Then Error 65294
'
'    CustoMedioProducaoApurado_Le_Mes = X.CustoMedioProducaoApurado_Le_Mes(sProduto, iMes, iAno, dCMPAtual)
'
'    Exit Function
'
'Erro_CustoMedioProducaoApurado_Le_Mes:
'
'    CustoMedioProducaoApurado_Le_Mes = Err
'
'    Select Case Err
'
'        Case 65294
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145004)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function EstoqueMes_Le_Apurado(objEstoqueMes As ClassEstoqueMes) As Long
''Le o último Ano/Mes apurado
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_EstoqueMes_Le_Apurado
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ESTOQUEMES_LE_APURADO", X)
'    If lErro <> SUCESSO Then Error 45117
'
'    EstoqueMes_Le_Apurado = X.EstoqueMes_Le_Apurado(objEstoqueMes)
'
'    Exit Function
'
'Erro_EstoqueMes_Le_Apurado:
'
'    EstoqueMes_Le_Apurado = Err
'
'    Select Case Err
'
'        Case 45117
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145005)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Calcula_CustoMedioProducao(iFilialEmpresa As Integer, sProduto As String, ByVal iAno As Integer, ByVal iMes As Integer, dCMPAtual As Double) As Long
''Calcula o Custo Médio de Produção do Produto sProduto
''iAno = indica o ano de onde deve ser lido os dados de SldMesEst
''iMes = indica o ultimo mes apurado ou zero. Zero indica que não houve apuração até o momento
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Calcula_CustoMedioProducao
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CALCULA_CUSTOMEDIOPRODUCAO", X)
'    If lErro <> SUCESSO Then Error 65293
'
'    Calcula_CustoMedioProducao = X.Calcula_CustoMedioProducao(iFilialEmpresa, sProduto, iAno, iMes, dCMPAtual)
'
'    Exit Function
'
'Erro_Calcula_CustoMedioProducao:
'
'    Calcula_CustoMedioProducao = Err
'
'    Select Case Err
'
'        Case 65293
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145006)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function SldMesEst_Le(objSldMesEst As ClassSldMesEst) As Long
''Le na tabela SldMesEst registro relativo a Ano, Produto (FilialEmpresa pega em giFilialEmpresa)
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst_Le
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SLDMESEST_LE", X)
'    If lErro <> SUCESSO Then Error 45403
'
'    SldMesEst_Le = X.SldMesEst_Le(objSldMesEst)
'
'    Exit Function
'
'Erro_SldMesEst_Le:
'
'    SldMesEst_Le = Err
'
'    Select Case Err
'
'        Case 45403
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145007)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst_Le_CustoRealProducao(ByVal lComando As Long, iFilialEmpresa As Integer, iAno As Integer, sProduto As String, iMes As Integer, dCustoRealProducao As Double) As Long
''le os dados de custo real de Produção de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst_Le_CustoRealProducao
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SLDMESEST_LE_CUSTOREALPRODUCAO", X)
'    If lErro <> SUCESSO Then Error 65468
'
'    SldMesEst_Le_CustoRealProducao = X.SldMesEst_Le_CustoRealProducao(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoRealProducao)
'
'    Exit Function
'
'Erro_SldMesEst_Le_CustoRealProducao:
'
'    SldMesEst_Le_CustoRealProducao = Err
'
'    Select Case Err
'
'        Case 65468
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145008)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function FilialFornecedor_Lock(lCodFornecedor As Long, iCodFilial As Integer) As Long
''Faz lock da filial do Fornecedor passado por parâmetro
''Chamada DENTRO DE TRANSAÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FilialFornecedor_Lock
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CRFAT_FILIALFORNECEDOR_LOCK", X)
'    If lErro <> SUCESSO Then Error 45597
'
'
'    FilialFornecedor_Lock = X.FilialFornecedor_Lock(lCodFornecedor, iCodFilial)
'
'    Exit Function
'
'Erro_FilialFornecedor_Lock:
'
'    FilialFornecedor_Lock = Err
'
'    Select Case Err
'
'        Case 45597
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145009)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function CondicaoPagto_Le(objCondicaoPagto As ClassCondicaoPagto) As Long
''le a Condicao de Pagto a partir da chave fornecida em objCondicaoPagto. Devolve os dados em objCondicaoPagto
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CondicaoPagto_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CPR_CONDICAOPAGTO_LE", X)
'    If lErro <> SUCESSO Then Error 14466
'
'
'    CondicaoPagto_Le = X.CondicaoPagto_Le(objCondicaoPagto)
'
'    Exit Function
'
'Erro_CondicaoPagto_Le:
'
'    CondicaoPagto_Le = Err
'
'    Select Case Err
'
'        Case 14466
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145010)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Fornecedor_Le_NomeReduzido(objFornecedor As ClassFornecedor) As Long
''le o Fornecedor a partir da chave(sNomeReduzido) fornecida em objFornecedor. Devolve os dados em objFornecedor
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Fornecedor_Le_NomeReduzido
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CPR_FORNECEDOR_LE_NOMEREDUZIDO", X)
'    If lErro <> SUCESSO Then Error 6722
'
'
'    Fornecedor_Le_NomeReduzido = X.Fornecedor_Le_NomeReduzido(objFornecedor)
'
'    Exit Function
'
'Erro_Fornecedor_Le_NomeReduzido:
'
'    Fornecedor_Le_NomeReduzido = Err
'
'    Select Case Err
'
'        Case 6722
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145011)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function FilialFornecedor_Le(objFilialFornecedor As ClassFilialFornecedor) As Long
''Le todos os dados da tabela FiliaisFornecedores, correspondente a chave passada por parametro
''??? eliminar FiliaisFornecedores_Le
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_FilialFornecedor_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("CPR_FILIAISFORNECEDORES_LE", X)
'    If lErro <> SUCESSO Then Error 14359
'
'
'    FilialFornecedor_Le = X.FilialFornecedor_Le(objFilialFornecedor)
'
'    Exit Function
'
'Erro_FilialFornecedor_Le:
'
'    FilialFornecedor_Le = Err
'
'    Select Case Err
'
'        Case 14359
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145012)
'
'    End Select
'
'    Exit Function
'
'End Function
'Function SldMesEst1_Le_CustoMedioConserto3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst1_Le_CustoMedioConserto3
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOCONSERTO3", X)
'    If lErro <> SUCESSO Then gError 71242
'
'    SldMesEst1_Le_CustoMedioConserto3 = X.SldMesEst1_Le_CustoMedioConserto3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst1_Le_CustoMedioConserto3:
'
'    SldMesEst1_Le_CustoMedioConserto3 = gErr
'
'    Select Case gErr
'
'        Case 71242
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145013)
'
'    End Select
'
'    Exit Function
'
'End Function
'Function SldMesEst1_Le_CustoMedioConsig3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst1_Le_CustoMedioConsig3
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOCONSIG3", X)
'    If lErro <> SUCESSO Then gError 71236
'
'    SldMesEst1_Le_CustoMedioConsig3 = X.SldMesEst1_Le_CustoMedioConsig3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst1_Le_CustoMedioConsig3:
'
'    SldMesEst1_Le_CustoMedioConsig3 = gErr
'
'    Select Case gErr
'
'        Case 71236
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145014)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst1_Le_CustoMedioBenef3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst1_Le_CustoMedioBenef3
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOBENEF3", X)
'    If lErro <> SUCESSO Then gError 71245
'
'    SldMesEst1_Le_CustoMedioBenef3 = X.SldMesEst1_Le_CustoMedioBenef3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst1_Le_CustoMedioBenef3:
'
'    SldMesEst1_Le_CustoMedioBenef3 = gErr
'
'    Select Case gErr
'
'        Case 71245
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145015)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst1_Le_CustoMedioOutros3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'
'Dim tSldMesEst1 As typeSldMesEst1
'Dim lErro As Long
'Dim iIndice As Integer
'Dim dQuant As Double
'Dim dValor As Double
'
'On Error GoTo Erro_SldMesEst1_Le_CustoMedioOutros3
'
'    'seleciona o registro de SldMesEst relativo ao Ano, FilialEmpresa e Produto passados como parametro
'    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros3, ValorInicialOutros3, SaldoQuantOutros31, SaldoValorOutros31, SaldoQuantOutros32, SaldoValorOutros32, SaldoQuantOutros33, SaldoValorOutros33, SaldoQuantOutros34, SaldoValorOutros34, SaldoQuantOutros35, SaldoValorOutros35, SaldoQuantOutros36, SaldoValorOutros36, SaldoQuantOutros37, SaldoValorOutros37, SaldoQuantOutros38, SaldoValorOutros38, SaldoQuantOutros39, SaldoValorOutros39, SaldoQuantOutros310, SaldoValorOutros310, SaldoQuantOutros311, SaldoValorOutros311, SaldoQuantOutros312, SaldoValorOutros312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
'                              tSldMesEst1.dQuantInicialOutros3, tSldMesEst1.dValorInicialOutros3, tSldMesEst1.adSaldoQuantOutros3(1), tSldMesEst1.adSaldoValorOutros3(1), tSldMesEst1.adSaldoQuantOutros3(2), tSldMesEst1.adSaldoValorOutros3(2), tSldMesEst1.adSaldoQuantOutros3(3), tSldMesEst1.adSaldoValorOutros3(3), tSldMesEst1.adSaldoQuantOutros3(4), tSldMesEst1.adSaldoValorOutros3(4), tSldMesEst1.adSaldoQuantOutros3(5), tSldMesEst1.adSaldoValorOutros3(5), tSldMesEst1.adSaldoQuantOutros3(6), tSldMesEst1.adSaldoValorOutros3(6), tSldMesEst1.adSaldoQuantOutros3(7), tSldMesEst1.adSaldoValorOutros3(7), tSldMesEst1.adSaldoQuantOutros3(8), tSldMesEst1.adSaldoValorOutros3(8), tSldMesEst1.adSaldoQuantOutros3(9), tSldMesEst1.adSaldoValorOutros3(9), _
'                              tSldMesEst1.adSaldoQuantOutros3(10), tSldMesEst1.adSaldoValorOutros3(10), tSldMesEst1.adSaldoQuantOutros3(11), tSldMesEst1.adSaldoValorOutros3(11), tSldMesEst1.adSaldoQuantOutros3(12), tSldMesEst1.adSaldoValorOutros3(12), iAno, iFilialEmpresa, sProduto)
'    If lErro <> AD_SQL_SUCESSO Then gError 71512
'
'    'le o SldMesEst
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71513
'
'    dQuant = tSldMesEst1.dQuantInicialOutros3
'    dValor = tSldMesEst1.dValorInicialOutros3
'
'    For iIndice = 1 To iMes
'
'        dQuant = dQuant + tSldMesEst1.adSaldoQuantOutros3(iIndice)
'        dValor = dValor + tSldMesEst1.adSaldoValorOutros3(iIndice)
'
'    Next
'
'    If dQuant <> 0 Then dCustoMedio = dValor / dQuant
'
'    SldMesEst1_Le_CustoMedioOutros3 = SUCESSO
'
'    Exit Function
'
'Erro_SldMesEst1_Le_CustoMedioOutros3:
'
'    SldMesEst1_Le_CustoMedioOutros3 = gErr
'
'    Select Case gErr
'
'        Case 71512, 71513
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno, iFilialEmpresa, sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145016)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst1_Le_CustoMedioDemo3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst1_Le_CustoMedioDemo3
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIODEMO3", X)
'    If lErro <> SUCESSO Then gError 71243
'
'    SldMesEst1_Le_CustoMedioDemo3 = X.SldMesEst1_Le_CustoMedioDemo3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst1_Le_CustoMedioDemo3:
'
'    SldMesEst1_Le_CustoMedioDemo3 = gErr
'
'    Select Case gErr
'
'        Case 71243
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145017)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst2_Le_CustoMedioConserto(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
''calcula o custo médio de material em conserto nosso em poder de terceiros.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst2_Le_CustoMedioConserto
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOCONSERTO", X)
'    If lErro <> SUCESSO Then gError 71237
'
'    SldMesEst2_Le_CustoMedioConserto = X.SldMesEst2_Le_CustoMedioConserto(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst2_Le_CustoMedioConserto:
'
'    SldMesEst2_Le_CustoMedioConserto = gErr
'
'    Select Case gErr
'
'        Case 71237
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145018)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst2_Le_CustoMedioDemo(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
''calcula o custo médio de material em Demonstração nosso em poder de terceiros.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst2_Le_CustoMedioDemo
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIODEMO", X)
'    If lErro <> SUCESSO Then gError 71238
'
'    SldMesEst2_Le_CustoMedioDemo = X.SldMesEst2_Le_CustoMedioDemo(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst2_Le_CustoMedioDemo:
'
'    SldMesEst2_Le_CustoMedioDemo = gErr
'
'    Select Case gErr
'
'        Case 71238
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145019)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst2_Le_CustoMedioOutros(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
''calcula o custo médio de Outros materiais nossos em poder de terceiros.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst2_Le_CustoMedioOutros
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOOUTROS", X)
'    If lErro <> SUCESSO Then gError 71239
'
'    SldMesEst2_Le_CustoMedioOutros = X.SldMesEst2_Le_CustoMedioOutros(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst2_Le_CustoMedioOutros:
'
'    SldMesEst2_Le_CustoMedioOutros = gErr
'
'    Select Case gErr
'
'        Case 71239
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145020)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst2_Le_CustoMedioBenef(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
''calcula o custo médio de materiais em beneficiamento nossos em poder de terceiros.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst2_Le_CustoMedioBenef
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOBENEF", X)
'    If lErro <> SUCESSO Then gError 71240
'
'    SldMesEst2_Le_CustoMedioBenef = X.SldMesEst2_Le_CustoMedioBenef(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst2_Le_CustoMedioBenef:
'
'    SldMesEst2_Le_CustoMedioBenef = gErr
'
'    Select Case gErr
'
'        Case 71240
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145021)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst2_Le_CustoMedioConsig(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
''calcula o custo médio de materiais em consignação nossos em poder de terceiros.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst2_Le_CustoMedioConsig
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOCONSIG", X)
'    If lErro <> SUCESSO Then gError 71241
'
'    SldMesEst2_Le_CustoMedioConsig = X.SldMesEst2_Le_CustoMedioConsig(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)
'
'    Exit Function
'
'Erro_SldMesEst2_Le_CustoMedioConsig:
'
'    SldMesEst2_Le_CustoMedioConsig = gErr
'
'    Select Case gErr
'
'        Case 71241
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145022)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Ccl_Le_Total(iTotalCcl As Integer) As Long
''Lê a quantidade de Ccl's já gravados na Tabela de Ccl e que são analíticos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Ccl_Le_Total
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CCL_LE_TOTAL", X)
'    If lErro <> SUCESSO Then gError 71270
'
'    Ccl_Le_Total = X.Ccl_Le_Total(iTotalCcl)
'
'    Exit Function
'
'Erro_Ccl_Le_Total:
'
'    Ccl_Le_Total = gErr
'
'    Select Case gErr
'
'        Case 71270
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145023)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TransacaoCTB_Le_TransacaoTela(objTransacaoCTB As ClassTransacaoCTB) As Long
''Le os dados da transacao passada como parametro e devolve os dados em objTransacaoCTB
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_TransacaoCTB_Le_TransacaoTela
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_TRANSACAOCTB_LE_TRANSACAOTELA", X)
'    If lErro <> SUCESSO Then gError 71347
'
'    TransacaoCTB_Le_TransacaoTela = X.TransacaoCTB_Le_TransacaoTela(objTransacaoCTB)
'
'    Exit Function
'
'Erro_TransacaoCTB_Le_TransacaoTela:
'
'    TransacaoCTB_Le_TransacaoTela = gErr
'
'    Select Case gErr
'
'        Case 71347
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145024)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Obtem_Subtipo_Contabil(sTabela As String, iCodigo As Integer, iSubtipo As Integer) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Obtem_Subtipo_Contabil
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_OBTEM_SUBTIPO_CONTABIL", X)
'    If lErro <> SUCESSO Then gError 71394
'
'    Obtem_Subtipo_Contabil = X.Obtem_Subtipo_Contabil(sTabela, iCodigo, iSubtipo)
'
'    Exit Function
'
'Erro_Obtem_Subtipo_Contabil:
'
'    Obtem_Subtipo_Contabil = gErr
'
'    Select Case gErr
'
'        Case 71394
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145025)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SldMesEst_Le_PrimeiroAno(objSldMesEst As ClassSldMesEst) As Long
''Le os dados iniciais do registro da tabela SldMesEst relativo ao Primeiro Ano do Produto (FilialEmpresa pega em giFilialEmpresa)
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_SldMesEst_Le_PrimeiroAno
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_SLDMESEST_LE_PRIMEIROANO", X)
'    If lErro <> SUCESSO Then gError 71395
'
'    SldMesEst_Le_PrimeiroAno = X.SldMesEst_Le_PrimeiroAno(objSldMesEst)
'
'    Exit Function
'
'Erro_SldMesEst_Le_PrimeiroAno:
'
'    SldMesEst_Le_PrimeiroAno = gErr
'
'    Select Case gErr
'
'        Case 71395
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145026)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCcl_Le_SldAnterior(iFilialEmpresa As Integer, iExercicio As Integer, sCcl As String, iPeriodo As Integer, dSaldo As Double, lComando As Long) As Long
''Le o Saldo anterior do iPeriodo passado
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCcl_Le_SldAnterior
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCCL_LE_SLDANTERIOR", X)
'    If lErro <> SUCESSO Then gError 71396
'
'    MvPerCcl_Le_SldAnterior = X.MvPerCcl_Le_SldAnterior(iFilialEmpresa, iExercicio, sCcl, iPeriodo, dSaldo, lComando)
'
'    Exit Function
'
'Erro_MvPerCcl_Le_SldAnterior:
'
'    MvPerCcl_Le_SldAnterior = gErr
'
'    Select Case gErr
'
'        Case 71396
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145027)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerCli_Le_SldAnterior(dSaldo As Double, lCliente As Long, dtData As Date, iFilialEmpresa As Integer, lComando As Long) As Long
''Le o Saldo anterior do cliente passado, a partir da tabela MvPerCli
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerCli_Le_SldAnterior
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERCLI_LE_SLDANTERIOR", X)
'    If lErro <> SUCESSO Then gError 71397
'
'    MvPerCli_Le_SldAnterior = X.MvPerCli_Le_SldAnterior(dSaldo, lCliente, dtData, iFilialEmpresa, lComando)
'
'    Exit Function
'
'Erro_MvPerCli_Le_SldAnterior:
'
'    MvPerCli_Le_SldAnterior = gErr
'
'    Select Case gErr
'
'        Case 71397
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145028)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvDiaCli_Le_SldAnterior(dSaldoDiario As Double, lCliente As Long, dtData As Date, iFilialEmpresa As Integer, lComando2 As Long) As Long
''Le o Saldo anterior do cliente passado, a partir da tabela MvDiaCli
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvDiaCli_Le_SldAnterior
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVDIACLI_LE_SLDANTERIOR", X)
'    If lErro <> SUCESSO Then gError 71398
'
'    MvDiaCli_Le_SldAnterior = X.MvDiaCli_Le_SldAnterior(dSaldoDiario, lCliente, dtData, iFilialEmpresa, lComando2)
'
'    Exit Function
'
'Erro_MvDiaCli_Le_SldAnterior:
'
'    MvDiaCli_Le_SldAnterior = gErr
'
'    Select Case gErr
'
'        Case 71398
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145029)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvPerForn_Le_SldAnterior(dSaldo As Double, lFornecedor As Long, dtData As Date, iFilialEmpresa As Integer, lComando As Long) As Long
''Le o Saldo anterior do Fornecedor passado, a partir da tabela MvPerForn
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvPerForn_Le_SldAnterior
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVPERFORN_LE_SLDANTERIOR", X)
'    If lErro <> SUCESSO Then gError 71399
'
'    MvPerForn_Le_SldAnterior = X.MvPerForn_Le_SldAnterior(dSaldo, lFornecedor, dtData, iFilialEmpresa, lComando)
'
'    Exit Function
'
'Erro_MvPerForn_Le_SldAnterior:
'
'    MvPerForn_Le_SldAnterior = gErr
'
'    Select Case gErr
'
'        Case 71399
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145030)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function MvDiaForn_Le_SldAnterior(dSaldoDiario As Double, lFornecedor As Long, dtData As Date, iFilialEmpresa As Integer, lComando2 As Long) As Long
''Le o Saldo anterior do Fornecedor passado, a partir da tabela MvDiaForn
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_MvDiaForn_Le_SldAnterior
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_MVDIAFORN_LE_SLDANTERIOR", X)
'    If lErro <> SUCESSO Then gError 71400
'
'    MvDiaForn_Le_SldAnterior = X.MvDiaForn_Le_SldAnterior(dSaldoDiario, lFornecedor, dtData, iFilialEmpresa, lComando2)
'
'    Exit Function
'
'Erro_MvDiaForn_Le_SldAnterior:
'
'    MvDiaForn_Le_SldAnterior = gErr
'
'    Select Case gErr
'
'        Case 71400
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145031)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Grava(objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Grava
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_GRAVA", X)
'    If lErro <> SUCESSO Then gError 83520
'
'    LanPendente_Grava = X.LanPendente_Grava(objLancamento_Cabecalho, colLancamento_Detalhe)
'
'    Exit Function
'
'Erro_LanPendente_Grava:
'
'    LanPendente_Grava = gErr
'
'    Select Case gErr
'
'        Case 83520
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145032)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function CustoMedio_Le(alComando() As Long, objLancamento_Detalhe As ClassLancamento_Detalhe, ByVal iApropriacao As Integer) As Long
''le o custo medio, standard ou custo dos escaninhos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CustoMedio_Le
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_CUSTOMEDIO_LE", X)
'    If lErro <> SUCESSO Then gError 83524
'
'    CustoMedio_Le = X.CustoMedio_Le(alComando, objLancamento_Detalhe, iApropriacao)
'
'    Exit Function
'
'Erro_CustoMedio_Le:
'
'    CustoMedio_Le = gErr
'
'    Select Case gErr
'
'        Case 83524
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145033)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Reprocessamento_DocOrigem(ByVal iOrigemLcto As Integer, ByVal lNumIntDocOrigem As Long, colExercicio As Collection, ByVal iFilialEmpresa As Integer) As Long
''utilizado pelo reprocessamento para alterar os lançamentos associados a docorigem que tiverem o produto preenchido e a apropriação seja pelo custo médio ou standard
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Reprocessamento_DocOrigem
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_REPROCESSAMENTO_DOCORIGEM", X)
'    If lErro <> SUCESSO Then gError 83585
'
'    Rotina_Reprocessamento_DocOrigem = X.Rotina_Reprocessamento_DocOrigem(iOrigemLcto, lNumIntDocOrigem, colExercicio, iFilialEmpresa)
'
'    Exit Function
'
'Erro_Rotina_Reprocessamento_DocOrigem:
'
'    Rotina_Reprocessamento_DocOrigem = gErr
'
'    Select Case gErr
'
'        Case 83585
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145034)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Reproc_Exclui_Lanc(iFilialEmpresa As Integer, ByVal iOrigemLcto As Integer, lNumIntDocOrigem As Long, colExercicio As Collection) As Long
''utilizado pelo reprocessamento para excluir as contabilizações associados ao ajuste do custo standard
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Reproc_Exclui_Lanc
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_REPROC_EXCLUI_LANC", X)
'    If lErro <> SUCESSO Then gError 83781
'
'    Rotina_Reproc_Exclui_Lanc = X.Rotina_Reproc_Exclui_Lanc(iFilialEmpresa, iOrigemLcto, lNumIntDocOrigem, colExercicio)
'
'    Exit Function
'
'Erro_Rotina_Reproc_Exclui_Lanc:
'
'    Rotina_Reproc_Exclui_Lanc = gErr
'
'    Select Case gErr
'
'        Case 83781
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145035)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Rotina_Reprocessamento_CProd(iFilialEmpresa As Integer, ByVal iMes As Integer, ByVal iAno As Integer, colExercicio As Collection) As Long
''altera os lançamentos do mes especificado com relação ao custo de produção
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Rotina_Reprocessamento_CProd
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_ROTINA_REPROCESSAMENTO_CPROD", X)
'    If lErro <> SUCESSO Then gError 83782
'
'    Rotina_Reprocessamento_CProd = X.Rotina_Reprocessamento_CProd(iFilialEmpresa, iMes, iAno, colExercicio)
'
'    Exit Function
'
'Erro_Rotina_Reprocessamento_CProd:
'
'    Rotina_Reprocessamento_CProd = gErr
'
'    Select Case gErr
'
'        Case 83782
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145036)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function LanPendente_Reprocessa(ByVal iOrigemLcto As Integer, ByVal lNumIntDocOrigem As Long, ByVal iFilialEmpresa As Integer) As Long
''atualiza o custo dos lançamentos pendentes que tem o documento origem passado como parametro.
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_LanPendente_Reprocessa
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANPENDENTE_REPROCESSA", X)
'    If lErro <> SUCESSO Then gError 83819
'
'    LanPendente_Reprocessa = X.LanPendente_Reprocessa(iOrigemLcto, lNumIntDocOrigem, iFilialEmpresa)
'
'    Exit Function
'
'Erro_LanPendente_Reprocessa:
'
'    LanPendente_Reprocessa = gErr
'
'    Select Case gErr
'
'        Case 83819
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145037)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Le_Doc1(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
''Tenta ler o documento em questão.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Doc1
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_DOC1", X)
'    If lErro <> SUCESSO Then gError 83871
'
'    Lancamento_Le_Doc1 = X.Lancamento_Le_Doc1(objLancamento_Cabecalho)
'
'    Exit Function
'
'Erro_Lancamento_Le_Doc1:
'
'    Lancamento_Le_Doc1 = gErr
'
'    Select Case gErr
'
'        Case 83871
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145038)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Altera_Status(iFilialEmpresa As Integer, iTransacao As Integer, lNumIntDoc As Long, colLancamentos As Collection, ByVal iStatus As Integer) As Long
''le os lançamentos relativos ao documento passado como parametro ,coloca-os em colLancamentos e atualiza o status
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Altera_Status
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_ALTERA_STATUS", X)
'    If lErro <> SUCESSO Then gError 83879
'
'    Lancamento_Altera_Status = X.Lancamento_Altera_Status(iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos, iStatus)
'
'    Exit Function
'
'Erro_Lancamento_Altera_Status:
'
'    Lancamento_Altera_Status = gErr
'
'    Select Case gErr
'
'        Case 83879
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145039)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Exclui_30(iFilialEmpresa As Integer, iTransacao As Integer, lNumIntDoc As Long, colLancamentos As Collection) As Long
''Exclui os lançamentos da transação em questão e acerta os totais dos lotes envolvidos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Exclui_30
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_EXCLUI_30", X)
'    If lErro <> SUCESSO Then gError 83894
'
'    Lancamento_Exclui_30 = X.Lancamento_Exclui_30(iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos)
'
'    Exit Function
'
'Erro_Lancamento_Exclui_30:
'
'    Lancamento_Exclui_30 = gErr
'
'    Select Case Err
'
'        Case 83894
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145040)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Le_Doc3(iFilialEmpresa As Integer, iTransacao As Integer, lNumIntDoc As Long, colLancamentos As Collection) As Long
''le os lançamentos relativos ao documento passado como parametro inclusive os com status <> 0 e coloca-os em colLancamentos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Le_Doc3
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_LE_DOC3", X)
'    If lErro <> SUCESSO Then gError 89102
'
'    Lancamento_Le_Doc3 = X.Lancamento_Le_Doc3(iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos)
'
'    Exit Function
'
'Erro_Lancamento_Le_Doc3:
'
'    Lancamento_Le_Doc3 = gErr
'
'    Select Case gErr
'
'        Case 89102
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145041)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Inicializa_Mascara_Ccl_MaskEd(objMaskEdCcl As Object) As Long
''inicializa a mascara de centro de custo
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Inicializa_Mascara_Ccl_MaskEd
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_INICIALIZA_MASCARA_CCL_MASKED", X)
'    If lErro <> SUCESSO Then gError 89254
'
'    Inicializa_Mascara_Ccl_MaskEd = X.Inicializa_Mascara_Ccl_MaskEd(objMaskEdCcl)
'    Exit Function
'
'Erro_Inicializa_Mascara_Ccl_MaskEd:
'
'    Inicializa_Mascara_Ccl_MaskEd = gErr
'
'    Select Case gErr
'
'        Case 89254
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145042)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Lancamento_Altera_Data(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, dtData As Date, dtDataEstoque As Date) As Long
''Exclui os lançamentos da transação em questão, acerta os totais dos lotes envolvidos e inclui um novo lançamento na nova data
''AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Lancamento_Altera_Data
'
'    lErro = Rotina_ObterObjetoFuncao("CTB_LANCAMENTO_ALTERA_DATA", X)
'    If lErro <> SUCESSO Then gError 89709
'
'    Lancamento_Altera_Data = X.Lancamento_Altera_Data(iFilialEmpresa, iTransacao, lNumIntDoc, dtData, dtDataEstoque)
'
'    Exit Function
'
'Erro_Lancamento_Altera_Data:
'
'    Lancamento_Altera_Data = gErr
'
'    Select Case gErr
'
'        Case 89709
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145043)
'
'
'    End Select
'
'    Exit Function
'
'End Function
'
