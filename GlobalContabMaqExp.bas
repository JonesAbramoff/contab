Attribute VB_Name = "GlobalContabMaqExp"
Option Explicit


Declare Function Compila_Formula_Contabil Lib "ADMAQEXP.DLL" Alias "AD_ME_Compila_Formula_Contabil" (ByVal objExeExp As ClassExeExp, ByVal MaqExp_Testa_Mnemonico As Long, ByVal sFormula As String, iTipo As Integer, iInicio As Integer, iTam As Integer) As Long

Declare Function Executa_Formula_Contabil Lib "ADMAQEXP.DLL" Alias "AD_ME_Executa_Formula_Contabil" (ByVal objExeExp As ClassExeExp, ByVal MaqExp_Grid_Linhas As Long, ByVal lExpComp As Long, ByVal MaqExp_Armazena_Campo As Long, ByVal MaqExp_Devolve_Valor_Mnemonico As Long, ByVal MaqExp_Devolve_Valor_Total_Mnemonico As Long) As Long

Declare Function Executa_Formula_Contabil_Comissoes Lib "ADMAQEXP.DLL" Alias "AD_ME_Executa_Formula_Contabil_Comissoes" (ByVal objExeExp As ClassExeExp, ByVal MaqExp_Grid_Linhas As Long, ByVal lExpComp As Long, ByVal MaqExp_Armazena_Campo As Long, ByVal MaqExp_Devolve_Valor_Mnemonico As Long, ByVal MaqExp_Devolve_Valor_Total_Mnemonico As Long, ByVal iIndice_Regra As Integer) As Long
'iIndice_Regra indica a linha das regras de comissoes que está sendo executado

Declare Function Devolve_Identificador_Mnemonico Lib "ADMAQEXP.DLL" Alias "AD_ME_Devolve_Identificador_Mnemonico" (ByVal sIdentificador As String, ByVal lpRel As Long) As Long

Declare Function Devolve_Identificador_Grid Lib "ADMAQEXP.DLL" Alias "AD_ME_Devolve_Identificador_Grid" (ByVal sIdentificador As String, ByVal lpRel As Long) As Long

Declare Function Devolve_Nome_Grid Lib "ADMAQEXP.DLL" Alias "AD_ME_Devolve_Nome_Grid" (ByVal sNomeGrid As String, ByVal lNomeGrid As Long) As Long

Declare Function Devolve_Valor_Campo_String Lib "ADMAQEXP.DLL" Alias "AD_ME_Devolve_Valor_Campo_String" (ByVal sValor As String, ByVal lpRel As Long) As Long

Declare Function Devolve_Valor_Param_String Lib "ADMAQEXP.DLL" Alias "AD_ME_Devolve_Valor_Param_String" (ByVal sValor As String, ByVal lpRel As Long, ByVal iParam As Integer) As Long

Declare Function Devolve_Valor_Param_Double Lib "ADMAQEXP.DLL" Alias "AD_ME_Devolve_Valor_Param_Double" (dValor As Double, ByVal lpRel As Long, ByVal iParam As Integer) As Long

Declare Function Envia_Valor_Mnemonico_String Lib "ADMAQEXP.DLL" Alias "AD_ME_Envia_Valor_Mnemonico_String" (ByVal sValor As String, ByVal lpRel As Long) As Long

Declare Function Envia_Valor_Mnemonico_Double Lib "ADMAQEXP.DLL" Alias "AD_ME_Envia_Valor_Mnemonico_Double" (ByVal dValor As Double, ByVal lpRel As Long) As Long

Declare Function Inicializa_Formula_Contabil Lib "ADMAQEXP.DLL" Alias "AD_ME_Inicializa_Formula_Contabil" (lExpComp As Long) As Long

Declare Function Inicializa_Formula_Contabil_Comissoes Lib "ADMAQEXP.DLL" Alias "AD_ME_Inicializa_Formula_Contabil_Comissoes" (lExpComp As Long) As Long
'difere da funcao original pelo seu conteudo. Os parametros sao os mesmos. Nesta funcao só descobre o numero de linhas dos grids envolvidos em todas as regras. Nao aloca memoria.

Declare Function Finaliza_Formula_Contabil Lib "ADMAQEXP.DLL" Alias "AD_ME_Finaliza_Formula_Contabil" (ByVal lExpComp As Long) As Long

Declare Function Finaliza_Formula_Contabil_Comissoes Lib "ADMAQEXP.DLL" Alias "AD_ME_Finaliza_Formula_Contabil_Comissoes" (ByVal lExpComp As Long) As Long
'difere da funcao original pelo seu conteudo. As desalocacoes de memoria sao feitas em funçao da quantidade alocada e nao fixo como a funcao original.

Declare Function CompilaExe_Formula_Contabil Lib "ADMAQEXP.DLL" Alias "AD_ME_CompilaExe_Formula_Contabil" (ByVal objExeExp As ClassExeExp, ByVal MaqExp_Testa_Mnemonico As Long, ByVal sFormula As String, iCampo As Integer, ByVal lExpComp As Long) As Long

Declare Function CompilaExe_Formula_Contabil_Comissoes Lib "ADMAQEXP.DLL" Alias "AD_ME_CompilaExe_Formula_Contabil_Comissoes" (ByVal objExeExp As ClassExeExp, ByVal MaqExp_Testa_Mnemonico As Long, ByVal sFormula As String, iCampo As Integer, ByVal lExpComp As Long) As Long
'difere da funcao original pois Aloca memoria para cada expressao compilada. Na original a alocacao é fixa e feita na funcao de inicializacao da execucao.

Declare Function Descobre_Numero_Linhas_Grids Lib "ADMAQEXP.DLL" Alias "AD_ME_Descobre_Numero_Linhas_Grids" (ByVal objExeExp As ClassExeExp, ByVal MaqExp_Grid_Linhas As Long, ByVal lExpComp As Long) As Long

Function MaqExp_Testa_Mnemonico(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal lNomeGrid As Long, iTipo As Integer, iNumParam As Integer, iParam1 As Integer, iParam2 As Integer, iParam3 As Integer, ByVal iInicio_Expressao As Integer) As Long

Dim lErro As Long
Dim sIdentificador As String
Dim objMnemonico As Object
Dim iIndice As Integer

On Error GoTo Erro_MaqExp_Testa_Mnemonico

    sIdentificador = String(255, 0)
    
    lErro = Devolve_Identificador_Mnemonico(sIdentificador, lpRel)
    If lErro <> SUCESSO Then Error 36019

    sIdentificador = StringZ(sIdentificador)
    
    Set objMnemonico = Nothing
    
    For iIndice = 1 To objExeExp.colMnemonico.Count
        If objExeExp.colMnemonico.Item(iIndice).sMnemonico = sIdentificador Then
            Set objMnemonico = objExeExp.colMnemonico.Item(iIndice)
            Exit For
        End If
    Next
    
    If objMnemonico Is Nothing Then Error 36203
    
    iTipo = objMnemonico.iTipo
    iNumParam = objMnemonico.iNumParam
    iParam1 = objMnemonico.iParam1
    iParam2 = objMnemonico.iParam2
    iParam3 = objMnemonico.iParam3
    
    lErro = Devolve_Nome_Grid(objMnemonico.sGrid, lNomeGrid)
    If lErro <> SUCESSO Then Error 36204
    
    MaqExp_Testa_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Testa_Mnemonico:

    MaqExp_Testa_Mnemonico = Err

    Select Case Err
    
        Case 36019, 36203, 36204
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161610)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Grid_Linhas(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, iNumLinhas As Integer) As Long
'retorna o numero de linhas do grid especificado.

Dim lErro As Long
Dim sIdentificador As String
Dim objMnemonico As ClassMnemonico
Dim objGridTransacao As ClassGridTransacao
Dim iAchou As Integer

On Error GoTo Erro_MaqExp_Grid_Linhas


    sIdentificador = String(255, 0)
    
    lErro = Devolve_Identificador_Grid(sIdentificador, lpRel)
    If lErro <> SUCESSO Then Error 36061

    sIdentificador = StringZ(sIdentificador)

    iNumLinhas = 1

    For Each objGridTransacao In objExeExp.colGridTransacao
    
        If objGridTransacao.sNomeGrid = sIdentificador Then
    
            iAchou = 1
            iNumLinhas = objGridTransacao.iNumLinhas
            Exit For
            
        End If
    
    Next
    
    'se não achou o grid
    If iAchou = 0 Then
    
        'verifica se o grid não é informado na tela em questão
        lErro = objExeExp.objTransacao.Calcula_Grid(sIdentificador, iNumLinhas)
        If lErro <> SUCESSO Then Error 36062
    
    End If
    
    MaqExp_Grid_Linhas = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Grid_Linhas:

    MaqExp_Grid_Linhas = Err

    Select Case Err
    
        Case 36061
    
        Case 36062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_ENCONTRADO", Err, sIdentificador)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161611)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Armazena_Campo(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iExp As Integer, ByVal iCampo As Integer, ByVal dValor As Double) As Long
'armazena em colLancamentos (item de objExeExp) o valor que vai ser retornado pela chamada da funcao Devolve_Valor_Formula. iExp indica se é o primeiro valor da linha que está sendo retornado (= 0) ou não.

Dim lErro As Long
Dim sValor As String
Dim objLancamentos As ClassLancamentos
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_MaqExp_Armazena_Campo

    ' Se é o primeiro campo de uma nova linha ==> tem que criar uma nova instancia de classLancamentos.
    If iExp = 0 Then
    
        Set objLancamentos = New ClassLancamentos
        objExeExp.colLancamentos.Add objLancamentos
        
    Else
    
        Set objLancamentos = objExeExp.colLancamentos.Item(objExeExp.colLancamentos.Count)
        
    End If

    'se o campo retorna uma string ==> chama a função que pega o valor do campo.
    'se o campo retorna um double este é passado como parametro na chamada da funcao (dValor)
    If iCampo = CAMPO_CONTA Or iCampo = CAMPO_CCL Or iCampo = CAMPO_HISTORICO Or iCampo = CAMPO_PRODUTO Then

        sValor = String(255, 0)
    
        lErro = Devolve_Valor_Campo_String(sValor, lpRel)
        If lErro <> SUCESSO Then Error 36063

        sValor = StringZ(sValor)
        
    End If
    
    Select Case iCampo
    
        Case CAMPO_CONTA
            lErro = CF("Conta_Formata", sValor, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then
                objLancamentos.sConta = ""
            Else
                objLancamentos.sConta = sValor
            End If
            
        Case CAMPO_CCL
            lErro = CF("Ccl_Formata", sValor, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then
                objLancamentos.sCcl = ""
            Else
                objLancamentos.sCcl = sValor
            End If
            
        Case CAMPO_CREDITO
            objLancamentos.dCredito = dValor
            
        Case CAMPO_DEBITO
            objLancamentos.dDebito = dValor
            
        Case CAMPO_HISTORICO
            objLancamentos.sHistorico = sValor
            
        Case CAMPO_PRODUTO
            lErro = CF("Produto_Formata", sValor, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then
                objLancamentos.sProduto = ""
            Else
                objLancamentos.sProduto = sValor
            End If
            
        Case CAMPO_AGLUTINA
            objLancamentos.iAglutina = dValor
            
        Case CAMPO_GERENCIAL
            objLancamentos.iGerencial = dValor
            
        Case CAMPO_ESCANINHO_CUSTO
            objLancamentos.iEscaninho_Custo = dValor
            
    End Select
    
    MaqExp_Armazena_Campo = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Armazena_Campo:

    MaqExp_Armazena_Campo = Err

    Select Case Err
    
        Case 36063
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161612)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Devolve_Valor_Mnemonico(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iIndice As Integer, ByVal iNumParam As Integer, ByVal iTipoParam1 As Integer, ByVal iTipoParam2 As Integer, ByVal iTipoParam3 As Integer) As Long
'devolve o valor do mnemonico passado como parametro

Dim lErro As Long
Dim sValor As String
Dim sMnemonico As String
Dim iParam As Integer
Dim aiTipoParam(1 To 3) As Integer
Dim avParam(1 To 3) As Variant
Dim dValor As Double
Dim objMnemonicoValor As ClassMnemonicoValor
Dim vValor As Variant
Dim objContabil As New ClassContabil

On Error GoTo Erro_MaqExp_Devolve_Valor_Mnemonico

    sMnemonico = String(255, 0)
    
    'recupera o nome do mnemonico
    lErro = Devolve_Identificador_Mnemonico(sMnemonico, lpRel)
    If lErro <> SUCESSO Then gError 36067

    sMnemonico = StringZ(sMnemonico)

    aiTipoParam(1) = iTipoParam1
    aiTipoParam(2) = iTipoParam2
    aiTipoParam(3) = iTipoParam3
    
    'recupera o valor dos parametros (se houverem)
    For iParam = 1 To iNumParam
    
        If aiTipoParam(iParam) = TIPO_TEXTO Then
        
            sValor = String(255, 0)
    
            lErro = Devolve_Valor_Param_String(sValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 36068

            avParam(iParam) = StringZ(sValor)
        
        
        Else

            lErro = Devolve_Valor_Param_Double(dValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 36069
            
            avParam(iParam) = dValor
            
        End If
        
    Next
    
    lErro = Procura_Colecao_Mnemonico(sMnemonico, iNumParam, avParam(), objMnemonicoValor, objExeExp.colMnemonicoValor)
    If lErro <> SUCESSO Then gError 36070
    
    'se o mnemonico ainda não foi calculado ==> calcula-o
    If objMnemonicoValor Is Nothing Then
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        
        Set objMnemonicoValor.colValor = New Collection
        objMnemonicoValor.sMnemonico = sMnemonico
        objMnemonicoValor.iParam = iNumParam
        
        For iParam = 1 To iNumParam
        
            objMnemonicoValor.vParam(iParam) = avParam(iParam)
            
        Next
    
        If objExeExp.objContexto Is Nothing Then
    
            lErro = objExeExp.objTransacao.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO And lErro <> CONTABIL_MNEMONICO_NAO_ENCONTRADO Then gError 36071
        
        Else
        
            lErro = objExeExp.objTransacao.Calcula_Mnemonico(objMnemonicoValor, objExeExp.objContexto)
            If lErro <> SUCESSO And lErro <> CONTABIL_MNEMONICO_NAO_ENCONTRADO Then gError 178250
        
        End If
        
        If lErro = CONTABIL_MNEMONICO_NAO_ENCONTRADO Then
        
            lErro = objContabil.Contabil_Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 36711
        
        End If
        
        'armazena os valores do mnemonico
        objExeExp.colMnemonicoValor.Add objMnemonicoValor
        
    End If
    
    'se o indice do elemento procurado for maior do que o numero de elementos
    If iIndice > objMnemonicoValor.colValor.Count Then gError 55858
    
    vValor = objMnemonicoValor.colValor.Item(iIndice)

    If VarType(vValor) = vbString Then
    
        lErro = Envia_Valor_Mnemonico_String(vValor, lpRel)
        If lErro <> SUCESSO Then gError 36073
        
    ElseIf VarType(vValor) = vbDate Then
            
        vValor = vValor + 693594
            
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 36074
    
    Else
    
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 36074
    
    End If

    MaqExp_Devolve_Valor_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Devolve_Valor_Mnemonico:

    MaqExp_Devolve_Valor_Mnemonico = gErr

    Select Case gErr
    
        Case 36067, 36068, 36069, 36070, 36071, 36073, 36074, 36711, 178250
    
        Case 55858
            Call Rotina_Erro(vbOKOnly, "ERRO_INDICE_MAIOR_LINHAS_GRID", gErr, sMnemonico, iIndice, objMnemonicoValor.colValor.Count)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161613)
    
    End Select
    
    Exit Function
    
End Function

Function MaqExp_Devolve_Valor_Total_Mnemonico(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iIndice As Integer, ByVal iNumParam As Integer, ByVal iTipoParam1 As Integer, ByVal iTipoParam2 As Integer, ByVal iTipoParam3 As Integer) As Long
'devolve o valor do mnemonico passado como parametro somando todas as linhas do grid

Dim lErro As Long
Dim sValor As String
Dim sMnemonico As String
Dim iParam As Integer
Dim aiTipoParam(1 To 3) As Integer
Dim avParam(1 To 3) As Variant
Dim dValor As Double
Dim objMnemonicoValor As ClassMnemonicoValor
Dim vValor As Variant
Dim objContabil As New ClassContabil

On Error GoTo Erro_MaqExp_Devolve_Valor_Total_Mnemonico

    sMnemonico = String(255, 0)
    
    'recupera o nome do mnemonico
    lErro = Devolve_Identificador_Mnemonico(sMnemonico, lpRel)
    If lErro <> SUCESSO Then Error 60806

    sMnemonico = StringZ(sMnemonico)

    aiTipoParam(1) = iTipoParam1
    aiTipoParam(2) = iTipoParam2
    aiTipoParam(3) = iTipoParam3
    
    'recupera o valor dos parametros (se houverem)
    For iParam = 1 To iNumParam
    
        If aiTipoParam(iParam) = TIPO_TEXTO Then
        
            sValor = String(255, 0)
    
            lErro = Devolve_Valor_Param_String(sValor, lpRel, iParam)
            If lErro <> SUCESSO Then Error 60807

            avParam(iParam) = StringZ(sValor)
        
        
        Else

            lErro = Devolve_Valor_Param_Double(dValor, lpRel, iParam)
            If lErro <> SUCESSO Then Error 60808
            
            avParam(iParam) = dValor
            
        End If
        
    Next
    
    lErro = Procura_Colecao_Mnemonico(sMnemonico, iNumParam, avParam(), objMnemonicoValor, objExeExp.colMnemonicoValor)
    If lErro <> SUCESSO Then Error 60809
    
    'se o mnemonico ainda não foi calculado ==> calcula-o
    If objMnemonicoValor Is Nothing Then
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        
        Set objMnemonicoValor.colValor = New Collection
        objMnemonicoValor.sMnemonico = sMnemonico
        objMnemonicoValor.iParam = iNumParam
        
        For iParam = 1 To iNumParam
        
            objMnemonicoValor.vParam(iParam) = avParam(iParam)
            
        Next
    
        lErro = objExeExp.objTransacao.Calcula_Mnemonico(objMnemonicoValor)
        If lErro <> SUCESSO And lErro <> CONTABIL_MNEMONICO_NAO_ENCONTRADO Then Error 60810
        
        If lErro = CONTABIL_MNEMONICO_NAO_ENCONTRADO Then
        
            lErro = objContabil.Contabil_Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then Error 60811
        
        End If
        
        'armazena os valores do mnemonico
        objExeExp.colMnemonicoValor.Add objMnemonicoValor
        
    End If
    
    vValor = 0
    
    'se o indice do elemento procurado for maior do que o numero de elementos
    If iIndice <= objMnemonicoValor.colValor.Count Then
    
        For iIndice = 1 To objMnemonicoValor.colValor.Count
            vValor = vValor + objMnemonicoValor.colValor.Item(iIndice)
        Next
    
    End If
   
    lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
    If lErro <> SUCESSO Then Error 60815
    
    MaqExp_Devolve_Valor_Total_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Devolve_Valor_Total_Mnemonico:

    MaqExp_Devolve_Valor_Total_Mnemonico = Err

    Select Case Err
    
        Case 60806, 60807, 60808, 60809, 60810, 60811, 60813, 60814, 60815
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161614)
    
    End Select
    
    Exit Function
    
End Function

Function MaqExp_Devolve_Valor_Total_MnemonicoComiss(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iIndice As Integer, ByVal iNumParam As Integer, ByVal iTipoParam1 As Integer, ByVal iTipoParam2 As Integer, ByVal iTipoParam3 As Integer) As Long
'devolve o valor do mnemonico passado como parametro somando todas as linhas do grid
'Versao especial para comissoes
'tulio070203

Dim lErro As Long
Dim sValor As String
Dim sMnemonico As String
Dim iParam As Integer
Dim aiTipoParam(1 To 3) As Integer
Dim avParam(1 To 3) As Variant
Dim dValor As Double
Dim objMnemonicoValor As ClassMnemonicoValor
Dim vValor As Variant
Dim objContabil As New ClassContabil

On Error GoTo Erro_MaqExp_Devolve_Valor_Total_MnemonicoComiss

    sMnemonico = String(255, 0)
    
    'recupera o nome do mnemonico
    lErro = Devolve_Identificador_Mnemonico(sMnemonico, lpRel)
    If lErro <> SUCESSO Then gError 111765

    sMnemonico = StringZ(sMnemonico)

    aiTipoParam(1) = iTipoParam1
    aiTipoParam(2) = iTipoParam2
    aiTipoParam(3) = iTipoParam3
    
    'recupera o valor dos parametros (se houverem)
    For iParam = 1 To iNumParam
    
        If aiTipoParam(iParam) = TIPO_TEXTO Then
        
            sValor = String(255, 0)
    
            lErro = Devolve_Valor_Param_String(sValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 111766

            avParam(iParam) = StringZ(sValor)
        
        
        Else

            lErro = Devolve_Valor_Param_Double(dValor, lpRel, iParam)
            If lErro <> SUCESSO Then Error 111767
            
            avParam(iParam) = dValor
            
        End If
        
    Next
    
    lErro = Procura_Colecao_Mnemonico(sMnemonico, iNumParam, avParam(), objMnemonicoValor, objExeExp.colMnemonicoValor)
    If lErro <> SUCESSO Then gError 111772
    
    'se o mnemonico ainda não foi calculado ==> calcula-o
    If objMnemonicoValor Is Nothing Then
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        
        Set objMnemonicoValor.colValor = New Collection
        objMnemonicoValor.sMnemonico = sMnemonico
        objMnemonicoValor.iParam = iNumParam
        
        For iParam = 1 To iNumParam
        
            objMnemonicoValor.vParam(iParam) = avParam(iParam)
            
        Next
    
        'chama as funcao que calcula os mnemonicos customizados
        lErro = objExeExp.objTransacao.objmnemonicoComissCalcAux.Calcula_Mnemonico_Comissoes(objMnemonicoValor)
        If lErro <> SUCESSO And lErro <> MNEMONICOCOMISSOES_NAO_ENCONTRADO Then gError 111769
        
        'se nao encontrou entre os mnemonicos customizados
        If lErro = MNEMONICOCOMISSOES_NAO_ENCONTRADO Then
        
            lErro = objExeExp.objTransacao.objMnemonicoComissCalc.Calcula_Mnemonico_Comissoes(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 111768
        
        End If
        
        'se nao encontrou o mnemonico => erro
        If lErro = MNEMONICOCOMISSOES_NAO_ENCONTRADO Then gError 111770
        
        'armazena os valores do mnemonico
        objExeExp.colMnemonicoValor.Add objMnemonicoValor
        
    End If
    
    vValor = 0
    
    'se o indice do elemento procurado for maior do que o numero de elementos
    If iIndice <= objMnemonicoValor.colValor.Count Then
    
        For iIndice = 1 To objMnemonicoValor.colValor.Count
            vValor = vValor + objMnemonicoValor.colValor.Item(iIndice)
        Next
    
    End If
   
    lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
    If lErro <> SUCESSO Then gError 111771
    
    MaqExp_Devolve_Valor_Total_MnemonicoComiss = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Devolve_Valor_Total_MnemonicoComiss:

    MaqExp_Devolve_Valor_Total_MnemonicoComiss = gErr

    Select Case gErr
    
        Case 111765 To 111772
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161615)
    
    End Select
    
    Exit Function
    
End Function

Private Function Procura_Colecao_Mnemonico(sMnemonico As String, ByVal iNumParam As Integer, avParam() As Variant, objMnemonicoValor As ClassMnemonicoValor, colMnemonicoValor As ClassColMnemonicoValor) As Long
'descobre os grids da tela em questão e coloca-os em colGridTransacao

Dim lErro As Long
Dim iParam As Integer
Dim iAchou As Integer

On Error GoTo Erro_Procura_Colecao_Mnemonico

    'pesquisa o mnemonico na colecao
    For Each objMnemonicoValor In colMnemonicoValor

        If objMnemonicoValor.sMnemonico = sMnemonico Then
    
            'se não tiver parametros ==> encontrou o mnemonico
            If iNumParam = 0 Then
                iAchou = 1
            Else
            
                iAchou = 1
                For iParam = 1 To iNumParam
                    If avParam(iParam) <> objMnemonicoValor.vParam(iParam) Then
                        iAchou = 0
                        Exit For
                    End If
                Next
            End If
            
            If iAchou = 1 Then Exit For
            
        End If
        
    Next
    
    If iAchou = 0 Then Set objMnemonicoValor = Nothing
                    
    Procura_Colecao_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Procura_Colecao_Mnemonico:

    Procura_Colecao_Mnemonico = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161616)
        
    End Select
    
    Exit Function
    
End Function

Function MaqExp_Testa_Mnemonico_FPreco(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal lNomeGrid As Long, iTipo As Integer, iNumParam As Integer, iParam1 As Integer, iParam2 As Integer, iParam3 As Integer, ByVal iInicio_Expressao As Integer) As Long

Dim lErro As Long
Dim sIdentificador As String
Dim objMnemonico As ClassMnemonicoFPreco
Dim iIndice As Integer

On Error GoTo Erro_MaqExp_Testa_Mnemonico_FPreco


    sIdentificador = String(255, 0)
    
    lErro = Devolve_Identificador_Mnemonico(sIdentificador, lpRel)
    If lErro <> SUCESSO Then gError 92272

    sIdentificador = StringZ(sIdentificador)
    
    If sIdentificador Like "L#" Or sIdentificador Like "L##" Then
        If objExeExp.iLinhaAtual <= CInt(Mid(sIdentificador, 2)) Then gError 92271
        
        iTipo = TIPO_NUMERICO
        iNumParam = 0
    
    Else
    
        Set objMnemonico = Nothing
        
        For iIndice = 1 To objExeExp.colMnemonico.Count
            If objExeExp.colMnemonico.Item(iIndice).sMnemonico = sIdentificador Then
                Set objMnemonico = objExeExp.colMnemonico.Item(iIndice)
                Exit For
            End If
        Next
        
        If objMnemonico Is Nothing Then gError 92273
        
        iTipo = objMnemonico.iTipo
        iNumParam = objMnemonico.iNumParam
        iParam1 = objMnemonico.iParam1
        iParam2 = objMnemonico.iParam2
        iParam3 = objMnemonico.iParam3
    
    End If
    
    MaqExp_Testa_Mnemonico_FPreco = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Testa_Mnemonico_FPreco:

    MaqExp_Testa_Mnemonico_FPreco = gErr

    Select Case gErr
    
        Case 92271, 92272, 92273
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161617)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Testa_Mnemonico_FPreco1(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal lNomeGrid As Long, iTipo As Integer, iNumParam As Integer, iParam1 As Integer, iParam2 As Integer, iParam3 As Integer, ByVal iInicio_Expressao As Integer) As Long

Dim lErro As Long
Dim sIdentificador As String
Dim objMnemonico As Object
Dim iIndice As Integer

On Error GoTo Erro_MaqExp_Testa_Mnemonico_FPreco1

    sIdentificador = String(255, 0)
    
    lErro = Devolve_Identificador_Mnemonico(sIdentificador, lpRel)
    If lErro <> SUCESSO Then gError 92290

    sIdentificador = StringZ(sIdentificador)
    
    If (sIdentificador Like "L#" Or sIdentificador Like "L##") Then
    
        If objExeExp.iInicio_Expressao <> iInicio_Expressao Then
    
            objExeExp.iInicio_Expressao = iInicio_Expressao
            
            If objExeExp.iLinha = CInt(Mid(sIdentificador, 2)) Then gError 92291
            
            If CInt(Mid(sIdentificador, 2)) > objExeExp.iLinha Then
            
                objExeExp.sExpressao1 = Mid(objExeExp.sExpressao1, 1, iInicio_Expressao + objExeExp.iCaracteres) & "L" & CStr(CInt(Mid(sIdentificador, 2)) - 1) & Mid(objExeExp.sExpressao1, iInicio_Expressao + Len(sIdentificador) + 1 + objExeExp.iCaracteres)
            
                If Len(Mid(sIdentificador, 2)) > Len(CStr(CInt(Mid(sIdentificador, 2)) - 1)) Then
                    objExeExp.iCaracteres = objExeExp.iCaracteres - 1
                End If
            
            End If
        
        End If
        
        iTipo = TIPO_NUMERICO
        iNumParam = 0
    
    Else
    
        Set objMnemonico = Nothing
        
        For iIndice = 1 To objExeExp.colMnemonico.Count
            If objExeExp.colMnemonico.Item(iIndice).sMnemonico = sIdentificador Then
                Set objMnemonico = objExeExp.colMnemonico.Item(iIndice)
                Exit For
            End If
        Next
        
        If objMnemonico Is Nothing Then gError 92292
        
        iTipo = objMnemonico.iTipo
        iNumParam = objMnemonico.iNumParam
        iParam1 = objMnemonico.iParam1
        iParam2 = objMnemonico.iParam2
        iParam3 = objMnemonico.iParam3
    
    End If
    
    MaqExp_Testa_Mnemonico_FPreco1 = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Testa_Mnemonico_FPreco1:

    MaqExp_Testa_Mnemonico_FPreco1 = gErr

    Select Case gErr
    
        Case 92290, 92292
    
        Case 92291
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FPRECO_LINHA_REFERENCIADA", gErr, objExeExp.iLinha, objExeExp.iLinhaAtual, sIdentificador)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161618)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Testa_Mnemonico_FPreco2(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal lNomeGrid As Long, iTipo As Integer, iNumParam As Integer, iParam1 As Integer, iParam2 As Integer, iParam3 As Integer, ByVal iInicio_Expressao As Integer) As Long

Dim lErro As Long
Dim sIdentificador As String
Dim objMnemonico As Object
Dim iIndice As Integer

On Error GoTo Erro_MaqExp_Testa_Mnemonico_FPreco2

    sIdentificador = String(255, 0)

    lErro = Devolve_Identificador_Mnemonico(sIdentificador, lpRel)
    If lErro <> SUCESSO Then gError 92294

    sIdentificador = StringZ(sIdentificador)

    If sIdentificador Like "L#" Or sIdentificador Like "L##" Then

        If objExeExp.iInicio_Expressao <> iInicio_Expressao Then

            objExeExp.iInicio_Expressao = iInicio_Expressao

            If CInt(Mid(sIdentificador, 2)) >= objExeExp.iLinha Then

                objExeExp.sExpressao1 = Mid(objExeExp.sExpressao1, 1, iInicio_Expressao + objExeExp.iCaracteres) & "L" & CStr(CInt(Mid(sIdentificador, 2)) + 1) & Mid(objExeExp.sExpressao1, iInicio_Expressao + Len(sIdentificador) + 1 + objExeExp.iCaracteres)

                If Len(Mid(sIdentificador, 2)) < Len(CStr(CInt(Mid(sIdentificador, 2)) + 1)) Then
                    objExeExp.iCaracteres = objExeExp.iCaracteres + 1
                End If

            End If

        End If

        iTipo = TIPO_NUMERICO
        iNumParam = 0

    Else

        Set objMnemonico = Nothing

        For iIndice = 1 To objExeExp.colMnemonico.Count
            If objExeExp.colMnemonico.Item(iIndice).sMnemonico = sIdentificador Then
                Set objMnemonico = objExeExp.colMnemonico.Item(iIndice)
                Exit For
            End If
        Next

        If objMnemonico Is Nothing Then gError 92292

        iTipo = objMnemonico.iTipo
        iNumParam = objMnemonico.iNumParam
        iParam1 = objMnemonico.iParam1
        iParam2 = objMnemonico.iParam2
        iParam3 = objMnemonico.iParam3

    End If

    MaqExp_Testa_Mnemonico_FPreco2 = SUCESSO

    Exit Function

Erro_MaqExp_Testa_Mnemonico_FPreco2:

    MaqExp_Testa_Mnemonico_FPreco2 = gErr

    Select Case gErr

        Case 92294, 92295

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161619)

    End Select

    Exit Function

End Function


Function MaqExp_Devolve_Valor_Mnemon_FPreco(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iIndice As Integer, ByVal iNumParam As Integer, ByVal iTipoParam1 As Integer, ByVal iTipoParam2 As Integer, ByVal iTipoParam3 As Integer) As Long
'devolve o valor do mnemonico passado como parametro

Dim lErro As Long
Dim sValor As String
Dim sMnemonico As String
Dim iParam As Integer
Dim aiTipoParam(1 To 3) As Integer
Dim avParam(1 To 3) As Variant
Dim dValor As Double
Dim objMnemonicoValor As ClassMnemonicoValor
Dim vValor As Variant
Dim objContabil As New ClassContabil
Dim objMnemonicoFPreco As Object
Dim iAchou As Integer
Dim sProduto As String

On Error GoTo Erro_MaqExp_Devolve_Valor_Mnemon_FPreco

    sMnemonico = String(255, 0)
    
    'recupera o nome do mnemonico
    lErro = Devolve_Identificador_Mnemonico(sMnemonico, lpRel)
    If lErro <> SUCESSO Then gError 92273

    sMnemonico = StringZ(sMnemonico)

    aiTipoParam(1) = iTipoParam1
    aiTipoParam(2) = iTipoParam2
    aiTipoParam(3) = iTipoParam3
    
    'recupera o valor dos parametros (se houverem)
    For iParam = 1 To iNumParam
    
        If aiTipoParam(iParam) = TIPO_TEXTO Then
        
            sValor = String(255, 0)
    
            lErro = Devolve_Valor_Param_String(sValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 92274

            avParam(iParam) = StringZ(sValor)
        
        Else

            lErro = Devolve_Valor_Param_Double(dValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 92275
            
            avParam(iParam) = dValor
            
        End If
        
    Next
    
    lErro = Procura_Colecao_Mnemonico(sMnemonico, iNumParam, avParam(), objMnemonicoValor, objExeExp.colMnemonicoValor)
    If lErro <> SUCESSO Then gError 92276
    
    'se o mnemonico ainda não foi calculado ==> calcula-o
    If objMnemonicoValor Is Nothing Then
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        
        Set objMnemonicoValor.colValor = New Collection
        objMnemonicoValor.sMnemonico = sMnemonico
        objMnemonicoValor.iParam = iNumParam
        
        For iParam = 1 To iNumParam
        
            objMnemonicoValor.vParam(iParam) = avParam(iParam)
            
        Next
    
        If sMnemonico Like "L#" Or sMnemonico Like "L##" Then
            objMnemonicoValor.colValor.Add 0
        Else
    
            iAchou = 0
    
            For Each objMnemonicoFPreco In objExeExp.colMnemonico
                If objMnemonicoFPreco.sMnemonico = sMnemonico Then
                    If objMnemonicoFPreco.iFuncao = MNEMONICOFPRECO_NAO_E_FUNCAO Then
                        objMnemonicoValor.colValor.Add CDbl(objMnemonicoFPreco.sExpressao)
                        iAchou = 1
                        Exit For
                    ElseIf objMnemonicoFPreco.iFuncao = MNEMONICOFPRECO_E_FUNCAO Then
                        iAchou = 2
                        Exit For
                    End If
                End If
            Next
    
            If iAchou = 2 Then
    
                sProduto = objExeExp.sProduto
    
                lErro = CF("Calcula_MnemonicoFPreco2", objMnemonicoValor, sProduto, objExeExp)
                If lErro <> SUCESSO And lErro <> 92413 Then gError 92414
            
            End If
            
            If iAchou = 0 Or lErro = 92413 Then gError 92415
            
            
    '        If lErro = CONTABIL_MNEMONICO_NAO_ENCONTRADO Then
    '
    '            lErro = objContabil.Contabil_Calcula_Mnemonico(objMnemonicoValor)
    '            If lErro <> SUCESSO Then Error 36711
    '
    '        End If
            
        End If
        
        'armazena os valores do mnemonico
        objExeExp.colMnemonicoValor.Add objMnemonicoValor
        
    End If
    
    vValor = objMnemonicoValor.colValor.Item(iIndice)

    If VarType(vValor) = vbString Then
    
        lErro = Envia_Valor_Mnemonico_String(vValor, lpRel)
        If lErro <> SUCESSO Then gError 92277
        
    ElseIf VarType(vValor) = vbDate Then
            
        vValor = vValor + 693594
            
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 92278
    
    Else
    
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 92279
    
    End If

    MaqExp_Devolve_Valor_Mnemon_FPreco = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Devolve_Valor_Mnemon_FPreco:

    MaqExp_Devolve_Valor_Mnemon_FPreco = gErr

    Select Case gErr
    
        Case 92273, 92274, 92275, 92276, 92277, 92278, 92279, 92414
    
        Case 92415
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOFPRECO_NAO_ENCONTRADO", gErr, sMnemonico)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161620)
    
    End Select
    
    Exit Function
    
End Function

Function MaqExp_Armazena_Campo_FPreco(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iExp As Integer, ByVal iCampo As Integer, ByVal dValor As Double) As Long
'armazena em colLancamentos (item de objExeExp) o valor que vai ser retornado pela chamada da funcao Devolve_Valor_Formula. iExp indica se é o primeiro valor da linha que está sendo retornado (= 0) ou não.

Dim lErro As Long
Dim sValor As String
Dim objLancamentos As ClassLancamentos
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_MaqExp_Armazena_Campo_FPreco

    objExeExp.vValor = dValor
    
    MaqExp_Armazena_Campo_FPreco = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Armazena_Campo_FPreco:

    MaqExp_Armazena_Campo_FPreco = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161621)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Devolve_Valor_MnemonicoComiss(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iIndice As Integer, ByVal iNumParam As Integer, ByVal iTipoParam1 As Integer, ByVal iTipoParam2 As Integer, ByVal iTipoParam3 As Integer) As Long
'devolve o valor do mnemonico passado como parametro

Dim lErro As Long
Dim sValor As String
Dim sMnemonico As String
Dim iParam As Integer
Dim aiTipoParam(1 To 3) As Integer
Dim avParam(1 To 3) As Variant
Dim dValor As Double
Dim objMnemonicoValor As ClassMnemonicoValor
Dim vValor As Variant

On Error GoTo Erro_MaqExp_Devolve_Valor_MnemonicoComiss

    'Inicializa a string
    sMnemonico = String(255, 0)
    
    'recupera o nome do mnemonico
    lErro = Devolve_Identificador_Mnemonico(sMnemonico, lpRel)
    If lErro <> SUCESSO Then gError 94961

    sMnemonico = StringZ(sMnemonico)

    aiTipoParam(1) = iTipoParam1
    aiTipoParam(2) = iTipoParam2
    aiTipoParam(3) = iTipoParam3
    
    'recupera o valor dos parametros (se houverem)
    For iParam = 1 To iNumParam
    
        If aiTipoParam(iParam) = TIPO_TEXTO Then
        
            sValor = String(255, 0)
    
            lErro = Devolve_Valor_Param_String(sValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 94962

            avParam(iParam) = StringZ(sValor)
        
        
        Else

            lErro = Devolve_Valor_Param_Double(dValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 94963
            
            avParam(iParam) = dValor
            
        End If
        
    Next
    
    lErro = Procura_Colecao_Mnemonico(sMnemonico, iNumParam, avParam(), objMnemonicoValor, objExeExp.colMnemonicoValor)
    If lErro <> SUCESSO Then gError 94964
    
    'se o mnemonico ainda não foi calculado ==> calcula-o
    If objMnemonicoValor Is Nothing Then
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        
        Set objMnemonicoValor.colValor = New Collection
        objMnemonicoValor.sMnemonico = sMnemonico
        objMnemonicoValor.iParam = iNumParam
        
        For iParam = 1 To iNumParam
        
            objMnemonicoValor.vParam(iParam) = avParam(iParam)
            
        Next
    
        'Chama a função que calcula mnemônicos customizados
        lErro = objExeExp.objTransacao.objmnemonicoComissCalcAux.Calcula_Mnemonico_Comissoes(objMnemonicoValor)
        If lErro <> SUCESSO And lErro <> MNEMONICOCOMISSOES_NAO_ENCONTRADO Then gError 94965

        'Se não encontrou o mnemônico entre os mnemônicos customizados
        If lErro = MNEMONICOCOMISSOES_NAO_ENCONTRADO Then
        
            lErro = objExeExp.objTransacao.objMnemonicoComissCalc.Calcula_Mnemonico_Comissoes(objMnemonicoValor)
            If lErro <> SUCESSO And lErro <> MNEMONICOCOMISSOES_NAO_ENCONTRADO Then gError 94966
            
            'Se não encontrou o mnemônico => erro
            If lErro = MNEMONICOCOMISSOES_NAO_ENCONTRADO Then gError 102031
        
        End If
        
        'armazena os valores do mnemonico
        objExeExp.colMnemonicoValor.Add objMnemonicoValor
        
    End If
    
    'se o indice do elemento procurado for maior do que o numero de elementos
    If iIndice > objMnemonicoValor.colValor.Count Then gError 94967
    
    vValor = objMnemonicoValor.colValor.Item(iIndice)

    If VarType(vValor) = vbString Then
    
        lErro = Envia_Valor_Mnemonico_String(vValor, lpRel)
        If lErro <> SUCESSO Then gError 94968
        
    ElseIf VarType(vValor) = vbDate Then
            
        vValor = vValor + 693594
            
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 94969
    
    Else
    
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 94970
    
    End If

    MaqExp_Devolve_Valor_MnemonicoComiss = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Devolve_Valor_MnemonicoComiss:

    MaqExp_Devolve_Valor_MnemonicoComiss = gErr

    Select Case gErr
    
        Case 94961 To 94966, 94968 To 94970
    
        Case 102031
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOCOMISSOES_NAO_TRATADO", gErr, objMnemonicoValor.sMnemonico)
            
        Case 94967
            Call Rotina_Erro(vbOKOnly, "ERRO_INDICE_MAIOR_LINHAS_GRID", gErr, sMnemonico, iIndice, objMnemonicoValor.colValor.Count)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161622)
    
    End Select
    
    Exit Function
    
End Function

'Função criada por Mauricio Maciel em 08/04/2003
Function MaqExp_Testa_Mnemonico_FPlanilha(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal lNomeGrid As Long, iTipo As Integer, iNumParam As Integer, iParam1 As Integer, iParam2 As Integer, iParam3 As Integer, ByVal iInicio_Expressao As Integer) As Long

Dim lErro As Long
Dim sIdentificador As String
Dim objMnemonico As ClassMnemonicoFPTipo
Dim iIndice As Integer

On Error GoTo Erro_MaqExp_Testa_Mnemonico_FPlanilha


    sIdentificador = String(255, 0)
    
    lErro = Devolve_Identificador_Mnemonico(sIdentificador, lpRel)
    If lErro <> SUCESSO Then gError 92272

    sIdentificador = StringZ(sIdentificador)
    
    If sIdentificador Like "L#" Or sIdentificador Like "L##" Then
        If objExeExp.iLinhaAtual <= CInt(Mid(sIdentificador, 2)) Then gError 92271
        
        iTipo = TIPO_NUMERICO
        iNumParam = 0
    
    Else
    
        Set objMnemonico = Nothing
        
        For iIndice = 1 To objExeExp.colMnemonico.Count
            If objExeExp.colMnemonico.Item(iIndice).sMnemonico = sIdentificador Then
                Set objMnemonico = objExeExp.colMnemonico.Item(iIndice)
                Exit For
            End If
        Next
        
        If objMnemonico Is Nothing Then gError 92273
        
        iTipo = objMnemonico.iTipo
        iNumParam = objMnemonico.iNumParam
        iParam1 = objMnemonico.iParam1
        iParam2 = objMnemonico.iParam2
        iParam3 = objMnemonico.iParam3
    
    End If
    
    MaqExp_Testa_Mnemonico_FPlanilha = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Testa_Mnemonico_FPlanilha:

    MaqExp_Testa_Mnemonico_FPlanilha = gErr

    Select Case gErr
    
        Case 92271, 92272, 92273
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161623)
    
    End Select
    
    Exit Function

End Function

'Função criada por Mauricio Maciel em 08/04/2003
Function MaqExp_Devolve_Valor_Mnemon_FPlanilha(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iIndice As Integer, ByVal iNumParam As Integer, ByVal iTipoParam1 As Integer, ByVal iTipoParam2 As Integer, ByVal iTipoParam3 As Integer) As Long
'devolve o valor do mnemonico passado como parametro

Dim lErro As Long
Dim sValor As String
Dim sMnemonico As String
Dim iParam As Integer
Dim aiTipoParam(1 To 3) As Integer
Dim avParam(1 To 3) As Variant
Dim dValor As Double
Dim objMnemonicoValor As ClassMnemonicoValor
Dim vValor As Variant
Dim objContabil As New ClassContabil
Dim objMnemonicoFPTipo As Object
Dim iAchou As Integer
Dim sProduto As String

On Error GoTo Erro_MaqExp_Devolve_Valor_Mnemon_FPlanilha

    sMnemonico = String(255, 0)
    
    'recupera o nome do mnemonico
    lErro = Devolve_Identificador_Mnemonico(sMnemonico, lpRel)
    If lErro <> SUCESSO Then gError 92273

    sMnemonico = StringZ(sMnemonico)

    aiTipoParam(1) = iTipoParam1
    aiTipoParam(2) = iTipoParam2
    aiTipoParam(3) = iTipoParam3
    
    'recupera o valor dos parametros (se houverem)
    For iParam = 1 To iNumParam
    
        If aiTipoParam(iParam) = TIPO_TEXTO Then
        
            sValor = String(255, 0)
    
            lErro = Devolve_Valor_Param_String(sValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 92274

            avParam(iParam) = StringZ(sValor)
        
        Else

            lErro = Devolve_Valor_Param_Double(dValor, lpRel, iParam)
            If lErro <> SUCESSO Then gError 92275
            
            avParam(iParam) = dValor
            
        End If
        
    Next
    
    lErro = Procura_Colecao_Mnemonico(sMnemonico, iNumParam, avParam(), objMnemonicoValor, objExeExp.colMnemonicoValor)
    If lErro <> SUCESSO Then gError 92276
    
    'se o mnemonico ainda não foi calculado ==> calcula-o
    If objMnemonicoValor Is Nothing Then
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        
        Set objMnemonicoValor.colValor = New Collection
        objMnemonicoValor.sMnemonico = sMnemonico
        objMnemonicoValor.iParam = iNumParam
        
        For iParam = 1 To iNumParam
        
            objMnemonicoValor.vParam(iParam) = avParam(iParam)
            
        Next
    
        If sMnemonico Like "L#" Or sMnemonico Like "L##" Then
            objMnemonicoValor.colValor.Add 0
        Else
    
            iAchou = 0
    
            For Each objMnemonicoFPTipo In objExeExp.colMnemonico
                If objMnemonicoFPTipo.sMnemonico = sMnemonico Then
                    If objMnemonicoFPTipo.iFuncao = MNEMONICOFPRECO_NAO_E_FUNCAO Then
                        objMnemonicoValor.colValor.Add StrParaDbl(objMnemonicoFPTipo.sExpressao)
                        iAchou = 1
                        Exit For
                    ElseIf objMnemonicoFPTipo.iFuncao = MNEMONICOFPRECO_E_FUNCAO Then
                        iAchou = 2
                        Exit For
                    End If
                End If
            Next
    
            If iAchou = 2 Then
    
                sProduto = objExeExp.sProduto
    
                lErro = CF("Calcula_MnemonicoFPreco2", objMnemonicoValor, sProduto, objExeExp)
                If lErro <> SUCESSO And lErro <> 92413 Then gError 92414
            
            End If
            
            If iAchou = 0 Or lErro = 92413 Then gError 92415
            
            
    '        If lErro = CONTABIL_MNEMONICO_NAO_ENCONTRADO Then
    '
    '            lErro = objContabil.Contabil_Calcula_Mnemonico(objMnemonicoValor)
    '            If lErro <> SUCESSO Then Error 36711
    '
    '        End If
            
        End If
        
        'armazena os valores do mnemonico
        objExeExp.colMnemonicoValor.Add objMnemonicoValor
        
    End If
    
    vValor = objMnemonicoValor.colValor.Item(iIndice)

    If VarType(vValor) = vbString Then
    
        lErro = Envia_Valor_Mnemonico_String(vValor, lpRel)
        If lErro <> SUCESSO Then gError 92277
        
    ElseIf VarType(vValor) = vbDate Then
            
        vValor = vValor + 693594
            
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 92278
    
    Else
    
        lErro = Envia_Valor_Mnemonico_Double(vValor, lpRel)
        If lErro <> SUCESSO Then gError 92279
    
    End If

    MaqExp_Devolve_Valor_Mnemon_FPlanilha = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Devolve_Valor_Mnemon_FPlanilha:

    MaqExp_Devolve_Valor_Mnemon_FPlanilha = gErr

    Select Case gErr
    
        Case 92273, 92274, 92275, 92276, 92277, 92278, 92279, 92414
    
        Case 92415
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOFPRECO_NAO_ENCONTRADO", gErr, sMnemonico)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161624)
    
    End Select
    
    Exit Function
    
End Function

'*************************************************************************************************************************************************************************************************************************************
' Fim das Funções e Declarações usados na maquina de expressões
'*************************************************************************************************************************************************************************************************************************************

Function MaqExp_Armazena_CampoWFW(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iExp As Integer, ByVal iCampo As Integer, ByVal dValor As Double) As Long
'armazena em colLancamentos (item de objExeExp) o valor que vai ser retornado pela chamada da funcao Devolve_Valor_Formula. iExp indica se é o primeiro valor da linha que está sendo retornado (= 0) ou não.

Dim lErro As Long
Dim sValor As String
Dim objRegraWFW As ClassRegraWFW

On Error GoTo Erro_MaqExp_Armazena_CampoWFW

    ' Se é o primeiro campo de uma nova linha ==> tem que criar uma nova instancia de classLancamentos.
    If iExp = 0 Then
    
        Set objRegraWFW = New ClassRegraWFW
        objExeExp.colRegras.Add objRegraWFW
        
    Else
    
        Set objRegraWFW = objExeExp.colRegras.Item(objExeExp.colRegras.Count)
        
    End If

    'se o campo retorna uma string ==> chama a função que pega o valor do campo.
    'se o campo retorna um double este é passado como parametro na chamada da funcao (dValor)
    If iCampo = CAMPO_RELSEL Or iCampo = CAMPO_RELANEXO Or iCampo = CAMPO_EMAILPARA Or iCampo = CAMPO_EMAILASSUNTO Or iCampo = CAMPO_EMAILMSG Or iCampo = CAMPO_AVISOMSG Or iCampo = CAMPO_LOGDOC Or iCampo = CAMPO_LOGMSG Then

        sValor = String(255, 0)
    
        lErro = Devolve_Valor_Campo_String(sValor, lpRel)
        If lErro <> SUCESSO Then gError 178147

        sValor = StringZ(sValor)
        
    End If
    
    Select Case iCampo
    
        Case CAMPO_REGRA
            objRegraWFW.dRegraRet = dValor
            
        Case CAMPO_EMAILPARA
            objRegraWFW.sEmailParaRet = sValor
            
        Case CAMPO_EMAILASSUNTO
            objRegraWFW.sEmailAssuntoRet = sValor
            
        Case CAMPO_EMAILMSG
            objRegraWFW.sEmailMsgRet = sValor
            
        Case CAMPO_AVISOMSG
            objRegraWFW.sAvisoMsgRet = sValor
            
        Case CAMPO_LOGDOC
            objRegraWFW.sLogDocRet = sValor
            
        Case CAMPO_LOGMSG
            objRegraWFW.sLogMsgRet = sValor
            
        Case CAMPO_RELSEL
            objRegraWFW.sRelSelRet = sValor
        
        Case CAMPO_RELANEXO
            objRegraWFW.sRelAnexoRet = sValor
            
    End Select
    
    MaqExp_Armazena_CampoWFW = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Armazena_CampoWFW:

    MaqExp_Armazena_CampoWFW = gErr

    Select Case gErr
    
        Case 178147
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178148)
    
    End Select
    
    Exit Function

End Function

Function MaqExp_Armazena_CampoMsg(ByVal objExeExp As ClassExeExp, ByVal lpRel As Long, ByVal iExp As Integer, ByVal iCampo As Integer, ByVal dValor As Double) As Long
'armazena em colLancamentos (item de objExeExp) o valor que vai ser retornado pela chamada da funcao Devolve_Valor_Formula. iExp indica se é o primeiro valor da linha que está sendo retornado (= 0) ou não.

Dim lErro As Long
Dim sValor As String
Dim objRegra As ClassRegrasMsg

On Error GoTo Erro_MaqExp_Armazena_CampoMsg

    ' Se é o primeiro campo de uma nova linha ==> tem que criar uma nova instancia de classLancamentos.
    If iExp = 0 Then
        Set objRegra = New ClassRegrasMsg
        objExeExp.colRegras.Add objRegra
    Else
        Set objRegra = objExeExp.colRegras.Item(objExeExp.colRegras.Count)
    End If

    'se o campo retorna uma string ==> chama a função que pega o valor do campo.
    'se o campo retorna um double este é passado como parametro na chamada da funcao (dValor)
    If iCampo <> CAMPO_REGRA Then

        sValor = String(255, 0)
    
        lErro = Devolve_Valor_Campo_String(sValor, lpRel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        sValor = StringZ(sValor)
        
    End If
    
    Select Case iCampo
    
        Case CAMPO_REGRA
            objRegra.dRegraRet = dValor
            
        Case Else
            objRegra.sMensagemRet = sValor
            
    End Select
    
    MaqExp_Armazena_CampoMsg = SUCESSO
    
    Exit Function
    
Erro_MaqExp_Armazena_CampoMsg:

    MaqExp_Armazena_CampoMsg = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178148)
    
    End Select
    
    Exit Function

End Function

