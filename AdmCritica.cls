VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmCritica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ConsisteInscricaoEstadual Lib "DllInscE32" (ByVal sIE As String, ByVal sUF As String) As Integer

'Function E_Zero(dNumero As Double) As Boolean
''Retorna True se dNumero for zero quando arredondado até precisão de FORMATO_ESTOQUE
'
'Dim lPrecisao As Long
'
'    'Número de casas decimais de FORMATO_ESTOQUE
'    lPrecisao = Len(Mid(FORMATO_ESTOQUE, InStr(FORMATO_ESTOQUE, "."))) - 1
'
'    'Arredonda
'    dNumero = Round(dNumero, lPrecisao)
'
'    'Testa se deu zero
'    If dNumero = 0 Then
'        E_Zero = True
'    Else
'        E_Zero = False
'    End If
'
'End Function
'
'Function IniciaLetra(sString As String) As Boolean
''Testa se sString começa por letra. Retorna True se começar e False caso contrário.
'
'Dim lErro As Long
'
'On Error GoTo Erro_IniciaLetra
'
'    If Len(sString) = 0 Then Error 25002
'
'    'Testa se começa por letra
'    If Asc(sString) >= 65 And Asc(sString) <= 90 Or Asc(sString) >= 97 And Asc(sString) <= 122 Then
'        IniciaLetra = True
'    Else
'        IniciaLetra = False
'    End If
'
'    Exit Function
'
'Erro_IniciaLetra:
'
'    IniciaLetra = False
'
'    Select Case Err
'
'        Case 25002  'Se String for vazia não começa por letra
'
'        Case Else
'
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142282)
'
'    End Select
'
'End Function
'Function Data_Critica(ByVal sData As String) As Long
'
'Dim lErro As Long
'Dim lPosition1 As Long
'Dim lPosition2 As Long
'Dim iDia As Integer
'
'On Error GoTo Erro_Data_Critica
'
'    If IsDate(sData) Then
'        'Posição do 1º separador "/" dentro de sData
'        lPosition1 = InStr(1, sData, "/")
'        'Posição do 2º separador "/" dentro de sData
'        lPosition2 = InStr(lPosition1 + 1, sData, "/")
'        'Não tem 2º separador então não é data
'        If lPosition2 = 0 Then Error 25026
'        'Isola os caracteres que antecedem a primeira barra "/"
'        iDia = CInt(Left(sData, lPosition1 - 1))
'        'Testa se iDia coincide com o dia da Data
'        If Not iDia = Day(sData) Then Error 25027
'    ElseIf sData <> "  /  /  " Then
'        Error 6001
'    End If
'
'    Data_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Data_Critica:
'
'    Data_Critica = Err
'
'    Select Case Err
'
'        Case 6001, 25026, 25027
'
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_DATA", Err)
'
'        Case Else
'
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142283)
'
'    End Select
'
'End Function
'Function Cpf_Critica(ByVal sCpf As String) As Long
'
'Dim lErro As Long
'Dim dCpf As Double
'Dim iDig(1 To STRING_CPF) As Integer
'Dim iDigVer(1 To 2) As Integer
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim iSum As Integer
'Dim dCpfMaximo As Double
'
'On Error GoTo Erro_Cpf_Critica
'
'    sCpf = Trim(sCpf)
'
'    If Len(sCpf) > 0 Then
'
'        If Not IsNumeric(sCpf) Then Error 6624
'
'        dCpf = CDbl(sCpf)
'
'        If dCpf <= 0 Then Error 6626
'
'        dCpfMaximo = CDbl(String(STRING_CPF, "9"))
'
'        If dCpf > dCpfMaximo Then Error 6896
'
'        If dCpf <> Int(dCpf) Then Error 6630
'
'        sCpf = CStr(dCpf)
'
'        If Len(sCpf) < STRING_CPF Then
'            sCpf = String(STRING_CPF - Len(sCpf), "0") & sCpf
'        End If
'
'        For iIndice = 1 To STRING_CPF
'            iDig(iIndice) = CInt(Mid(sCpf, iIndice, 1))
'        Next
'
'        For iIndice2 = 1 To 2
'
'            iSum = 0
'
'            For iIndice = 1 To STRING_CPF - 2
'                iSum = iSum + (10 + iIndice2 - iIndice) * iDig(iIndice)
'            Next
'
'            iSum = IIf(iIndice2 = 1, iSum, iSum + 2 * iDigVer(1))
'
'            iDigVer(iIndice2) = ((iSum * 10) Mod 11) Mod 10
'
'        Next
'
'
'        If iDigVer(1) <> iDig(STRING_CPF - 1) Or iDigVer(2) <> iDig(STRING_CPF) Then Error 6633
'
'    End If
'
'    Cpf_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Cpf_Critica:
'
'    Cpf_Critica = Err
'
'    Select Case Err
'
'        Case 6624
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_NAO_NUMERICO", Err, sCpf)
'
'        Case 6626
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_MENOR_OU_IGUAL_ZERO", Err, sCpf)
'
'        Case 6630
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_NAO_INTEIRO", Err, sCpf)
'
'        Case 6633
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_INVALIDO", Err, sCpf)
'
'        Case 6896
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_MAIOR_QUE_CPFMAXIMO", Err, sCpf)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142284)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Cgc_Critica(ByVal sCgc As String) As Long
'
'Dim lErro As Long
'Dim dCgc As Double
'Dim iDig(1 To STRING_CGC) As Integer
'Dim iDigVer(1 To 3) As Integer
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim iSum As Integer
'Dim iParcela As Integer
'Dim dCgcMaximo As Double
'
'On Error GoTo Erro_Cgc_Critica
'
'    sCgc = Trim(sCgc)
'
'    If Len(sCgc) > 0 Then
'
'        If Not IsNumeric(sCgc) Then Error 6641
'
'        dCgc = CDbl(sCgc)
'
'        If dCgc <= 0 Then Error 6642
'
'        dCgcMaximo = CDbl(String(STRING_CGC, "9"))
'
'        If dCgc > dCgcMaximo Then Error 6897
'
'        If dCgc <> Int(dCgc) Then Error 6643
'
'        sCgc = CStr(dCgc)
'
'        If Len(sCgc) < STRING_CGC Then
'            sCgc = String(STRING_CGC - Len(sCgc), "0") & sCgc
'        End If
'
'        For iIndice = 1 To STRING_CGC
'            iDig(iIndice) = CInt(Mid(sCgc, iIndice, 1))
'        Next
'
'        For iIndice2 = 1 To 2
'
'            iSum = 0
'
'            For iIndice = 1 To 3 + iIndice2
'                iSum = iSum + (5 + iIndice2 - iIndice) * iDig(iIndice)
'            Next
'
'            For iIndice = 4 + iIndice2 To STRING_CGC - 2
'                iSum = iSum + (13 + iIndice2 - iIndice) * iDig(iIndice)
'            Next
'
'            iSum = IIf(iIndice2 = 1, iSum, iSum + 2 * iDigVer(1))
'
'            iDigVer(iIndice2) = ((iSum * 10) Mod 11) Mod 10
'
'        Next
'
'        iSum = 0
'
'        For iIndice = 1 To 7
'
'            iParcela = iDig(iIndice) * ((iIndice Mod 2) + 1)
'            iSum = iSum + (iParcela Mod 10) + Int(iParcela / 10)
'
'        Next
'
''       este o 3º DV foi eliminado em 94/95
''       iDigVer(3) = IIf(iSum / 10 = Int(iSum / 10), iSum / 10, Int(iSum / 10) + 1) * 10 - iSum
'
''       If iDigVer(1) <> iDig(STRING_CGC - 1) Or iDigVer(2) <> iDig(STRING_CGC) Or iDigVer(3) <> iDig(8) Then Error 6644
'
'        If iDigVer(1) <> iDig(STRING_CGC - 1) Or iDigVer(2) <> iDig(STRING_CGC) Then Error 6644
'
'    End If
'
'    Cgc_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Cgc_Critica:
'
'    Cgc_Critica = Err
'
'    Select Case Err
'
'        Case 6641
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_NUMERICO", Err, sCgc)
'
'        Case 6642
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_MENOR_OU_IGUAL_ZERO", Err, sCgc)
'
'        Case 6643
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INTEIRO", Err, sCgc)
'
'        Case 6644
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_INVALIDO", Err, sCgc)
'
'       Case 6897
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_MAIOR_QUE_CGCMAXIMO", Err, sCgc)
'
'       Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142285)
'
'    End Select
'
'    Exit Function
'
'End Function
'Function Inteiro_Critica(sNumero As String) As Long
''Critica se é Tipo inteiro positivo (entre 1 e 9999)
'
'Dim lErro As Long
'Dim dNumero As Double
'
'On Error GoTo Error_Inteiro_Critica
'
'    If IsNumeric(sNumero) Then
'        dNumero = 0#
'        dNumero = CDbl(sNumero)
'    Else
'        Error 12310
'    End If
'
'    If dNumero <> Int(dNumero) Then Error 12358
'
'    If dNumero < 1 Then Error 6748
'
'    If dNumero > 9999 Then Error 6749
'
'    Inteiro_Critica = SUCESSO
'
'Exit Function
'
'Error_Inteiro_Critica:
'
'    Inteiro_Critica = Err
'
'    Select Case Err
'
'        Case 12310
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sNumero)
'
'        Case 12358
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sNumero)
'
'        Case 6748
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
'
'        Case 6749
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sNumero)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142286)
'
'    End Select
'
'    Exit Function
'
'End Function
'Function Long_Critica(sNumero As String) As Long
''Critica se é tipo Long positivo (entre 1 e 999 999 999)
'
'Dim lErro As Long
'Dim dNumero As Double
'
'On Error GoTo Error_Long_Critica
'
'    If IsNumeric(sNumero) Then
'        dNumero = 0#
'        dNumero = CDbl(sNumero)
'    Else
'        Error 12311
'    End If
'
'    If dNumero <> Int(dNumero) Then Error 12359
'
'    If dNumero < 1 Then Error 6750
'
'    If dNumero > 999999999 Then Error 6751
'
'    Long_Critica = SUCESSO
'
'Exit Function
'
'Error_Long_Critica:
'
'    Long_Critica = Err
'
'    Select Case Err
'
'        Case 12311
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sNumero)
'
'        Case 12359
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sNumero)
'
'        Case 6750
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
'
'        Case 6751
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LONG_OVERFLOW", Err, sNumero)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142287)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Porcentagem_Critica(sString As String) As Long
''Critica o valor passado em sString , se nao for entre 0 e 100 retorna erro.
'
'Dim dValor As Double
'Dim lErro As Long
'
'On Error GoTo Error_Porcentagem_Critica
'
'    If IsNumeric(sString) Then
'        dValor = 0#
'        dValor = CDbl(sString)
'    Else
'        Error 12530
'    End If
'
'    If dValor > 100 Or dValor < 0 Then Error 12314
'
'    Porcentagem_Critica = SUCESSO
'
'    Exit Function
'
'Error_Porcentagem_Critica:
'
'    Porcentagem_Critica = Err
'
'    Select Case Err
'
'        Case 12530
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
'
'        Case 12314
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM", Err, dValor)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142288)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Valor_Critica(sValor As String) As Long
''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO
'
'Dim curTeste As Currency
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Critica
'
'    curTeste = CCur(sValor)
'
'    Valor_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Critica:
'
'    Valor_Critica = Err
'
'   lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
'
'    Exit Function
'
'End Function
'
'Function Valor_Inteiro_Critica(sValor As String) As Long
''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
'
'Dim iTeste As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Inteiro_Critica
'
'    iTeste = CInt(sValor)
'
'    Valor_Inteiro_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Inteiro_Critica:
'
'    Valor_Inteiro_Critica = Err
'
'    Exit Function
'
'End Function
'
'Function Valor_Positivo_Critica(sValor As String) As Long
''critica se o valor passado como parametro é valido e positivo. Se estiver tudo ok retorna SUCESSO
'
'Dim curTeste As Currency
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Positivo_Critica
'
'    curTeste = CCur(sValor)
'
'    If curTeste <= 0 Then Error 10143
'
'    Valor_Positivo_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Positivo_Critica:
'
'    Valor_Positivo_Critica = Err
'
'    Select Case Err
'
'        Case 10143
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_POSITIVO", Err, sValor)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Valor_Positivo_Critica_Double(sValor As String) As Long
''verifica se um valor double é positivo
'
'Dim dTeste As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Positivo_Critica_Double
'
'    dTeste = CDbl(sValor)
'
'    If dTeste <= 0 Then Error 60769
'
'    Valor_Positivo_Critica_Double = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Positivo_Critica_Double:
'
'    Valor_Positivo_Critica_Double = Err
'
'    Select Case Err
'
'        Case 60769
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_POSITIVO", Err, sValor)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
'
'    End Select
'
'    Exit Function
'
'    Exit Function
'
'End Function
'
'Function Valor_NaoNegativo_Critica(sValor As String) As Long
''critica se o valor passado como parametro é valido e não é negativo. Se estiver tudo ok retorna SUCESSO
'
'Dim curTeste As Currency
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_NaoNegativo_Critica
'
'    curTeste = CCur(sValor)
'
'    If curTeste < 0 Then Error 6752
'
'    Valor_NaoNegativo_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_NaoNegativo_Critica:
'
'    Valor_NaoNegativo_Critica = Err
'
'    Select Case Err
'
'        Case 6752
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, sValor)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Produto_Critica(ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer, Optional ByVal bAceitaInativo As Boolean = False) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido

Dim lErro As Long
Dim sProdutoFormatado As String

On Error GoTo Erro_Produto_Critica

    If Len(sProduto) > 0 Then
    
        'Formata o código do Produto como no BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 25039
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            'Carrega sProdutoFormatado em objProduto
            objProduto.sCodigo = sProdutoFormatado
            
            'Tenta ler o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 25040
    
            'Ausência de Produto no BD
            If lErro = 28030 Then Error 25041
            
            'Critica se é produto ATIVO
            If objProduto.iAtivo <> Ativo And Not bAceitaInativo Then Error 25042
            
            'Critica se é produto NAO_GERENCIAL
            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 25043
        
        End If
    
    End If
    
    Produto_Critica = SUCESSO
    
    Exit Function

Erro_Produto_Critica:

    Produto_Critica = Err
    
    Select Case Err
    
        Case 25039, 25040  'Erro criticado na rotina chamada
        
        Case 25041  'Produto não está cadastrado
        
        Case 25042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
        
        Case 25043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142289)
        
    End Select
    
    Exit Function

End Function

Function Produto_Critica1(ByVal lComando As Long, ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'Devolve os dados do Produto em objProduto
'O produto deve ser fornecido no formato do BD

Dim lErro As Long
Dim sProdutoFormatado As String

On Error GoTo Erro_Produto_Critica1


    If Len(sProduto) > 0 Then
    
'        'Formata o código do Produto como no BD
'        lErro = CF("Produto_Formata",sProduto, sProdutoFormatado, iProdutoPreenchido)
'        If lErro Then Error 20581
'
'        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            'Carrega sProdutoFormatado em objProduto
            objProduto.sCodigo = sProduto
            
            'Tenta ler o Produto
            lErro = CF("Produto_Le2", lComando, objProduto)
            If lErro <> SUCESSO And lErro <> 20579 Then Error 20580
    
            'Ausência de Produto no BD
            If lErro = 20579 Then Error 20582
            
'            'Critica se é produto ATIVO
'            If objProduto.iAtivo <> Ativo Then Error 20583
            
            'Critica se é produto NAO_GERENCIAL
            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 20584
        
'        End If
    
    End If
    
    Produto_Critica1 = SUCESSO
    
    Exit Function

Erro_Produto_Critica1:

    Produto_Critica1 = Err
    
    Select Case Err
    
        Case 20580, 20581  'Erro criticado na rotina chamada
        
        Case 20582  'Produto não está cadastrado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, sProduto)
        
        Case 20583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
        
        Case 20584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142290)
        
    End Select
    
    Exit Function

End Function
Function Produto_Critica_Filial(ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
'Verifica em Produto está relacionado com a FilialEmpresa.

Dim lErro As Long
Dim sProdutoFormatado As String
Dim lComando As Long
Dim iFilialEmpresa As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Produto_Critica_Filial

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51380

    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then Error 51385
    If lErro <> SUCESSO Then Error 51381
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM ProdutosFilial WHERE FilialEmpresa =? AND Produto = ?", iFilialEmpresa, giFilialEmpresa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 51382
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51383
    If lErro = AD_SQL_SEM_DADOS Then Error 51384
           
    Call Comando_Fechar(lComando)
    
    Produto_Critica_Filial = SUCESSO
    
    Exit Function

Erro_Produto_Critica_Filial:

    Produto_Critica_Filial = Err
    
    Select Case Err
    
        Case 51380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 51381
        
        Case 51382, 51383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, giFilialEmpresa, sProduto)
        
        Case 51384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", Err, sProduto, giFilialEmpresa)
        
        Case 51385
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142291)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Function Long_Critica1(sNumero As String) As Long
''Critica se é tipo Long positivo (entre 1 e 99 999)
'
'Dim lErro As Long
'Dim dNumero As Double
'
'On Error GoTo Error_Long_Critica1
'
'    If IsNumeric(sNumero) Then
'        dNumero = 0#
'        dNumero = CDbl(sNumero)
'    Else
'        Error 44027
'    End If
'
'    If dNumero <> Int(dNumero) Then Error 44028
'
'    If dNumero < 1 Then Error 44029
'
'    If dNumero > 99999 Then Error 44030
'
'    Long_Critica1 = SUCESSO
'
'Exit Function
'
'Error_Long_Critica1:
'
'    Long_Critica1 = Err
'
'    Select Case Err
'
'        Case 44027, 44028, 44029, 44030
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142292)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Valor_Long_Critica(sValor As String) As Long
''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
'
'Dim lTeste As Long
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Long_Critica
'
'    lTeste = CLng(sValor)
'
'    Valor_Long_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Long_Critica:
'
'    Valor_Long_Critica = Err
'
'    Exit Function
'
'End Function
'
'Function Valor_Date_Critica(sValor As String) As Long
''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
'
'Dim dtTeste As Date
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Date_Critica
'
'    dtTeste = CDate(sValor)
'
'    Valor_Date_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Date_Critica:
'
'    Valor_Date_Critica = Err
'
'    Exit Function
'
'End Function
'
'Function Valor_Double_Critica(sValor As String) As Long
''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
'
'Dim dTeste As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Valor_Double_Critica
'
'    dTeste = CDbl(sValor)
'
'    Valor_Double_Critica = SUCESSO
'
'    Exit Function
'
'Erro_Valor_Double_Critica:
'
'    Valor_Double_Critica = Err
'
'    Exit Function
'
'End Function
'
'Function Porcentagem_Critica_Nao_Zero(sString As String) As Long
''Critica o valor passado em sString , se nao for entre -100 e 100 ou se for zero retorna erro.
'
'Dim dValor As Double
'Dim lErro As Long
'
'On Error GoTo Error_Porcentagem_Critica_Nao_Zero
'
'    If IsNumeric(sString) Then
'        dValor = 0#
'        dValor = CDbl(sString)
'    Else
'        Error 58075
'    End If
'
'    If dValor < -99.99 Then Error 58076
'
'    If dValor = 0 Then Error 58077
'
'    Porcentagem_Critica_Nao_Zero = SUCESSO
'
'    Exit Function
'
'Error_Porcentagem_Critica_Nao_Zero:
'
'    Porcentagem_Critica_Nao_Zero = Err
'
'    Select Case Err
'
'        Case 58075
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
'
'        Case 58076
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM2", Err, dValor)
'
'        Case 58077
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ZERO", Err)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142293)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Porcentagem_Critica_Negativa(sString As String) As Long
''Critica o valor passado em sString aceitando valores negativos até -99.99
'
'Dim dValor As Double
'Dim lErro As Long
'
'On Error GoTo Error_Porcentagem_Critica_Negativa
'
'    If IsNumeric(sString) Then
'        dValor = 0#
'        dValor = CDbl(sString)
'    Else
'        Error 59063
'    End If
'
'    If dValor < -99.99 Then Error 59064
'
'    Porcentagem_Critica_Negativa = SUCESSO
'
'    Exit Function
'
'Error_Porcentagem_Critica_Negativa:
'
'    Porcentagem_Critica_Negativa = Err
'
'    Select Case Err
'
'        Case 59063
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
'
'        Case 59064
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM2", Err, dValor)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142294)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Porcentagem_Critica2(sString As String) As Long
''Critica o valor passado em sString , se nao for entre 0 e 99 retorna erro.
'
'Dim dValor As Double
'Dim lErro As Long
'
'On Error GoTo Error_Porcentagem_Critica2
'
'    If IsNumeric(sString) Then
'        dValor = 0#
'        dValor = CDbl(sString)
'    Else
'        Error 58454
'    End If
'
'    If dValor > 99 Or dValor < 0 Then Error 58455
'
'    Porcentagem_Critica2 = SUCESSO
'
'    Exit Function
'
'Error_Porcentagem_Critica2:
'
'    Porcentagem_Critica2 = Err
'
'    Select Case Err
'
'        Case 58454
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
'
'        Case 58455
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM3", Err, dValor)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142295)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'

Function Produto_Critica2(ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO
'Se ele não for GERENCIAL retorna erro sem msg.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido

Dim lErro As Long
Dim sProdutoFormatado As String

On Error GoTo Erro_Produto_Critica2

    If Len(sProduto) > 0 Then
    
        'Formata o código do Produto como no BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 25039
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            'Carrega sProdutoFormatado em objProduto
            objProduto.sCodigo = sProdutoFormatado
            
            'Tenta ler o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 25040
    
            'Ausência de Produto no BD
            If lErro = 28030 Then Error 25041
            
            'Critica se é produto ATIVO
            If objProduto.iAtivo <> Ativo Then Error 25042
            
            'Critica se é produto NAO_GERENCIAL
            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 25043
        
        End If
    
    End If
    
    Produto_Critica2 = SUCESSO
    
    Exit Function

Erro_Produto_Critica2:

    Produto_Critica2 = Err
    
    Select Case Err
    
        Case 25039, 25040  'Erro criticado na rotina chamada
        
        Case 25041  'Produto não está cadastrado
        
        Case 25042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
        
        Case 25043
'''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142296)
        
    End Select
    
    Exit Function

End Function


Function Produto_Critica_Filial2(ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO .
'Se ele não for GERENCIAL retorna erro (86295) sem msg.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
'Verifica em Produto está relacionado com a FilialEmpresa.

Dim lErro As Long
Dim sProdutoFormatado As String
Dim lComando As Long
Dim iFilialEmpresa As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Produto_Critica_Filial2

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 51380

    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 51385
    If lErro = 25043 Then gError 86295
    
    If lErro <> SUCESSO Then gError 51381
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM ProdutosFilial WHERE FilialEmpresa =? AND Produto = ?", iFilialEmpresa, giFilialEmpresa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 51382
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 51383
    If lErro = AD_SQL_SEM_DADOS Then gError 51384
           
    Call Comando_Fechar(lComando)
    
    Produto_Critica_Filial2 = SUCESSO
    
    Exit Function

Erro_Produto_Critica_Filial2:

    Produto_Critica_Filial2 = gErr
    
    Select Case gErr
    
        Case 51380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 51381, 51385, 86295
        
        Case 51382, 51383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, giFilialEmpresa, sProduto)
        
        Case 51384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", gErr, sProduto, giFilialEmpresa)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142297)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function InscricaoEstadual_Critica(ByVal sIE As String, ByVal sUF As String, Optional bValidaIsento As Boolean = True) As Long

Dim lErro As Long

On Error GoTo Erro_InscricaoEstadual_Critica

    sIE = Replace(sIE, "/", "")
    sIE = Replace(sIE, "\", "")
    sIE = Replace(sIE, ".", "")
    sIE = Replace(sIE, "-", "")

    If Len(Trim(sIE)) > 0 Then

        If Len(Trim(sUF)) = 0 Then gError 189573
    
        lErro = InscricaoEstadual_Critica2(sUF, sIE)
        
        If lErro = 189569 Then gError 189571
        If lErro = 189570 Then gError 189572
        
        If bValidaIsento Then
        
            If UCase(sIE) = "ISENTO" Then gError 189574
        
        End If
    
    End If
    
    InscricaoEstadual_Critica = SUCESSO
    
    Exit Function

Erro_InscricaoEstadual_Critica:

    InscricaoEstadual_Critica = gErr
    
    Select Case gErr

        Case 189571
            Call Rotina_Erro(vbOKOnly, "ERRO_INSCESTADUAL_INVALIDA", gErr, sUF)
        
        Case 189572
            Call Rotina_Erro(vbOKOnly, "ERRO_INSCESTADUAL_PARAM_INVALIDOS", gErr, sIE, sUF)

        Case 189573
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_PREECHIDA", gErr)

        Case 189574
            Call Rotina_Erro(vbOKOnly, "ERRO_INSCESTADUAL_INSENTO_NAO_PERMITIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189567)
        
    End Select
    
    Exit Function

End Function

Function InscricaoEstadual_Critica2(ByVal sUF As String, ByVal sIE As String) As Long

Dim lErro As Long
Dim lRetorno As Long

On Error GoTo Erro_InscricaoEstadual_Critica2

    lRetorno = ConsisteInscricaoEstadual(sIE, sUF)
    If lRetorno <> 0 Then
        If lRetorno = 1 Then
           gError 189569 'Call MsgBox("Inscrição inválida para " + sUF, vbError)
        Else
           gError 189570 'Call MsgBox("Parâmetros inválidos", vbError)
        End If
    End If
    
    InscricaoEstadual_Critica2 = SUCESSO
    
    Exit Function

Erro_InscricaoEstadual_Critica2:

    InscricaoEstadual_Critica2 = gErr
    
    Select Case gErr
    
        Case 189569
            'Call Rotina_Erro(vbOKOnly, "ERRO_INSCESTADUAL_INVALIDA", gErr, sUF)
        
        Case 189570
            'Call Rotina_Erro(vbOKOnly, "ERRO_INSCESTADUAL_PARAM_INVALIDOS", gErr)

        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189568)
        
    End Select
    
    Exit Function

End Function

