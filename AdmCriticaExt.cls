VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmCriticaExt"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function E_Zero(dNumero As Double) As Boolean
'Retorna True se dNumero for zero quando arredondado até precisão de FORMATO_ESTOQUE

Dim lPrecisao As Long

    'Número de casas decimais de FORMATO_ESTOQUE
    lPrecisao = Len(Mid(FORMATO_ESTOQUE, InStr(FORMATO_ESTOQUE, "."))) - 1
    
    'Arredonda
    dNumero = Round(dNumero, lPrecisao)
        
    'Testa se deu zero
    If dNumero = 0 Then
        E_Zero = True
    Else
        E_Zero = False
    End If

End Function

Function IniciaLetra(sString As String) As Boolean
'Testa se sString começa por letra. Retorna True se começar e False caso contrário.

Dim lErro As Long

On Error GoTo Erro_IniciaLetra
    
    If Len(sString) = 0 Then Error 25002

    If IsNumeric(sString) = True Then
        IniciaLetra = False
    Else
        IniciaLetra = True
    End If
    
    Exit Function

Erro_IniciaLetra:

    IniciaLetra = False
    
    Select Case Err

        Case 25002  'Se String for vazia não começa por letra
            
        Case Else
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142298)

    End Select

End Function

Function Data_Critica(ByVal sData As String) As Long
Attribute Data_Critica.VB_Description = "Retorna SUCESSO se sData for válida ou estiver vazia."

Dim lErro As Long
Dim lPosition1 As Long
Dim lPosition2 As Long
Dim iDia As Integer

On Error GoTo Erro_Data_Critica

    If IsDate(sData) Then
        'Posição do 1º separador "/" dentro de sData
        lPosition1 = InStr(1, sData, "/")
        'Posição do 2º separador "/" dentro de sData
        lPosition2 = InStr(lPosition1 + 1, sData, "/")
        'Não tem 2º separador então não é data
        If lPosition2 = 0 Then Error 25026
        'Isola os caracteres que antecedem a primeira barra "/"
        iDia = CInt(left(sData, lPosition1 - 1))
        'Testa se iDia coincide com o dia da Data
        If Not iDia = Day(sData) Then Error 25027
    ElseIf sData <> "  /  /  " Then
        Error 6001
    End If

    Data_Critica = SUCESSO
    
    Exit Function

Erro_Data_Critica:

    Data_Critica = Err
    
    Select Case Err

        Case 6001, 25026, 25027
           
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_DATA", Err, sData)

        Case Else
    
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142299)

    End Select

End Function

Function Cpf_Critica(ByVal sCpf As String) As Long

Dim lErro As Long
Dim dCpf As Double
Dim iDig(1 To STRING_CPF) As Integer
Dim iDigVer(1 To 2) As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iSum As Integer
Dim dCpfMaximo As Double

On Error GoTo Erro_Cpf_Critica

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
  
    Cpf_Critica = SUCESSO
    
    Exit Function

Erro_Cpf_Critica:

    Cpf_Critica = Err
    
    Select Case Err
    
        Case 6624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_NAO_NUMERICO", Err, sCpf)
        
        Case 6626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_MENOR_OU_IGUAL_ZERO", Err, sCpf)
        
        Case 6630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_NAO_INTEIRO", Err, sCpf)
        
        Case 6633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_INVALIDO", Err, sCpf)
        
        Case 6896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CPF_MAIOR_QUE_CPFMAXIMO", Err, sCpf)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142300)
        
    End Select
    
    Exit Function

End Function

Function Cgc_Critica(ByVal sCgc As String) As Long

Dim lErro As Long
Dim dCgc As Double
Dim iDig(1 To STRING_CGC) As Integer
Dim iDigVer(1 To 3) As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iSum As Integer
Dim iParcela As Integer
Dim dCgcMaximo As Double

On Error GoTo Erro_Cgc_Critica

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
  
    Cgc_Critica = SUCESSO
    
    Exit Function

Erro_Cgc_Critica:

    Cgc_Critica = Err
    
    Select Case Err
    
        Case 6641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_NUMERICO", Err, sCgc)
        
        Case 6642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_MENOR_OU_IGUAL_ZERO", Err, sCgc)
        
        Case 6643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INTEIRO", Err, sCgc)
        
        Case 6644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_INVALIDO", Err, sCgc)
        
       Case 6897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_MAIOR_QUE_CGCMAXIMO", Err, sCgc)
         
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142301)
        
    End Select
    
    Exit Function

End Function

Function Inteiro_Critica(sNumero As String) As Long
'Critica se é Tipo inteiro positivo (entre 1 e 9999)

Dim lErro As Long
Dim dNumero As Double

On Error GoTo Error_Inteiro_Critica

    If IsNumeric(sNumero) Then
        dNumero = 0#
        dNumero = CDbl(sNumero)
    Else
        Error 12310
    End If

    If dNumero <> Int(dNumero) Then Error 12358
    
    If dNumero < 1 Then Error 6748

    If dNumero > 9999 Then Error 6749
    
    Inteiro_Critica = SUCESSO

Exit Function

Error_Inteiro_Critica:

    Inteiro_Critica = Err

    Select Case Err

        Case 12310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sNumero)

        Case 12358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sNumero)
            
        Case 6748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case 6749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sNumero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142302)

    End Select

    Exit Function

End Function

Function Long_Critica(sNumero As String) As Long
'Critica se é tipo Long positivo (entre 1 e 999 999 999)

Dim lErro As Long
Dim dNumero As Double

On Error GoTo Error_Long_Critica
    
    If IsNumeric(sNumero) Then
        dNumero = 0#
        dNumero = CDbl(sNumero)
    Else
        Error 12311
    End If

    If dNumero <> Int(dNumero) Then Error 12359
    
    If dNumero < 1 Then Error 6750
    
    If dNumero > 999999999 Then Error 6751

    Long_Critica = SUCESSO

Exit Function

Error_Long_Critica:

    Long_Critica = Err

    Select Case Err
        
        Case 12311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sNumero)

        Case 12359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sNumero)
            
        Case 6750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case 6751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LONG_OVERFLOW", Err, sNumero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142303)

    End Select

    Exit Function

End Function

Function Porcentagem_Critica(sString As String) As Long
'Critica o valor passado em sString , se nao for entre 0 e 100 retorna erro.

Dim dValor As Double
Dim lErro As Long

On Error GoTo Error_Porcentagem_Critica

    If IsNumeric(sString) Then
        dValor = 0#
        dValor = CDbl(sString)
    Else
        Error 12530
    End If

    If dValor > 100 Or dValor < 0 Then Error 12314

    Porcentagem_Critica = SUCESSO

    Exit Function

Error_Porcentagem_Critica:

    Porcentagem_Critica = Err

    Select Case Err

        Case 12530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
            
        Case 12314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM", Err, dValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142304)

    End Select

    Exit Function

End Function


Function Valor_Critica(sValor As String) As Long
'critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO

Dim curTeste As Currency
Dim lErro As Long

On Error GoTo Erro_Valor_Critica

    curTeste = CCur(sValor)
    
    Valor_Critica = SUCESSO

    Exit Function

Erro_Valor_Critica:

    Valor_Critica = Err

   lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
        
    Exit Function

End Function

Function Valor_Inteiro_Critica(sValor As String) As Long
'critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.

Dim iTeste As Integer
Dim lErro As Long

On Error GoTo Erro_Valor_Inteiro_Critica

    iTeste = CInt(sValor)
    
    Valor_Inteiro_Critica = SUCESSO

    Exit Function

Erro_Valor_Inteiro_Critica:

    Valor_Inteiro_Critica = Err

    Exit Function

End Function

Function Valor_Positivo_Critica(sValor As String) As Long
'critica se o valor passado como parametro é valido e positivo. Se estiver tudo ok retorna SUCESSO

Dim curTeste As Currency
Dim lErro As Long

On Error GoTo Erro_Valor_Positivo_Critica

    curTeste = CCur(sValor)
    
    If curTeste <= 0 Then Error 10143
    
    Valor_Positivo_Critica = SUCESSO

    Exit Function

Erro_Valor_Positivo_Critica:

    Valor_Positivo_Critica = Err
    
    Select Case Err
    
        Case 10143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_POSITIVO", Err, sValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
            
    End Select
        
    Exit Function

End Function

Function Valor_Positivo_Critica_Double(sValor As String) As Long
'verifica se um valor double é positivo

Dim dTeste As Double
Dim lErro As Long

On Error GoTo Erro_Valor_Positivo_Critica_Double

    dTeste = CDbl(sValor)
    
    If dTeste <= 0 Then Error 60769
    
    Valor_Positivo_Critica_Double = SUCESSO

    Exit Function

Erro_Valor_Positivo_Critica_Double:

    Valor_Positivo_Critica_Double = Err

    Select Case Err
    
        Case 60769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_POSITIVO", Err, sValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
            
    End Select
        
    Exit Function

    Exit Function

End Function

Function Valor_NaoNegativo_Critica(sValor As String) As Long
'critica se o valor passado como parametro é valido e não é negativo. Se estiver tudo ok retorna SUCESSO

Dim curTeste As Currency
Dim lErro As Long

On Error GoTo Erro_Valor_NaoNegativo_Critica

    curTeste = CCur(sValor)
    
    If curTeste < 0 Then Error 6752
    
    Valor_NaoNegativo_Critica = SUCESSO

    Exit Function

Erro_Valor_NaoNegativo_Critica:

    Valor_NaoNegativo_Critica = Err
    
    Select Case Err
    
        Case 6752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, sValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", Err, sValor)
            
    End Select
        
    Exit Function

End Function

''Function Produto_Critica(ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'''Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'''Devolve os dados do Produto em objProduto
'''Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
''
''Dim lErro As Long
''Dim sProdutoFormatado As String
''
''On Error GoTo Erro_Produto_Critica
''
''
''    If Len(sProduto) > 0 Then
''
''        'Formata o código do Produto como no BD
''        lErro = CF("Produto_Formata",sProduto, sProdutoFormatado, iProdutoPreenchido)
''        If lErro Then Error 25039
''
''        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
''
''            'Carrega sProdutoFormatado em objProduto
''            objProduto.sCodigo = sProdutoFormatado
''
''            'Tenta ler o Produto
''            lErro = CF("Produto_Le",objProduto)
''            If lErro <> SUCESSO And lErro <> 28030 Then Error 25040
''
''            'Ausência de Produto no BD
''            If lErro = 28030 Then Error 25041
''
''            'Critica se é produto ATIVO
''            If objProduto.iAtivo <> Ativo Then Error 25042
''
''            'Critica se é produto NAO_GERENCIAL
''            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 25043
''
''        End If
''
''    End If
''
''    Produto_Critica = SUCESSO
''
''    Exit Function
''
''Erro_Produto_Critica:
''
''    Produto_Critica = Err
''
''    Select Case Err
''
''        Case 25039, 25040  'Erro criticado na rotina chamada
''
''        Case 25041  'Produto não está cadastrado
''
''        Case 25042
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
''
''        Case 25043
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142305)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Produto_Critica1(lComando As Long, ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'''Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'''Devolve os dados do Produto em objProduto
'''Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
''
''Dim lErro As Long
''Dim sProdutoFormatado As String
''
''On Error GoTo Erro_Produto_Critica1
''
''
''    If Len(sProduto) > 0 Then
''
''        'Formata o código do Produto como no BD
''        lErro = CF("Produto_Formata",sProduto, sProdutoFormatado, iProdutoPreenchido)
''        If lErro Then Error 20581
''
''        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
''
''            'Carrega sProdutoFormatado em objProduto
''            objProduto.sCodigo = sProdutoFormatado
''
''            'Tenta ler o Produto
''            lErro = CF("Produto_Le2",lComando, objProduto)
''            If lErro <> SUCESSO And lErro <> 20579 Then Error 20580
''
''            'Ausência de Produto no BD
''            If lErro = 20579 Then Error 20582
''
''            'Critica se é produto ATIVO
''            If objProduto.iAtivo <> Ativo Then Error 20583
''
''            'Critica se é produto NAO_GERENCIAL
''            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 20584
''
''        End If
''
''    End If
''
''    Produto_Critica1 = SUCESSO
''
''    Exit Function
''
''Erro_Produto_Critica1:
''
''    Produto_Critica1 = Err
''
''    Select Case Err
''
''        Case 20580, 20581  'Erro criticado na rotina chamada
''
''        Case 20582  'Produto não está cadastrado
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, sProduto)
''
''        Case 20583
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
''
''        Case 20584
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142306)
''
''    End Select
''
''    Exit Function
''
''End Function
''Function Produto_Critica_Filial(ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'''Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'''Devolve os dados do Produto em objProduto
'''Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
'''Verifica em Produto está relacionado com a FilialEmpresa.
''
''Dim lErro As Long
''Dim sProdutoFormatado As String
''Dim lComando As Long
''Dim iFilialEmpresa As Integer
''Dim vbMsgRes As VbMsgBoxResult
''
''On Error GoTo Erro_Produto_Critica_Filial
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 51380
''
''    lErro = CF("Produto_Critica",sProduto, objProduto, iProdutoPreenchido)
''    If lErro <> SUCESSO And lErro <> 25041 Then Error 51385
''    If lErro <> SUCESSO Then Error 51381
''
''    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM ProdutosFilial WHERE FilialEmpresa =? AND Produto = ?", iFilialEmpresa, giFilialEmpresa, objProduto.sCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 51382
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51383
''    If lErro = AD_SQL_SEM_DADOS Then Error 51384
''
''    Call Comando_Fechar(lComando)
''
''    Produto_Critica_Filial = SUCESSO
''
''    Exit Function
''
''Erro_Produto_Critica_Filial:
''
''    Produto_Critica_Filial = Err
''
''    Select Case Err
''
''        Case 51380
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 51381
''
''        Case 51382, 51383
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, giFilialEmpresa, sProduto)
''
''        Case 51384
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE", Err, sProduto, giFilialEmpresa)
''
''        Case 51385
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142307)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function

Function Long_Critica1(sNumero As String) As Long
'Critica se é tipo Long positivo (entre 1 e 99 999)

Dim lErro As Long
Dim dNumero As Double

On Error GoTo Error_Long_Critica1
    
    If IsNumeric(sNumero) Then
        dNumero = 0#
        dNumero = CDbl(sNumero)
    Else
        Error 44027
    End If

    If dNumero <> Int(dNumero) Then Error 44028
    
    If dNumero < 1 Then Error 44029
    
    If dNumero > 99999 Then Error 44030

    Long_Critica1 = SUCESSO

Exit Function

Error_Long_Critica1:

    Long_Critica1 = Err

    Select Case Err
        
        Case 44027, 44028, 44029, 44030

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142308)

    End Select

    Exit Function

End Function

Function Valor_Long_Critica(sValor As String) As Long
'critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.

Dim lTeste As Long
Dim lErro As Long

On Error GoTo Erro_Valor_Long_Critica

    lTeste = CLng(sValor)
    
    Valor_Long_Critica = SUCESSO

    Exit Function

Erro_Valor_Long_Critica:

    Valor_Long_Critica = Err

    Exit Function

End Function

Function Valor_Date_Critica(sValor As String) As Long
'critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.

Dim dtTeste As Date
Dim lErro As Long

On Error GoTo Erro_Valor_Date_Critica

    dtTeste = CDate(sValor)
    
    Valor_Date_Critica = SUCESSO

    Exit Function

Erro_Valor_Date_Critica:

    Valor_Date_Critica = Err

    Exit Function

End Function

Function Valor_Double_Critica(sValor As String) As Long
'critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.

Dim dTeste As Double
Dim lErro As Long

On Error GoTo Erro_Valor_Double_Critica

    dTeste = CDbl(sValor)
    
    Valor_Double_Critica = SUCESSO

    Exit Function

Erro_Valor_Double_Critica:

    Valor_Double_Critica = Err

    Exit Function

End Function

Function Porcentagem_Critica_Nao_Zero(sString As String) As Long
'Critica o valor passado em sString , se nao for entre -100 e 100 ou se for zero retorna erro.

Dim dValor As Double
Dim lErro As Long

On Error GoTo Error_Porcentagem_Critica_Nao_Zero

    If IsNumeric(sString) Then
        dValor = 0#
        dValor = CDbl(sString)
    Else
        Error 58075
    End If

    If dValor < -99.99 Then Error 58076
    
    If dValor = 0 Then Error 58077
    
    Porcentagem_Critica_Nao_Zero = SUCESSO

    Exit Function

Error_Porcentagem_Critica_Nao_Zero:

    Porcentagem_Critica_Nao_Zero = Err

    Select Case Err

        Case 58075
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
            
        Case 58076
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM2", Err, dValor)
        
        Case 58077
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ZERO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142309)

    End Select

    Exit Function

End Function

Function Porcentagem_Critica_Negativa(sString As String) As Long
'Critica o valor passado em sString aceitando valores negativos até -99.99

Dim dValor As Double
Dim lErro As Long

On Error GoTo Error_Porcentagem_Critica_Negativa

    If IsNumeric(sString) Then
        dValor = 0#
        dValor = CDbl(sString)
    Else
        Error 59063
    End If

    If dValor < -99.99 Then Error 59064
    
    Porcentagem_Critica_Negativa = SUCESSO

    Exit Function

Error_Porcentagem_Critica_Negativa:

    Porcentagem_Critica_Negativa = Err

    Select Case Err

        Case 59063
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
            
        Case 59064
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM2", Err, dValor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142310)

    End Select

    Exit Function

End Function

Function Porcentagem_Critica2(sString As String) As Long
'Critica o valor passado em sString , se nao for entre 0 e 99 retorna erro.

Dim dValor As Double
Dim lErro As Long

On Error GoTo Error_Porcentagem_Critica2

    If IsNumeric(sString) Then
        dValor = 0#
        dValor = CDbl(sString)
    Else
        Error 58454
    End If

    If dValor > 99 Or dValor < 0 Then Error 58455

    Porcentagem_Critica2 = SUCESSO

    Exit Function

Error_Porcentagem_Critica2:

    Porcentagem_Critica2 = Err

    Select Case Err

        Case 58454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sString)
            
        Case 58455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PORCENTAGEM3", Err, dValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142311)

    End Select

    Exit Function

End Function

Function Hora_Critica(ByVal sHora As String) As Long
'verifica se sHora contém uma hora válida

Dim lErro As Long
Dim dHora As Double

On Error GoTo Erro_Hora_Critica

    If IsDate(sHora) Then
    
        dHora = CDbl(CDate(sHora))
        
        If dHora >= 1 Then gError 89668
    
    Else
        gError 89669
    End If

    Hora_Critica = SUCESSO
    
    Exit Function

Erro_Hora_Critica:

    Hora_Critica = gErr
    
    Select Case gErr

        Case 89668, 89669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMATO_HORA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142312)

    End Select

End Function

Function Long_Critica2(sNumero As String) As Long
'Critica se é tipo Long positivo (entre 1 e 999 999 999)

Dim lErro As Long
Dim dNumero As Double

On Error GoTo Error_Long_Critica2
    
    If IsNumeric(sNumero) Then
        dNumero = 0#
        dNumero = CDbl(sNumero)
    Else
        Error 12311
    End If

    If dNumero <> Int(dNumero) Then Error 12359
    
    'If dNumero < 1 Then Error 6750
    
    If dNumero > 999999999 Then Error 6751

    Long_Critica2 = SUCESSO

Exit Function

Error_Long_Critica2:

    Long_Critica2 = Err

    Select Case Err
        
        Case 12311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, sNumero)

        Case 12359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sNumero)
            
        'Case 6750
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case 6751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LONG_OVERFLOW", Err, sNumero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142303)

    End Select

    Exit Function

End Function

Function CB_Calcula(ByVal sBanco As String, ByVal sMoeda As String, ByVal dValor As Double, ByVal dtVencimento As Date, ByVal sLivre As String, sCodigoBarra As String) As Long
'Calcula o Código de Barras para geração de Boleto Bancário - Alterado por Jorge Specian - 08/03/2005

Dim lErro As Long
Dim sCodigoSequencia As String
Dim dtDataBase As Date
Dim iFator As Integer
Dim iDV As Integer

On Error GoTo Erro_CB_Calcula
    
    'Database para calculo do fator
    dtDataBase = CDate("7/10/1997")
    iFator = DateDiff("d", dtDataBase, Format(dtVencimento, "dd/mm/yyyy"))
    dValor = Round(dValor * 100, 0)
    sBanco = Format(sBanco, "000")
    sLivre = Format(sLivre, "0000000000000000000000000")
    
    'Sequencia sem o DV
    sCodigoSequencia = sBanco & sMoeda & iFator & Format(dValor, "0000000000") & sLivre
    
    'Calculo do DV
    lErro = CB_Calcula_DV11(sCodigoSequencia, iDV)
    If lErro <> SUCESSO Then gError 198673
    
    'Monta a sequencia para o codigo de barras com o DV
    sCodigoBarra = left(sCodigoSequencia, 4) & iDV & right(sCodigoSequencia, 39)

    CB_Calcula = 0

    Exit Function
    
Erro_CB_Calcula:

    CB_Calcula = gErr
    
    Select Case gErr

        Case 198673
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198672)

    End Select
    
    Exit Function
    
End Function

Private Function CB_Calcula_DV11(ByVal sSequencia As String, iDV As Integer) As Long
'Calcula o Dígito Verificador do Código de Barras do Boleto Bancário - Alterado por Jorge Specian - 08/03/2005
'Cálculo através do módulo 11, com base de cálculo igual a 9

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim iResto As Integer
Dim iResultado As Integer
Dim sCaracter As String

On Error GoTo Erro_CB_Calcula_DV11

    iMultiplicador = 2
    
    For iContador = 1 To 43
       sCaracter = Mid(right(sSequencia, iContador), 1, 1)
       If iMultiplicador > 9 Then
              iMultiplicador = 2
             iNumero = 0
       End If
       iNumero = sCaracter * iMultiplicador
       iTotalNumero = iTotalNumero + iNumero
       iMultiplicador = iMultiplicador + 1
    Next
    
    iResto = iTotalNumero Mod 11
    
    iResultado = 11 - iResto
    
    If iResultado = 10 Or iResultado = 11 Then
        iDV = 1
    Else
        iDV = iResultado
    End If

    CB_Calcula_DV11 = 0

    Exit Function

Erro_CB_Calcula_DV11:

    CB_Calcula_DV11 = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198671)
    
    End Select
    
    Exit Function

End Function

Function CB_Converte_CB_LinDig_NaoArrec(ByVal sCodigoBarras As String, sLinhaDigitavel As String) As Long
'Calcula os Números da Linha Digitável de um Boleto Bancário

Dim lErro As Long
Dim sSeq1 As String
Dim sSeq2 As String
Dim sSeq3 As String
Dim sDVCodBarras As String
Dim sFatorVenc As String
Dim sValor As String
Dim iDv1 As Integer
Dim iDv2 As Integer
Dim iDv3 As Integer

On Error GoTo Erro_CB_Converte_CB_LinDig_NaoArrec
        
    'Separa as sequencias
    sSeq1 = left(sCodigoBarras, 4) & Mid(sCodigoBarras, 20, 5)
    sDVCodBarras = Mid(sCodigoBarras, 5, 1)
    sFatorVenc = Mid(sCodigoBarras, 6, 4)
    sValor = Mid(sCodigoBarras, 10, 10)
    sSeq2 = Mid(sCodigoBarras, 25, 10)
    sSeq3 = right(sCodigoBarras, 10)
    
    'Calcula os DVs
    lErro = CB_Calcula_DV10(sSeq1, iDv1)
    If lErro <> SUCESSO Then gError 198668
    
    lErro = CB_Calcula_DV10(sSeq2, iDv2)
    If lErro <> SUCESSO Then gError 198669
    
    lErro = CB_Calcula_DV10(sSeq3, iDv3)
    If lErro <> SUCESSO Then gError 198670
    
'    'Formata as sequencias
'    sSeq1 = Left(sSeq1 & iDv1, 5) & "." & Right(sSeq1 & iDv1, 5)
'    sSeq2 = Left(sSeq2 & iDv2, 5) & "." & Right(sSeq2 & iDv2, 6)
'    sSeq3 = Left(sSeq3 & iDv3, 5) & "." & Right(sSeq3 & iDv3, 6)
    
    'Concatena as sequencias
    sLinhaDigitavel = sSeq1 & CStr(iDv1) & sSeq2 & CStr(iDv2) & sSeq3 & CStr(iDv3) & sDVCodBarras & sFatorVenc & sValor

    CB_Converte_CB_LinDig_NaoArrec = 0

    Exit Function
    
Erro_CB_Converte_CB_LinDig_NaoArrec:

    CB_Converte_CB_LinDig_NaoArrec = gErr
    
    Select Case gErr
    
        Case 198668 To 198670
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198667)

    End Select
    
    Exit Function
    
End Function

Function CB_Calcula_DV10(ByVal sSequencia As String, iDigito As Integer) As Long
'Calcula o Digito Verificador no Módulo 10 para Linha Digitável de um Boleto Bancário
'Alterado por Jorge Specian - 09/03/2005

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim iDezenaSuperior As Integer

On Error GoTo Erro_CB_Calcula_DV10

    'Se nao for um valor numerico -> erro
    If Not IsNumeric(sSequencia) Then gError 198666
        
    'Inicia o multiplicador
    iMultiplicador = 2
    
    'Pega cada caracter do numero a partir da direita
    For iContador = Len(sSequencia) To 1 Step -1
        
        'Extrai o caracter e multiplica pelo multiplicador
        iNumero = Val(Mid(sSequencia, iContador, 1)) * iMultiplicador
        
        'Se o resultado for maior que nove soma os algarismos do resultado
        If iNumero > 9 Then
            
            iNumero = Val(left(iNumero, 1)) + Val(right(iNumero, 1))
        
        End If
        
        'Soma o resultado para totalização
        iTotalNumero = iTotalNumero + iNumero
        
        'Se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
        iMultiplicador = IIf(iMultiplicador = 2, 1, 2)
        
    Next

    If iTotalNumero < 10 Then
        iDezenaSuperior = 10
    Else
        iDezenaSuperior = 10 * (Val(left(CStr(iTotalNumero), Len(CStr(iTotalNumero)) - 1)) + 1)
    End If
    
    iDigito = iDezenaSuperior - iTotalNumero

    'verifica as exceções ( 10 -> DV=0 )
    If iDigito = 10 Then iDigito = 0
    
    CB_Calcula_DV10 = 0

    Exit Function

Erro_CB_Calcula_DV10:

    CB_Calcula_DV10 = gErr
    
    Select Case gErr
        
        Case 198666
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198665)
    
    End Select
    
    Exit Function

End Function

Function CB_Retona_Tipo(ByVal sCodigo As String, iLinhaDigitavel As Integer, iArrecadacao As Integer) As Long

On Error GoTo Erro_CB_Retona_Tipo

    iLinhaDigitavel = MARCADO
    iArrecadacao = DESMARCADO
    
    If Len(Trim(sCodigo)) Then

        If Len(Trim(sCodigo)) = 44 Then
            iLinhaDigitavel = DESMARCADO
        End If
        
        If left(sCodigo, 1) = "8" Then
            iArrecadacao = MARCADO
        End If

    End If
    
    CB_Retona_Tipo = SUCESSO

    Exit Function

Erro_CB_Retona_Tipo:

    CB_Retona_Tipo = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198636)
    
    End Select
    
    Exit Function

End Function

Function CB_LinDig_Retona_Mascara(ByVal sCodigo As String, sMascara As String) As Long

Dim iLinhaDigitavel As Integer
Dim iArrecadacao As Integer

On Error GoTo Erro_CB_LinDig_Retona_Mascara

    Call CB_Retona_Tipo(sCodigo, iLinhaDigitavel, iArrecadacao)

    If iArrecadacao = MARCADO Then
        sMascara = "###########.#.###########.#.###########.#.###########.#"
    Else
        sMascara = "#####.#####.#####.######.#####.######.#.#################"
    End If

    CB_LinDig_Retona_Mascara = SUCESSO

    Exit Function

Erro_CB_LinDig_Retona_Mascara:

    CB_LinDig_Retona_Mascara = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198637)
    
    End Select
    
    Exit Function

End Function

Function CB_Valida(ByVal sCodigo As String, Optional ByVal iTipo As Integer = CB_VALIDACAO_TIPO_NORMAL, Optional ByVal sDadosTitulo As String) As Long

Dim lErro As Long
Dim iDV As Integer
Dim sParte1 As String
Dim sParte2 As String
Dim sParte3 As String
Dim sParte4 As String
Dim iLinhaDigitavel As Integer
Dim iArrecadacao As Integer
Dim sCodigoBarras As String
Dim sLinDigFormatada As String

On Error GoTo Erro_CB_Valida

    sCodigo = Replace(Replace(sCodigo, ".", ""), " ", "")
    
    If Len(Trim(sCodigo)) < 33 Then gError 198674

    Call CB_Retona_Tipo(sCodigo, iLinhaDigitavel, iArrecadacao)

    If iLinhaDigitavel = DESMARCADO Then
        'Valida código de barras para não dar erro na conversão

        If iArrecadacao = MARCADO Then
        
            'retira o dígito verificador que fica na quarta casa
            sParte1 = left(sCodigo, 3) & right(sCodigo, 40)
            
            'Calcula o DV módulo 10
            lErro = CB_Calcula_DV10(sParte1, iDV)
            If lErro <> SUCESSO Then gError 198638
        
            'Código de barras errado
            If iDV <> CInt(Mid(sCodigo, 4, 1)) Then gError 198639
            
        Else
        
            'retira o dígito verificador que fica na quarta casa
            sParte1 = left(sCodigo, 4) & right(sCodigo, 39)
        
            'Calcula o DV módulo 11
            lErro = CB_Calcula_DV11(sParte1, iDV)
            If lErro <> SUCESSO Then gError 198640
        
            'Verifioca se o código de barras está errado
            If iDV <> CInt(Mid(sCodigo, 5, 1)) Then gError 198641
            
        End If
        
    Else
        Call CB_LinDig_Formata(sCodigo, sLinDigFormatada)
        
        If Len(Trim(sCodigo)) < 47 Then sCodigo = sCodigo & String(47 - Len(sCodigo), "0")
        
        'Linha digitável
    
        If iArrecadacao = MARCADO Then
        
            'Pega as quatro partes do código de barras sem o dígito verificador
            sParte1 = left(sCodigo, 11)
            sParte2 = Mid(sCodigo, 13, 11)
            sParte3 = Mid(sCodigo, 25, 11)
            sParte4 = Mid(sCodigo, 37, 11)
        
            lErro = CB_Calcula_DV10(sParte1, iDV)
            If lErro <> SUCESSO Then gError 198642
        
            'Verifica o primerio dígito
            If iDV <> CInt(Mid(sCodigo, 12, 1)) Then gError 198643
            
            lErro = CB_Calcula_DV10(sParte2, iDV)
            If lErro <> SUCESSO Then gError 198644
        
            'Verifica o segundo dígito
            If iDV <> CInt(Mid(sCodigo, 24, 1)) Then gError 198645
            
            lErro = CB_Calcula_DV10(sParte3, iDV)
            If lErro <> SUCESSO Then gError 198646
        
            'Verifica o terceiro dígito
            If iDV <> CInt(Mid(sCodigo, 36, 1)) Then gError 198647
            
            lErro = CB_Calcula_DV10(sParte4, iDV)
            If lErro <> SUCESSO Then gError 198648
        
            'Verifica o quarto dígito
            If iDV <> CInt(Mid(sCodigo, 48, 1)) Then gError 198649
            
            lErro = CB_Calcula_DV10(left(sParte1, 3) & right(sParte1, 7) & sParte2 & sParte3 & sParte4, iDV)
            If lErro <> SUCESSO Then gError 198650
        
            'Verifica dígito verificador geral
            If iDV <> CInt(Mid(sCodigo, 4, 1)) Then gError 198651
            
        Else
        
            'Pega as quatro partes do código de barras sem o dígito verificador
            sParte1 = left(sCodigo, 9)
            sParte2 = Mid(sCodigo, 11, 10)
            sParte3 = Mid(sCodigo, 22, 10)
        
            lErro = CB_Calcula_DV10(sParte1, iDV)
            If lErro <> SUCESSO Then gError 198652
        
            'Verifica o primerio dígito
            If iDV <> CInt(Mid(sCodigo, 10, 1)) Then gError 198653
            
            lErro = CB_Calcula_DV10(sParte2, iDV)
            If lErro <> SUCESSO Then gError 198654
        
            'Verifica o segundo dígito
            If iDV <> CInt(Mid(sCodigo, 21, 1)) Then gError 198655
            
            lErro = CB_Calcula_DV10(sParte3, iDV)
            If lErro <> SUCESSO Then gError 198656
            
            'Verifica o terceiro dígito
            If iDV <> CInt(Mid(sCodigo, 32, 1)) Then gError 198657
            
            lErro = CB_LinDig_Converte_CB_NaoArrec(sCodigo, sCodigoBarras)
            If lErro <> SUCESSO Then gError 198680
            
            'retira o dígito verificador que fica na quarta casa
            sParte4 = left(sCodigoBarras, 4) & right(sCodigoBarras, 39)
        
            'Calcula o DV módulo 11
            lErro = CB_Calcula_DV11(sParte4, iDV)
            If lErro <> SUCESSO Then gError 198681
        
            'Verifioca se o código de barras está errado
            If iDV <> CInt(Mid(sCodigoBarras, 5, 1)) Then gError 198682
            
            If iArrecadacao = DESMARCADO And Len(Trim(sCodigo)) > 47 Then gError 198674
            
        End If
       
    End If
    
    CB_Valida = SUCESSO

    Exit Function

Erro_CB_Valida:

    CB_Valida = gErr
    
    Select Case gErr
    
        Case 198638, 198640, 198642, 198644, 198646, 198648, 198650, 198652, 198654, 198656, 198680, 198681
        
        Case 198639, 198641, 198643, 198645, 198647, 198649, 198651, 198653, 198655, 198657, 198674, 198682
            
            If iTipo = CB_VALIDACAO_TIPO_BATCH Then
            
                Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOBARRAS_INVALIDO_BATCH", gErr, sLinDigFormatada, sDadosTitulo)
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOBARRAS_INVALIDO", gErr)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198658)
    
    End Select
    
    Exit Function

End Function

Function CB_Converte_LinDig(ByVal sCodigoBarras As String, sLinhaDigitavel As String) As Long

Dim lErro As Long
Dim iLinhaDigitavel As Integer
Dim iArrecadacao As Integer
Dim iDV As Integer

On Error GoTo Erro_CB_Converte_LinDig

    sCodigoBarras = Replace(Replace(sCodigoBarras, ".", ""), " ", "")
    sLinhaDigitavel = ""

    Call CB_Retona_Tipo(sCodigoBarras, iLinhaDigitavel, iArrecadacao)

    lErro = CB_Valida(sCodigoBarras)
    If lErro <> SUCESSO Then gError 198659
    
    If iLinhaDigitavel <> MARCADO Then
            
        If iArrecadacao = DESMARCADO Then
        
            'Converte para a linha digitável
            lErro = CB_Converte_CB_LinDig_NaoArrec(sCodigoBarras, sLinhaDigitavel)
            If lErro <> SUCESSO Then gError 198660
            
        Else
            'Insere os DVs
            
            lErro = CB_Calcula_DV10(left(sCodigoBarras, 11), iDV)
            If lErro <> SUCESSO Then gError 198661
        
            sLinhaDigitavel = left(sCodigoBarras, 11) & CStr(iDV)
        
            lErro = CB_Calcula_DV10(Mid(sCodigoBarras, 12, 11), iDV)
            If lErro <> SUCESSO Then gError 198662
        
            sLinhaDigitavel = sLinhaDigitavel & Mid(sCodigoBarras, 12, 11) & CStr(iDV)
        
            lErro = CB_Calcula_DV10(Mid(sCodigoBarras, 23, 11), iDV)
            If lErro <> SUCESSO Then gError 198663
        
            sLinhaDigitavel = sLinhaDigitavel & Mid(sCodigoBarras, 23, 11) & CStr(iDV)
        
            lErro = CB_Calcula_DV10(Mid(sCodigoBarras, 34, 11), iDV)
            If lErro <> SUCESSO Then gError 198664
        
            sLinhaDigitavel = sLinhaDigitavel & Mid(sCodigoBarras, 34, 11) & CStr(iDV)
        
        End If
    Else
        sLinhaDigitavel = sCodigoBarras
    End If
    
    CB_Converte_LinDig = SUCESSO

    Exit Function

Erro_CB_Converte_LinDig:

    CB_Converte_LinDig = gErr
    
    Select Case gErr
    
        Case 198659 To 198664
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198659)
    
    End Select
    
    Exit Function

End Function

Function CB_LinDig_Converte_CB(ByVal sLinhaDigitavel As String, sCodigoBarras As String) As Long

Dim lErro As Long
Dim iLinhaDigitavel As Integer
Dim iArrecadacao As Integer

On Error GoTo Erro_CB_LinDig_Converte_CB

    sLinhaDigitavel = Replace(Replace(sLinhaDigitavel, ".", ""), " ", "")
    sCodigoBarras = ""

    Call CB_Retona_Tipo(sLinhaDigitavel, iLinhaDigitavel, iArrecadacao)

    lErro = CB_Valida(sLinhaDigitavel)
    If lErro <> SUCESSO Then gError 198675
        
    If iLinhaDigitavel = MARCADO Then
            
        If iArrecadacao = DESMARCADO Then
    
            'Preenche com brancos até completar o tamanho
            If Len(Trim(sLinhaDigitavel)) < 47 Then sLinhaDigitavel = sLinhaDigitavel & String(47 - Len(sLinhaDigitavel), "0")
    
            'Converte para a linha digitável
            lErro = CB_LinDig_Converte_CB_NaoArrec(sLinhaDigitavel, sCodigoBarras)
            If lErro <> SUCESSO Then gError 198676
            
        Else
        
            sCodigoBarras = left(sLinhaDigitavel, 11)
            sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 13, 11)
            sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 25, 11)
            sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 37, 11)
        
        End If
    Else
        sCodigoBarras = sLinhaDigitavel
    End If

    CB_LinDig_Converte_CB = SUCESSO

    Exit Function

Erro_CB_LinDig_Converte_CB:

    CB_LinDig_Converte_CB = gErr
    
    Select Case gErr
    
        Case 198675 To 198676
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198677)
    
    End Select
    
    Exit Function

End Function

Function CB_LinDig_Converte_CB_NaoArrec(ByVal sLinhaDigitavel As String, sCodigoBarras As String) As Long
'Calcula os Números da Linha Digitável de um Boleto Bancário

Dim lErro As Long

On Error GoTo Erro_CB_LinDig_Converte_CB_NaoArrec

    sCodigoBarras = Mid(sLinhaDigitavel, 1, 4)
    sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 33, 1)
    sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 34, 4)
    sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 38, 10)
    sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 5, 5)
    sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 11, 10)
    sCodigoBarras = sCodigoBarras & Mid(sLinhaDigitavel, 22, 10)
    
    CB_LinDig_Converte_CB_NaoArrec = 0

    Exit Function
    
Erro_CB_LinDig_Converte_CB_NaoArrec:

    CB_LinDig_Converte_CB_NaoArrec = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198679)

    End Select
    
    Exit Function
    
End Function

Function CB_LinDig_Formata(ByVal sLinhaDigitavel As String, sLinhaDigitavelFormatada As String) As Long
'Calcula os Números da Linha Digitável de um Boleto Bancário

Dim lErro As Long
Dim iLinhaDigitavel As Integer
Dim iArrecadacao As Integer

On Error GoTo Erro_CB_LinDig_Formata

    Call CB_Retona_Tipo(sLinhaDigitavel, iLinhaDigitavel, iArrecadacao)
    
    If iLinhaDigitavel = MARCADO Then

        sLinhaDigitavel = Replace(Replace(sLinhaDigitavel, ".", ""), " ", "")

        If iArrecadacao = DESMARCADO Then
        
            sLinhaDigitavelFormatada = Mid(sLinhaDigitavel, 1, 5) & "."
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 6, 5) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 11, 5) & "."
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 16, 6) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 22, 5) & "."
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 27, 6) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 33, 1) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 34, 14)
            
        Else
        
            sLinhaDigitavelFormatada = Mid(sLinhaDigitavel, 1, 11) & "-"
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 12, 1) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 13, 11) & "-"
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 24, 1) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 25, 11) & "-"
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 36, 1) & " "
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 37, 11) & "-"
            sLinhaDigitavelFormatada = sLinhaDigitavelFormatada & Mid(sLinhaDigitavel, 48, 1)
        
        End If
            
    End If
    
    CB_LinDig_Formata = 0

    Exit Function
    
Erro_CB_LinDig_Formata:

    CB_LinDig_Formata = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 198679)

    End Select
    
    Exit Function
    
End Function



