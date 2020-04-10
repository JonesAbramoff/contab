VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoLoja"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function TP_Produto_Le_Col(X As ClassProdReferencia, x1 As ClassProdCodBarra, x2 As ClassProdNome, sProduto As String, objProduto As ClassProduto) As Long
'preenche objProduto lendo no bd a partir da referencia ou codigo de barras

Dim Produto As MaskEdBox
Dim eTipoProduto As enumTipoProd
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim bAchou As Boolean
Dim objProduto1 As ClassProduto
Dim sCodBarras As String
Dim iIndice As Integer
Dim sProdRef As String

On Error GoTo Erro_TP_Produto_Le_Col

    Set objProduto = New ClassProduto
    
    'Verifica se código do produto é Código de Barras ou referência
    'sProduto = Trim(Produto.Text)



    'pesquisa se é um codigo de barras
    Set objProduto = x1.Busca(sProduto)
    
    If giDebug = 1 Then MsgBox ("x1.Busca")
    
    'Não encontrou o Produto
    If objProduto Is Nothing Then
        
        'pesquisa pela referencia
        Set objProduto = X.Busca(sProduto)
        
        If giDebug = 1 Then MsgBox ("x.Busca")
        
        
        'Não encontrou o Produto
        If objProduto Is Nothing Then
        
            'pesquisa pelo nome
            Set objProduto = x2.Busca(sProduto)
            
            If giDebug = 1 Then MsgBox ("x2.Busca")
            
            
            'Não encontrou o Produto
            If objProduto Is Nothing Then
                If Len(Trim(sProduto)) = 13 And IsNumeric(sProduto) Then
                    sProdRef = Mid(sProduto, 2, 5)
                                
                    'pesquisa pela referencia
                    Set objProduto1 = X.Busca(sProdRef)
                    
                    If Not (objProduto1 Is Nothing) Then
                        If objProduto1.iUsaBalanca <> NAO_USA_BALANCA Then
                            objProduto1.iUsaBalanca = USA_BALANCA_PARA_ETIQUETA
                            Set objProduto = objProduto1
                        End If
                    End If
                End If
            End If
            

            'Não encontrou o Produto
            If objProduto Is Nothing Then gError 99605
        
        End If
        
    End If
            
    TP_Produto_Le_Col = SUCESSO

    Exit Function

Erro_TP_Produto_Le_Col:

    TP_Produto_Le_Col = gErr

    Select Case gErr
        
        Case 99605, 99606
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_NAO_CADASTRADO1, gErr, sProduto)
                   
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 153624)

    End Select

    Exit Function
    
End Function

Private Function Tipo_Produto(ByVal sProduto As String) As enumTipoProd
'Faz a verificação se é Código de Barras ou produto
    
    'Se o tamanho da String estiver entre 13 e 15, então é um código de Barras
    If Len(Trim(sProduto)) <= PRODUTO_CODBARRAS_MAX And Len(Trim(sProduto)) >= PRODUTO_CODBARRAS_MIN And IsNumeric(sProduto) Then
    
        Tipo_Produto = TIPO_PROD_COD_BARRA

    Else
        Tipo_Produto = TIPO_PROD_REFERENCIA

    End If

    Exit Function

End Function

Public Function TP_Produto_Le_Col1(X As ClassProdReferencia, x1 As ClassProdCodBarra, x2 As ClassProdNome, objProdMask As Object, objProduto As ClassProduto) As Long
'preenche objProduto lendo no bd a partir da referencia ou codigo de barras

Dim sProduto As String
Dim Produto As MaskEdBox
Dim eTipoProduto As enumTipoProd
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim bAchou As Boolean
Dim objProduto1 As ClassProduto
Dim sCodBarras As String
Dim iIndice As Integer

On Error GoTo Erro_TP_Produto_Le_Col1

    Set Produto = objProdMask
    Set objProduto = New ClassProduto
    
    'Verifica se código do produto é Código de Barras ou referência
    'sProduto = Trim(Produto.Text)
    sProduto = Produto.Text
            
    Set objProduto = x1.Busca(sProduto)
    
    'Não encontrou o Produto
    If objProduto Is Nothing Then
            
        Set objProduto = X.Busca(sProduto)
            
        'Não encontrou o Produto
        If objProduto Is Nothing Then
            
            Set objProduto = x2.Busca(sProduto)
            
            'Não encontrou o Produto
            If objProduto Is Nothing Then gError 99606
            
        End If
            
    End If

    TP_Produto_Le_Col1 = SUCESSO

    Exit Function

Erro_TP_Produto_Le_Col1:

    TP_Produto_Le_Col1 = gErr

    Select Case gErr
        
        Case 99605, 99606
                               
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 153625)

    End Select

    Exit Function
    
End Function

Function Caixa_TP_Cliente_Le(objClienteMaskEdBox As Object, objCliente As ClassCliente, iCodFilial As Integer) As Long
'Lê o Cliente com Código ou NomeRed ou CGC/CPF em sCliente
'Devolve em objCliente e o código da filial em iCodFilial

Dim sCliente As String
Dim Cliente As Object
Dim eTipoCliente As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Caixa_TP_Cliente_Le

    Set Cliente = objClienteMaskEdBox
    sCliente = Trim(Cliente.Text)
    eTipoCliente = Tipo_Cliente(sCliente)

    Select Case eTipoCliente

    Case TIPO_STRING

        objCliente.sNomeReduzido = sCliente
        lErro = CF_ECF("Caixa_Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO Then gError 109541
        
        Cliente.Text = sCliente

    Case TIPO_CODIGO

        objCliente.lCodigo = CLng(sCliente)
        lErro = CF_ECF("Caixa_Cliente_Le_Codigo", objCliente)
        If lErro <> SUCESSO Then gError 109542
                
        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sCliente)
        If lErro <> SUCESSO Then gError 109543
        
        objCliente.sCGC = sCliente
        
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF_ECF("Caixa_Cliente_Le_CPFCGC", objCliente)
        If lErro <> SUCESSO Then gError 109544
                
        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sCliente)
        If lErro <> SUCESSO Then gError 109545
        
        objCliente.sCGC = sCliente
        
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF_ECF("Caixa_Cliente_Le_CPFCGC", objCliente)
        If lErro <> SUCESSO Then gError 109546
        
        Cliente.Text = objCliente.sNomeReduzido
    
    Case TIPO_OVERFLOW

        gError 109547

    Case TIPO_DECIMAL

        gError 109548

    Case TIPO_NAO_POSITIVO

        gError 109549

    End Select

    Caixa_TP_Cliente_Le = SUCESSO

    Exit Function

Caixa_TP_Cliente_Le:

    Caixa_TP_Cliente_Le = gErr

    Select Case gErr
                
        Case 109541 To 109546
        
        Case 109547
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_CGC_OVERFLOW1, gErr, sCliente)

        Case 109548
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_INTEIRO1, gErr, sCliente)

        Case 109549
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_POSITIVO1, gErr, sCliente)

        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 153626)

    End Select

    Exit Function

End Function

Private Function Tipo_Cliente(ByVal sCliente As String) As enumTipo

    If Len(Trim(sCliente)) = 0 Then
        Tipo_Cliente = TIPO_VAZIO
    ElseIf Not IsNumeric(sCliente) Then
        Tipo_Cliente = TIPO_STRING
    ElseIf Int(CDbl(sCliente)) <> CDbl(sCliente) Then
        Tipo_Cliente = TIPO_DECIMAL
    ElseIf CDbl(sCliente) <= 0 Then
        Tipo_Cliente = TIPO_NAO_POSITIVO
    ElseIf Len(Trim(sCliente)) > STRING_CGC Then
        Tipo_Cliente = TIPO_OVERFLOW
    ElseIf Len(Trim(sCliente)) > STRING_CPF Then
        Tipo_Cliente = TIPO_CGC
    ElseIf CDbl(sCliente) > NUM_MAX_CLIENTES Then
        Tipo_Cliente = TIPO_CPF
    Else
        Tipo_Cliente = TIPO_CODIGO
    End If

End Function

Public Function TP_Caixa_Le1(ByVal objCaixaMaskEdBox As Object, ByVal objCaixa As ClassCaixa) As Long
'Lê a Caixa com Código ou NomeRed em objCaixaMaskEdBox.Text
'Devolve em objCaixa. Coloca código-NomeReduzido no .Text

Dim Caixa As Object
Dim lErro As Long

On Error GoTo Erro_TP_Caixa_Le1

    Set Caixa = objCaixaMaskEdBox

    'Se é do tipo código
    If objCaixa.iCodigo > 0 Then

        If giFilialEmpresa = EMPRESA_TODA Then
        
            'verifica se o codigo existe
            lErro = CF("Caixas_Le1", objCaixa)
            If lErro <> SUCESSO And lErro <> 79405 Then gError 116174
        
        Else
        
            objCaixa.iFilialEmpresa = giFilialEmpresa
            
            'verifica se o codigo existe
            lErro = CF("Caixas_Le", objCaixa)
            If lErro <> SUCESSO And lErro <> 79405 Then gError 116174
        
        End If
        
        'sem dados
        If lErro = 79405 Then gError 116175

        Caixa.Text = CStr(objCaixa.iCodigo) & Separador & objCaixa.sNomeReduzido

    Else  'Se é do tipo String

         objCaixa.sNomeReduzido = Caixa

         'verifica se o nome reduzido existe
         lErro = CF("Caixa_Le_NomeReduzido", objCaixa)
         If lErro <> SUCESSO And lErro <> 79582 Then gError 116176

         'sem dados
         If lErro = 79582 Then gError 116177

         'NomeControle.text recebe codigo - nome_reduzido
         Caixa.Text = CStr(objCaixa.iCodigo) & Separador & objCaixa.sNomeReduzido

    End If

    TP_Caixa_Le1 = SUCESSO

    Exit Function

Erro_TP_Caixa_Le1:

    TP_Caixa_Le1 = gErr

    Select Case gErr

        Case 116176, 116174 'Tratados nas rotinas chamadas

        Case 116175, 116177 'Caixa com Codigo / NomeReduzido não cadastrado

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153627)

    End Select

    Exit Function

End Function


Public Function TP_Operador_Le(ByVal objOperadorMaskEdBox As Object, ByVal objOperador As ClassOperador) As Long
'Lê o Operador com Código ou NomeRed em objOperadorMaskEdBox.Text
'Devolve em objOperador. Coloca código-NomeReduzido no .Text

Dim Operador As Object
Dim lErro As Long

On Error GoTo Erro_TP_Operador_Le

    Set Operador = objOperadorMaskEdBox
        
    'Guarda o código do operador no objeto
    objOperador.iCodigo = Codigo_Extrai(Operador.Text)
    
    'Se é do tipo código
    If objOperador.iCodigo > 0 Then
    
        'Se está acessando empresa toda
        If giFilialEmpresa = EMPRESA_TODA Then
        
            'Verifica se o operador está cadastrado sem filtro de filial
            lErro = CF("Operador_Le1", objOperador)
            If lErro <> SUCESSO And lErro <> 102480 Then gError 102476
        
            'sem dados
            If lErro = 102480 Then gError 117117
        
        'Senão, ou seja, se está acessando uma filial específica
        Else
        
            'Guarda a filial da empresa que está sendo acessada
            objOperador.iFilialEmpresa = giFilialEmpresa
            
            'Verifica se o operador está cadastrado, considerando o filtro de filial
            lErro = CF("Operador_Le", objOperador)
            If lErro <> SUCESSO And lErro <> 81026 Then gError 117116
            
            'sem dados
            If lErro = 81026 Then gError 117117
        
        End If
        
        'Joga no controle código-nome do operador
        Operador.Text = CStr(objOperador.iCodigo) & Separador & objOperador.sNome
        
    Else  'Se é do tipo String
                     
        objOperador.sNome = Trim(Operador.Text)
        
        'Se está acessando empresa toda
        If giFilialEmpresa = EMPRESA_TODA Then
         
            'Verifica se o operador está cadastrado sem filtro de filial
            lErro = CF("Operador_Le_NomeReduzido1", objOperador)
            If lErro <> SUCESSO And lErro <> 102484 Then gError 102485
            
            'sem dados
            If lErro = 102484 Then gError 117119
        
        Else
         
            'Guarda a filial da empresa que está sendo acessada
            objOperador.iFilialEmpresa = giFilialEmpresa
            
            'Verifica se o operador está cadastrado, considerando o filtro de filial
            lErro = CF("Operador_Le_NomeReduzido", objOperador)
            If lErro <> SUCESSO And lErro <> 116189 Then gError 117118
            
            'sem dados
            If lErro = 116189 Then gError 117119
        
        End If
        
         'NomeControle.text recebe codigo - nome_reduzido
         Operador.Text = CStr(objOperador.iCodigo) & Separador & objOperador.sNome
        
    End If

    TP_Operador_Le = SUCESSO

    Exit Function

Erro_TP_Operador_Le:

    TP_Operador_Le = gErr

    Select Case gErr
        
        Case 117118, 117116, 102476, 102485

        Case 117117, 117119 'Caixa com Codigo / NomeReduzido não cadastrado
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153628)

    End Select

    Exit Function

End Function


Public Function TP_Produto_Le_ColNaka(X As ClassProdReferencia, x1 As ClassProdCodBarra, x2 As ClassProdNome, objProdMask As Object, objProduto As ClassProduto) As Long
'preenche objProduto lendo no bd a partir da referencia ou codigo de barras

Dim sProduto As String
Dim Produto As MaskEdBox
Dim eTipoProduto As enumTipoProd
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim bAchou As Boolean
Dim objProduto1 As ClassProduto
Dim sCodBarras As String
Dim iIndice As Integer

On Error GoTo Erro_TP_Produto_Le_ColNaka

    Set Produto = objProdMask
    Set objProduto = New ClassProduto
    
    'Verifica se código do produto é Código de Barras ou referência
    'sProduto = Trim(Produto.Text)
    sProduto = Produto.Text

    'pesquisa se é um codigo de barras
    Set objProduto = x1.BuscaNaka(sProduto)
    
    'Não encontrou o Produto
    If objProduto Is Nothing Then
        
        'pesquisa pela referencia
        Set objProduto = X.Busca(sProduto)
        
        'Não encontrou o Produto
        If objProduto Is Nothing Then
        
            'pesquisa pelo nome
            Set objProduto = x2.Busca(sProduto)
            
            'Não encontrou o Produto
            If objProduto Is Nothing Then gError 99605
        
        End If
        
    End If
            
    TP_Produto_Le_ColNaka = SUCESSO

    Exit Function

Erro_TP_Produto_Le_ColNaka:

    TP_Produto_Le_ColNaka = gErr

    Select Case gErr
        
        Case 99605, 99606
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_NAO_CADASTRADO1, gErr, sProduto)
                   
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 153624)

    End Select

    Exit Function
    
End Function

Public Function TP_Produto_Le_Col2(X As ClassProdCodigo, sProduto As String, objProduto As ClassProduto) As Long
'preenche objProduto lendo no bd a partir do  codigo

Dim Produto As MaskEdBox
Dim eTipoProduto As enumTipoProd
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim bAchou As Boolean
Dim objProduto1 As ClassProduto
Dim sCodBarras As String
Dim iIndice As Integer

On Error GoTo Erro_TP_Produto_Le_Col2

    Set objProduto = New ClassProduto
    
    'Verifica se código do produto é Código de Barras ou referência
    'sProduto = Trim(Produto.Text)

    'pesquisa pelo codigo
    Set objProduto = X.Busca(sProduto)
    
    'Não encontrou o Produto
    If objProduto Is Nothing Then gError 210472
        
    TP_Produto_Le_Col2 = SUCESSO

    Exit Function

Erro_TP_Produto_Le_Col2:

    TP_Produto_Le_Col2 = gErr

    Select Case gErr
        
        Case 210472
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_NAO_CADASTRADO1, gErr, sProduto)
                   
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 153624)

    End Select

    Exit Function
    
End Function

