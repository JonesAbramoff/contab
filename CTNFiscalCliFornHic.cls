VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalCliFornHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Property Get PercDesconto(ByVal objCT As Object) As Object
'     Set PercDesconto = objCT.objUserControl.Controls("PercDesconto")
'End Property

Public Property Get RazaoSocial(ByVal objCT As Object) As Object
     Set RazaoSocial = objCT.objUserControl.Controls("RazaoSocial")
End Property

Public Property Get CNPJ(ByVal objCT As Object) As Object
     Set CNPJ = objCT.objUserControl.Controls("CNPJ")
End Property

Function NFCF_Fornecedor_Traz_NF_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_NFCF_Fornecedor_Traz_NF_Tela

    If objNFiscal.lFornecedor <> 0 Then
        
         objFornecedor.lCodigo = objNFiscal.lFornecedor
        
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 131843
               
        'Se não encontrou o cliente => erro
        If lErro = SUCESSO Then
            
            objFilialForn.lCodFornecedor = objFornecedor.lCodigo
            objFilialForn.iCodFilial = objNFiscal.iFilialForn
    
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 131844
                       
            CNPJ(objCT).Caption = objFilialForn.sCgc
            RazaoSocial(objCT).Caption = objFornecedor.sRazaoSocial
        
            lErro = Formata_CNPJ(objCT)
            If lErro <> SUCESSO Then gError 131845
        
        End If
    
    End If
    
    NFCF_Fornecedor_Traz_NF_Tela = SUCESSO
     
    Exit Function
    
Erro_NFCF_Fornecedor_Traz_NF_Tela:

    NFCF_Fornecedor_Traz_NF_Tela = gErr
     
    Select Case gErr
          
        Case 131843 To 131845 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179273)
     
    End Select
     
    Exit Function
    
End Function

Function NFCF_Fornecedor_Validate(ByVal objCT As Object, ByVal lFornecedor As Long, ByVal iFilial As Integer) As Long

Dim lErro As Long, dSaldoAtraso As Double

Dim objFornecedor As New ClassFornecedor
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_NFCF_Fornecedor_Validate

    If lFornecedor <> 0 Then
                
         objFornecedor.lCodigo = lFornecedor
        
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 131840
               
        'Se não encontrou o cliente => erro
        If lErro = SUCESSO Then
            
            objFilialForn.lCodFornecedor = objFornecedor.lCodigo
            objFilialForn.iCodFilial = iFilial
    
            lErro = CF("FilialFornecedor_Le", objFilialForn)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 131841
                       
            CNPJ(objCT).Caption = objFilialForn.sCgc
            RazaoSocial(objCT).Caption = objFornecedor.sRazaoSocial
        
            lErro = Formata_CNPJ(objCT)
            If lErro <> SUCESSO Then gError 131842
        
        
        End If
    
    End If
    
    NFCF_Fornecedor_Validate = SUCESSO
     
    Exit Function
    
Erro_NFCF_Fornecedor_Validate:

    NFCF_Fornecedor_Validate = gErr
     
    Select Case gErr
          
        Case 131840 To 131842 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179274)
     
    End Select
     
    Exit Function

End Function

Public Function NFCF_Limpa_Tela(ByVal objCT As Object) As Long

    CNPJ(objCT).Caption = ""
    RazaoSocial(objCT).Caption = ""

    NFCF_Limpa_Tela = SUCESSO
    
End Function

Public Function Formata_CNPJ(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sFormato As String

On Error GoTo Erro_Formata_CNPJ
    
    If Len(Trim(CNPJ(objCT).Caption)) <> 0 Then
    
        Select Case Len(Trim(CNPJ(objCT).Caption))
    
            Case STRING_CPF
                
                'Critica Cpf
                lErro = Cpf_Critica(CNPJ(objCT).Caption)
                If lErro <> SUCESSO Then gError 131846
                
                'Formata e coloca na Tela
                sFormato = "000\.000\.000-00; ; ; "
                CNPJ(objCT).Caption = Format(CNPJ(objCT).Caption, sFormato)
    
            Case STRING_CGC 'CGC
                
                'Critica CGC
                lErro = Cgc_Critica(CNPJ(objCT).Caption)
                If lErro <> SUCESSO Then gError 131847
                
                'Formata e Coloca na Tela
                sFormato = "00\.000\.000\/0000-00; ; ; "
                CNPJ(objCT).Caption = Format(CNPJ(objCT).Caption, sFormato)
    
            Case Else
                    
                gError 131848
    
        End Select
    
    End If
    
    Formata_CNPJ = SUCESSO
    
    Exit Function

Erro_Formata_CNPJ:

    Formata_CNPJ = gErr

    Select Case gErr

        Case 131846 To 131847

        Case 131848
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179275)

    End Select

    Exit Function

End Function

Public Function NFCF_Fornecedor_Filial_Validate(ByVal objCT As Object, ByVal objFilialForn As ClassFilialFornecedor) As Long
    
Dim lErro As Long

On Error GoTo Erro_NFCF_Fornecedor_Filial_Validate

    CNPJ(objCT).Caption = objFilialForn.sCgc
    
    lErro = Formata_CNPJ(objCT)
    If lErro <> SUCESSO Then gError 131849
    
    NFCF_Fornecedor_Filial_Validate = SUCESSO
    
    Exit Function

Erro_NFCF_Fornecedor_Filial_Validate:

    NFCF_Fornecedor_Filial_Validate = gErr

    Select Case gErr
    
        Case 131849

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179276)

    End Select

    Exit Function
    
End Function

Function NFCF_Cliente_Traz_NF_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_NFCF_Cliente_Traz_NF_Tela

    If objNFiscal.lCliente <> 0 Then
        
        'Guarda em objCliente o nome reduzido que foi informado no Orçamento
        objCliente.lCodigo = objNFiscal.lCliente
        
        'Lê os dados do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 131850
        
        'Se não encontrou o cliente => erro
        If lErro = SUCESSO Then
            
            objFilialCliente.lCodCliente = objCliente.lCodigo
            objFilialCliente.iCodFilial = objNFiscal.iFilialCli
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 131851
            
            CNPJ(objCT).Caption = objFilialCliente.sCgc
            RazaoSocial(objCT).Caption = objCliente.sRazaoSocial
        
            lErro = Formata_CNPJ(objCT)
            If lErro <> SUCESSO Then gError 131852
        
        End If
    
    End If
    
    NFCF_Cliente_Traz_NF_Tela = SUCESSO
     
    Exit Function
    
Erro_NFCF_Cliente_Traz_NF_Tela:

    NFCF_Cliente_Traz_NF_Tela = gErr
     
    Select Case gErr
          
        Case 131850 To 131852 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179277)
     
    End Select
     
    Exit Function
    
End Function

Function NFCF_Cliente_Validate(ByVal objCT As Object, ByVal lCliente As Long, ByVal iFilial As Integer) As Long

Dim lErro As Long, dSaldoAtraso As Double

Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_NFCF_Cliente_Validate

    If lCliente <> 0 Then
        
        'Guarda em objCliente o nome reduzido que foi informado no Orçamento
        objCliente.lCodigo = lCliente
        
        'Lê os dados do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 131853
        
        'Se não encontrou o cliente => erro
        If lErro = SUCESSO Then
        
            objFilialCliente.lCodCliente = lCliente
            objFilialCliente.iCodFilial = iFilial
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 131854
            
            CNPJ(objCT).Caption = objFilialCliente.sCgc
            RazaoSocial(objCT).Caption = objCliente.sRazaoSocial
        
            lErro = Formata_CNPJ(objCT)
            If lErro <> SUCESSO Then gError 131855
        
        End If
    
    End If
    
    NFCF_Cliente_Validate = SUCESSO
     
    Exit Function
    
Erro_NFCF_Cliente_Validate:

    NFCF_Cliente_Validate = gErr
     
    Select Case gErr
          
        Case 131853 To 131855 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179278)
     
    End Select
     
    Exit Function

End Function

Public Function NFCF_Cliente_Filial_Validate(ByVal objCT As Object, ByVal objFilialCliente As ClassFilialCliente) As Long
    
Dim lErro As Long

On Error GoTo Erro_NFCF_Cliente_Filial_Validate

    CNPJ(objCT).Caption = objFilialCliente.sCgc
    
    lErro = Formata_CNPJ(objCT)
    If lErro <> SUCESSO Then gError 131859
    
    NFCF_Cliente_Filial_Validate = SUCESSO
    
    Exit Function

Erro_NFCF_Cliente_Filial_Validate:

    NFCF_Cliente_Filial_Validate = gErr

    Select Case gErr
    
        Case 131859

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179279)

    End Select

    Exit Function
    
End Function

'Public Sub PercDesconto_Change(ByVal objCT As Object)
'
'    objCT.iAlterado = REGISTRO_ALTERADO
'    objCT.iValorDescontoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercDesconto_Validate(ByVal objCT As Object, Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorProdutos As Double
'Dim iIndice As Integer
'Dim dPercDesconto As Double
'
'On Error GoTo Erro_PercDesconto_Validate
'
'    'Vale o desconto que foi colocado aqui
'    objCT.giValorDescontoManual = 1
'
'    dPercDesconto = StrParaDbl(PercDesconto(objCT).Text)
'
'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
'        If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) > 0 Then
'            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    'Verifica se o Valor está preenchido
'    If dPercDesconto > 0 Then
'
'        'Faz a Crítica do Valor digitado
'        lErro = Porcentagem_Critica(PercDesconto(objCT).Text)
'        If lErro <> SUCESSO Then gError 132017
'
'    End If
'
'    objCT.ValorDesconto.Text = Format((StrParaDbl(objCT.ValorProdutos.Caption) + StrParaDbl(objCT.ValorDesconto)) * (dPercDesconto / 100), "STANDARD")
'
'    'Recalcula o Valor do desconto
'    lErro = objCT.ValorDesconto_Validate(bSGECancelDummy)
'    If lErro <> SUCESSO Then gError 132018
'
'    objCT.iValorDescontoAlterado = 0
'
'    Exit Sub
'
'Erro_PercDesconto_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 132017 To 132018
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179280)
'
'    End Select
'
'    Exit Sub
'
'End Sub'Public Sub PercDesconto_Change(ByVal objCT As Object)
'
'    objCT.iAlterado = REGISTRO_ALTERADO
'    objCT.iValorDescontoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercDesconto_Validate(ByVal objCT As Object, Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorProdutos As Double
'Dim iIndice As Integer
'Dim dPercDesconto As Double
'
'On Error GoTo Erro_PercDesconto_Validate
'
'    'Vale o desconto que foi colocado aqui
'    objCT.giValorDescontoManual = 1
'
'    dPercDesconto = StrParaDbl(PercDesconto(objCT).Text)
'
'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
'        If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) > 0 Then
'            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    'Verifica se o Valor está preenchido
'    If dPercDesconto > 0 Then
'
'        'Faz a Crítica do Valor digitado
'        lErro = Porcentagem_Critica(PercDesconto(objCT).Text)
'        If lErro <> SUCESSO Then gError 132017
'
'    End If
'
'    objCT.ValorDesconto.Text = Format((StrParaDbl(objCT.ValorProdutos.Caption) + StrParaDbl(objCT.ValorDesconto)) * (dPercDesconto / 100), "STANDARD")
'
'    'Recalcula o Valor do desconto
'    lErro = objCT.ValorDesconto_Validate(bSGECancelDummy)
'    If lErro <> SUCESSO Then gError 132018
'
'    objCT.iValorDescontoAlterado = 0
'
'    Exit Sub
'
'Erro_PercDesconto_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 132017 To 132018
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179280)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function NFCF_ValorDesconto_Validate(ByVal objCT As Object) As Long

'Dim lErro As Long
'Dim dValorDesconto As Double
'Dim dValorProdutos As Double
'
'On Error GoTo Erro_NFCF_ValorDesconto_Validate
'
'    dValorDesconto = StrParaDbl(objCT.ValorDesconto)
'    dValorProdutos = StrParaDbl(objCT.ValorProdutos.Caption)
'
'    If 0 > dValorProdutos Then gError 132016
'
'    If dValorDesconto + dValorProdutos <> 0 Then
'        PercDesconto(objCT).Text = (dValorDesconto / (dValorDesconto + dValorProdutos)) * 100
'    Else
'        PercDesconto(objCT).Text = "0"
'    End If
'
'    NFCF_ValorDesconto_Validate = SUCESSO
'
'    Exit Function
'
'Erro_NFCF_ValorDesconto_Validate:
'
'    NFCF_ValorDesconto_Validate = gErr
'
'    Select Case gErr
'
'        Case 132016
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos + dValorDesconto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179281)
'
'    End Select
'
'    Exit Function
    
End Function

Function NFCF_Default_DataSaida(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_Default_DataSaida

    NFCF_Default_DataSaida = SUCESSO

    Exit Function

Erro_NFCF_Default_DataSaida:

    NFCF_Default_DataSaida = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179282)

    End Select

    Exit Function
    
End Function

Function NFCF_DataSaida(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_DataSaida

    objNFiscal.dtDataSaida = DATA_NULA

    NFCF_DataSaida = SUCESSO
    
    Exit Function

Erro_NFCF_DataSaida:

    NFCF_DataSaida = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179283)

    End Select

    Exit Function
    
End Function

Function NFCF_Move_DataSaida_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_Move_DataSaida_Memoria
  
    If Len(Trim(objCT.DataSaida.ClipText)) > 0 Then
        objNFiscal.dtDataSaida = CDate(objCT.DataSaida.Text)
        objNFiscal.iSemDataSaida = 0
    
        If Len(Trim(objCT.HoraSaida.ClipText)) > 0 Then
            objNFiscal.dtHoraSaida = CDate(objCT.HoraSaida.Text)
        Else
            objNFiscal.dtHoraSaida = Time
        End If
    
    Else
        objNFiscal.dtDataSaida = DATA_NULA
        objNFiscal.iSemDataSaida = 1
    End If

  
    NFCF_Move_DataSaida_Memoria = SUCESSO

    Exit Function

Erro_NFCF_Move_DataSaida_Memoria:

    NFCF_Move_DataSaida_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179284)

    End Select

    Exit Function
    
End Function

Function NFCF_DataSaida2(ByVal objCT As Object, ByVal objCTNFiscalPedido As CTNFiscalPedido) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_DataSaida2

    objCTNFiscalPedido.DataSaida.PromptInclude = False
    objCTNFiscalPedido.DataSaida.Text = Format(DATA_NULA, "dd/mm/yy")
    objCTNFiscalPedido.DataSaida.PromptInclude = True

    NFCF_DataSaida2 = SUCESSO
    
    Exit Function

Erro_NFCF_DataSaida2:

    NFCF_DataSaida2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179285)

    End Select

    Exit Function
    
End Function
'###########################################################################
