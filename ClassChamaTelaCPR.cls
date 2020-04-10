VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassChamaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function BaixaRec_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBaixaRec As ClassBaixaReceber

On Error GoTo Erro_BaixaRec_Chama_Tela

    Load X.objBaixaRec
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45239
            
    If IsMissing(vP1) Then
    
        Set objBaixaRec = Nothing
        
    Else
    
        Set objBaixaRec = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBaixaRec.Trata_Parametros(objBaixaRec)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45240
            
        If lErro <> SUCESSO Then Error 45241
        
    End If
    
    X.objBaixaRec.Show
            
    BaixaRec_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaRec_Chama_Tela:

    BaixaRec_Chama_Tela = Err

    Select Case Err
    
        Case 45239, 45240
            Unload X.objBaixaRec
            
        Case 45241
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146676)
        
    End Select
        
    Exit Function

End Function

Function Paises_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPaises As ClassPais

On Error GoTo Erro_Paises_Chama_Tela

    Load X.objPaises
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45242
            
    If IsMissing(vP1) Then
    
        Set objPaises = Nothing
        
    Else
    
        Set objPaises = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPaises.Trata_Parametros(objPaises)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45243
            
        If lErro <> SUCESSO Then Error 45244
        
    End If
    
    X.objPaises.Show
            
    Paises_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Paises_Chama_Tela:

    Paises_Chama_Tela = Err

    Select Case Err
    
        Case 45242, 45243
            Unload X.objPaises
            
        Case 45244
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146677)
        
    End Select
        
    Exit Function

End Function

Function Clientes_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_Clientes_Chama_Tela

    Load X.objClientes
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9447
            
    If IsMissing(vP1) Then
    
        Set objCliente = Nothing
        
    Else
    
        Set objCliente = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objClientes.Trata_Parametros(objCliente)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9448
            
        If lErro <> SUCESSO Then Error 9449
        
    End If
    
    X.objClientes.Show
            
    Clientes_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Clientes_Chama_Tela:

    Clientes_Chama_Tela = Err

    Select Case Err
    
        Case 9447, 9448
            Unload X.objClientes
            
        Case 9449
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146678)
        
    End Select
        
    Exit Function

End Function

Function Bancos_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBanco As ClassBanco

On Error GoTo Erro_Bancos_Chama_Tela

    Load X.objBancos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9450
            
    If IsMissing(vP1) Then
    
        Set objBanco = Nothing
        
    Else
    
        Set objBanco = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBancos.Trata_Parametros(objBanco)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9451
            
        If lErro <> SUCESSO Then Error 9452
        
    End If
    
    X.objBancos.Show
            
    Bancos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Bancos_Chama_Tela:

    Bancos_Chama_Tela = Err

    Select Case Err
    
        Case 9450, 9451
            Unload X.objBancos
            
        Case 9452
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146679)
        
    End Select
        
    Exit Function

End Function

Function DebitosReceb_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objDebitosReceb As ClassDebitoRecCli

On Error GoTo Erro_DebitosReceb_Chama_Tela

    Load X.objDebitosReceb
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24678
            
    If IsMissing(vP1) Then
    
        Set objDebitosReceb = Nothing
        
    Else
    
        Set objDebitosReceb = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objDebitosReceb.Trata_Parametros(objDebitosReceb)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24679
            
        If lErro <> SUCESSO Then Error 24684
        
    End If
    
    X.objDebitosReceb.Show
            
    DebitosReceb_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DebitosReceb_Chama_Tela:

    DebitosReceb_Chama_Tela = Err

    Select Case Err
    
        Case 24678, 24679
            Unload X.objDebitosReceb
            
        Case 24684
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146680)
        
    End Select
        
    Exit Function

End Function

Function PagtoCancelar_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPagtoCancelar As ClassMovContaCorrente

On Error GoTo Erro_PagtoCancelar_Chama_Tela

    Load X.objPagtoCancelar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24685
            
    If IsMissing(vP1) Then
    
        Set objPagtoCancelar = Nothing
        
    Else
    
        Set objPagtoCancelar = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPagtoCancelar.Trata_Parametros(objPagtoCancelar)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24686
            
        If lErro <> SUCESSO Then Error 24687
        
    End If
    
    X.objPagtoCancelar.Show
            
    PagtoCancelar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PagtoCancelar_Chama_Tela:

    PagtoCancelar_Chama_Tela = Err

    Select Case Err
    
        Case 24685, 24686
            Unload X.objPagtoCancelar
            
        Case 24687
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146681)
        
    End Select
        
    Exit Function

End Function

Function RegiaoVenda_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRegiao As ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Chama_Tela

    Load X.objRegiaoVenda
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9554
            
    If IsMissing(vP1) Then
    
        Set objRegiao = Nothing
        
    Else
    
        Set objRegiao = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRegiaoVenda.Trata_Parametros(objRegiao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9555
            
        If lErro <> SUCESSO Then Error 9556
        
    End If
    
    X.objRegiaoVenda.Show
            
    RegiaoVenda_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RegiaoVenda_Chama_Tela:

    RegiaoVenda_Chama_Tela = Err

    Select Case Err
    
        Case 9554, 9555
            Unload X.objRegiaoVenda
            
        Case 9556
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146682)
        
    End Select
        
    Exit Function

End Function

Function TipoVendedor_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoVendedor As ClassTipoVendedor

On Error GoTo Erro_TipoVendedor_Chama_Tela

    Load X.objTipoVendedor
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14301
            
    If IsMissing(vP1) Then
    
        Set objTipoVendedor = Nothing
        
    Else
    
        Set objTipoVendedor = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoVendedor.Trata_Parametros(objTipoVendedor)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14302
            
        If lErro <> SUCESSO Then Error 14303
        
    End If
    
    X.objTipoVendedor.Show
            
    TipoVendedor_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoVendedor_Chama_Tela:

    TipoVendedor_Chama_Tela = Err

    Select Case Err
    
        Case 14301, 14302
            Unload X.objTipoVendedor
            
        Case 14303
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146683)
        
    End Select
        
    Exit Function

End Function

Function HistMovCta_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objHistMovCta1 As ClassHistMovCta

On Error GoTo Erro_HistMovCta_Chama_Tela

    Load X.objHistMovCta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9626
            
    If IsMissing(vP1) Then
    
        Set objHistMovCta1 = Nothing
        
    Else
    
        Set objHistMovCta1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objHistMovCta.Trata_Parametros(objHistMovCta1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9627
            
        If lErro <> SUCESSO Then Error 9628
        
    End If
    
    X.objHistMovCta.Show
            
    HistMovCta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_HistMovCta_Chama_Tela:

    HistMovCta_Chama_Tela = Err

    Select Case Err
    
        Case 9626, 9627
            Unload X.objHistMovCta
            
        Case 9628
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146684)
        
    End Select
        
    Exit Function

End Function

Function Favorecidos_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFavorecidos1 As ClassFavorecidos

On Error GoTo Erro_Favorecidos_Chama_Tela

    Load X.objFavorecidos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9630
            
    If IsMissing(vP1) Then
    
        Set objFavorecidos1 = Nothing
        
    Else
    
        Set objFavorecidos1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFavorecidos.Trata_Parametros(objFavorecidos1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9631
            
        If lErro <> SUCESSO Then Error 9632
        
    End If
    
    X.objFavorecidos.Show
            
    Favorecidos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Favorecidos_Chama_Tela:

    Favorecidos_Chama_Tela = Err

    Select Case Err
    
        Case 9630, 9631
            Unload X.objFavorecidos
            
        Case 9632
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146685)
        
    End Select
        
    Exit Function

End Function

Function Fornecedores_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Fornecedores_Chama_Tela

    Load X.objFornecedores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9786
            
    If IsMissing(vP1) Then
    
        Set objFornecedor = Nothing
        
    Else
    
        Set objFornecedor = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFornecedores.Trata_Parametros(objFornecedor)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9787
            
        If lErro <> SUCESSO Then Error 9788
        
    End If
    
    X.objFornecedores.Show
            
    Fornecedores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Fornecedores_Chama_Tela:

    Fornecedores_Chama_Tela = Err

    Select Case Err
    
        Case 9786, 9787
            Unload X.objFornecedores
            
        Case 9788
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146686)
        
    End Select
        
    Exit Function

End Function

Function FiliaisClientes_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFilialCliente As ClassFilialCliente

On Error GoTo Erro_FiliaisClientes_Chama_Tela

    Load X.objFiliaisClientes
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9790
            
    If IsMissing(vP1) Then
    
        Set objFilialCliente = Nothing
        
    Else
    
        Set objFilialCliente = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFiliaisClientes.Trata_Parametros(objFilialCliente)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9791
            
        If lErro <> SUCESSO Then Error 9792
        
    End If
    
    X.objFiliaisClientes.Show
            
    FiliaisClientes_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FiliaisClientes_Chama_Tela:

    FiliaisClientes_Chama_Tela = Err

    Select Case Err
    
        Case 9790, 9791
            Unload X.objFiliaisClientes
            
        Case 9792
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146687)
        
    End Select
        
    Exit Function

End Function

Function CreditosPagar_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCreditosPagar As ClassCreditoPagar

On Error GoTo Erro_CreditosPagar_Chama_Tela

    Load X.objCreditosPagar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14363
            
    If IsMissing(vP1) Then
    
        Set objCreditosPagar = Nothing
        
    Else
    
        Set objCreditosPagar = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCreditosPagar.Trata_Parametros(objCreditosPagar)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14364
            
        If lErro <> SUCESSO Then Error 14365
        
    End If
    
    X.objCreditosPagar.Show
            
    CreditosPagar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CreditosPagar_Chama_Tela:

    CreditosPagar_Chama_Tela = Err

    Select Case Err
    
        Case 14363, 14364
            Unload X.objCreditosPagar
            
        Case 14365
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146688)
        
    End Select
        
    Exit Function

End Function

Function TipoFornecedor_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoFornecedor As ClassTipoFornecedor

On Error GoTo Erro_TipoFornecedor_Chama_Tela

    Load X.objTipoFornecedor
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14370
            
    If IsMissing(vP1) Then
    
        Set objTipoFornecedor = Nothing
        
    Else
    
        Set objTipoFornecedor = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoFornecedor.Trata_Parametros(objTipoFornecedor)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14371
            
        If lErro <> SUCESSO Then Error 14372
        
    End If
    
    X.objTipoFornecedor.Show
            
    TipoFornecedor_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoFornecedor_Chama_Tela:

    TipoFornecedor_Chama_Tela = Err

    Select Case Err
    
        Case 14370, 14371
            Unload X.objTipoFornecedor
            
        Case 14372
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146689)
        
    End Select
        
    Exit Function

End Function

Function CondicoesPagto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCondicoesPagto As ClassCondicaoPagto

On Error GoTo Erro_CondicoesPagto_Chama_Tela

    Load X.objCondicoesPagto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14373
            
    If IsMissing(vP1) Then
    
        Set objCondicoesPagto = Nothing
        
    Else
    
        Set objCondicoesPagto = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCondicoesPagto.Trata_Parametros(objCondicoesPagto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14374
            
        If lErro <> SUCESSO Then Error 14375
        
    End If
    
    X.objCondicoesPagto.Show
            
    CondicoesPagto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CondicoesPagto_Chama_Tela:

    CondicoesPagto_Chama_Tela = Err

    Select Case Err
    
        Case 14373, 14374
            Unload X.objCondicoesPagto
            
        Case 14375
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146690)
        
    End Select
        
    Exit Function

End Function

Function Mensagens_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objMensagens As ClassMensagem

On Error GoTo Erro_Mensagens_Chama_Tela

    Load X.objMensagens
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14376
            
    If IsMissing(vP1) Then
    
        Set objMensagens = Nothing
        
    Else
    
        Set objMensagens = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objMensagens.Trata_Parametros(objMensagens)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14377
            
        If lErro <> SUCESSO Then Error 14378
        
    End If
    
    X.objMensagens.Show
            
    Mensagens_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Mensagens_Chama_Tela:

    Mensagens_Chama_Tela = Err

    Select Case Err
    
        Case 14376, 14377
            Unload X.objMensagens
            
        Case 14378
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146691)
        
    End Select
        
    Exit Function

End Function

Function FiliaisFornecedores_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFiliaisFornecedores As ClassFilialFornecedor

On Error GoTo Erro_FiliaisFornecedores_Chama_Tela

    Load X.objFiliaisFornecedores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14379
            
    If IsMissing(vP1) Then
    
        Set objFiliaisFornecedores = Nothing
        
    Else
    
        Set objFiliaisFornecedores = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFiliaisFornecedores.Trata_Parametros(objFiliaisFornecedores)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14380
            
        If lErro <> SUCESSO Then Error 14381
        
    End If
    
    X.objFiliaisFornecedores.Show
            
    FiliaisFornecedores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FiliaisFornecedores_Chama_Tela:

    FiliaisFornecedores_Chama_Tela = Err

    Select Case Err
    
        Case 14379, 14380
            Unload X.objFiliaisFornecedores
            
        Case 14381
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146692)
        
    End Select
        
    Exit Function

End Function

Function AntecipPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAntecipPag As ClassAntecipPag

On Error GoTo Erro_AntecipPag_Chama_Tela

    Load X.objAntecipPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14389
            
    If IsMissing(vP1) Then
    
        Set objAntecipPag = Nothing
        
    Else
    
        Set objAntecipPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAntecipPag.Trata_Parametros(objAntecipPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14390
            
        If lErro <> SUCESSO Then Error 14391
        
    End If
    
    X.objAntecipPag.Show
            
    AntecipPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AntecipPag_Chama_Tela:

    AntecipPag_Chama_Tela = Err

    Select Case Err
    
        Case 14389, 14390
            Unload X.objAntecipPag
            
        Case 14391
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146693)
        
    End Select
        
    Exit Function

End Function

Function BaixaPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBaixaPag As ClassBaixaPagar

On Error GoTo Erro_BaixaPag_Chama_Tela

    Load X.objBaixaPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14392
            
    If IsMissing(vP1) Then
    
        Set objBaixaPag = Nothing
        
    Else
    
        Set objBaixaPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBaixaPag.Trata_Parametros(objBaixaPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14393
            
        If lErro <> SUCESSO Then Error 14394
        
    End If
    
    X.objBaixaPag.Show
            
    BaixaPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaPag_Chama_Tela:

    BaixaPag_Chama_Tela = Err

    Select Case Err
    
        Case 14392, 14393
            Unload X.objBaixaPag
            
        Case 14394
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146694)
        
    End Select
        
    Exit Function

End Function

Function Cobradores_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCobradores As ClassCobrador

On Error GoTo Erro_Cobradores_Chama_Tela

    Load X.objCobradores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14398
            
    If IsMissing(vP1) Then
    
        Set objCobradores = Nothing
        
    Else
    
        Set objCobradores = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCobradores.Trata_Parametros(objCobradores)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14399
            
        If lErro <> SUCESSO Then Error 14400
        
    End If
    
    X.objCobradores.Show
            
    Cobradores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Cobradores_Chama_Tela:

    Cobradores_Chama_Tela = Err

    Select Case Err
    
        Case 14398, 14399
            Unload X.objCobradores
            
        Case 14400
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146695)
        
    End Select
        
    Exit Function

End Function

Function Vendedores_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objVendedores As ClassVendedor

On Error GoTo Erro_Vendedores_Chama_Tela

    Load X.objVendedores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14401
            
    If IsMissing(vP1) Then
    
        Set objVendedores = Nothing
        
    Else
    
        Set objVendedores = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objVendedores.Trata_Parametros(objVendedores)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14402
            
        If lErro <> SUCESSO Then Error 14403
        
    End If
    
    X.objVendedores.Show
            
    Vendedores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Vendedores_Chama_Tela:

    Vendedores_Chama_Tela = Err

    Select Case Err
    
        Case 14401, 14402
            Unload X.objVendedores
            
        Case 14403
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146696)
        
    End Select
        
    Exit Function

End Function

Function TipoCliente_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoCliente As ClassTipoCliente

On Error GoTo Erro_TipoCliente_Chama_Tela

    Load X.objTipoCliente
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14404
            
    If IsMissing(vP1) Then
    
        Set objTipoCliente = Nothing
        
    Else
    
        Set objTipoCliente = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoCliente.Trata_Parametros(objTipoCliente)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14405
            
        If lErro <> SUCESSO Then Error 14406
        
    End If
    
    X.objTipoCliente.Show
            
    TipoCliente_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoCliente_Chama_Tela:

    TipoCliente_Chama_Tela = Err

    Select Case Err
    
        Case 14404, 14405
            Unload X.objTipoCliente
            
        Case 14406
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146697)
        
    End Select
        
    Exit Function

End Function

Function TituloReceber_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_TituloReceber_Chama_Tela

    Load X.objTituloReceber
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14410
            
    If IsMissing(vP1) Then
    
        Set objTituloReceber = Nothing
        
    Else
    
        Set objTituloReceber = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTituloReceber.Trata_Parametros(objTituloReceber)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14411
            
        If lErro <> SUCESSO Then Error 14412
        
    End If
    
    X.objTituloReceber.Show
            
    TituloReceber_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TituloReceber_Chama_Tela:

    TituloReceber_Chama_Tela = Err

    Select Case Err
    
        Case 14410, 14411
            Unload X.objTituloReceber
            
        Case 14412
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146698)
        
    End Select
        
    Exit Function

End Function

Function PadroesCobranca_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPadroesCobranca As ClassPadraoCobranca

On Error GoTo Erro_PadroesCobranca_Chama_Tela

    Load X.objPadroesCobranca
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14413
            
    If IsMissing(vP1) Then
    
        Set objPadroesCobranca = Nothing
        
    Else
    
        Set objPadroesCobranca = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPadroesCobranca.Trata_Parametros(objPadroesCobranca)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14414
            
        If lErro <> SUCESSO Then Error 14415
        
    End If
    
    X.objPadroesCobranca.Show
            
    PadroesCobranca_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadroesCobranca_Chama_Tela:

    PadroesCobranca_Chama_Tela = Err

    Select Case Err
    
        Case 14413, 14414
            Unload X.objPadroesCobranca
            
        Case 14415
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146699)
        
    End Select
        
    Exit Function

End Function

Function ChequePre_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequePre As ClassChequePre

On Error GoTo Erro_ChequePre_Chama_Tela

    Load X.objChequePre
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14416
            
    If IsMissing(vP1) Then
    
        Set objChequePre = Nothing
        
    Else
    
        Set objChequePre = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequePre.Trata_Parametros(objChequePre)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14417
            
        If lErro <> SUCESSO Then Error 14418
        
    End If
    
    X.objChequePre.Show
            
    ChequePre_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePre_Chama_Tela:

    ChequePre_Chama_Tela = Err

    Select Case Err
    
        Case 14416, 14417
            Unload X.objChequePre
            
        Case 14418
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146700)
        
    End Select
        
    Exit Function

End Function

Function FaturasPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFaturasPag As ClassTituloPagar

On Error GoTo Erro_FaturasPag_Chama_Tela

    Load X.objFaturasPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19985
            
    If IsMissing(vP1) Then
    
        Set objFaturasPag = Nothing
        
    Else
    
        Set objFaturasPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFaturasPag.Trata_Parametros(objFaturasPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19986
            
        If lErro <> SUCESSO Then Error 19987
        
    End If
    
    X.objFaturasPag.Show
            
    FaturasPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FaturasPag_Chama_Tela:

    FaturasPag_Chama_Tela = Err

    Select Case Err
    
        Case 19985, 19986
            Unload X.objFaturasPag
            
        Case 19987
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146701)
        
    End Select
        
    Exit Function

End Function

Function CarteirasCobranca_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCarteirasCobranca As ClassCarteiraCobranca

On Error GoTo Erro_CarteirasCobranca_Chama_Tela

    Load X.objCarteirasCobranca
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19995
            
    If IsMissing(vP1) Then
    
        Set objCarteirasCobranca = Nothing
        
    Else
    
        Set objCarteirasCobranca = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCarteirasCobranca.Trata_Parametros(objCarteirasCobranca)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19996
            
        If lErro <> SUCESSO Then Error 19997
        
    End If
    
    X.objCarteirasCobranca.Show
            
    CarteirasCobranca_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CarteirasCobranca_Chama_Tela:

    CarteirasCobranca_Chama_Tela = Err

    Select Case Err
    
        Case 19995, 19996
            Unload X.objCarteirasCobranca
            
        Case 19997
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146702)
        
    End Select
        
    Exit Function

End Function

Function DetPag_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_DetPag_Chama_Tela

    Load X.objDetPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19988
    
    X.objDetPag.Show
            
    DetPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DetPag_Chama_Tela:

    DetPag_Chama_Tela = Err

    Select Case Err
    
        Case 19988
            Unload X.objDetPag
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146703)
        
    End Select
        
    Exit Function

End Function

Function BaixaRecCancelar_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BaixaRecCancelar_Chama_Tela

    Load X.objBaixaRecCancelar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45245
    
    X.objBaixaRecCancelar.Show
            
    BaixaRecCancelar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaRecCancelar_Chama_Tela:

    BaixaRecCancelar_Chama_Tela = Err

    Select Case Err
    
        Case 45245
            Unload X.objBaixaRecCancelar
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146704)
        
    End Select
        
    Exit Function

End Function

Function AlteracoesCobranca_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_AlteracoesCobranca_Chama_Tela

    Load X.objAlteracoesCobranca
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45246
    
    X.objAlteracoesCobranca.Show
            
    AlteracoesCobranca_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlteracoesCobranca_Chama_Tela:

    AlteracoesCobranca_Chama_Tela = Err

    Select Case Err
    
        Case 45246
            Unload X.objAlteracoesCobranca
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146705)
        
    End Select
        
    Exit Function

End Function

Function TransfCartCobr_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_TransfCartCobr_Chama_Tela

    Load X.objTransfCartCobr
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45247
    
    X.objTransfCartCobr.Show
            
    TransfCartCobr_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TransfCartCobr_Chama_Tela:

    TransfCartCobr_Chama_Tela = Err

    Select Case Err
    
        Case 45247
            Unload X.objTransfCartCobr
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146706)
        
    End Select
        
    Exit Function

End Function

Function BaixaPagCancelar_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BaixaPagCancelar_Chama_Tela

    Load X.objBaixaPagCancelar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19989
    
    X.objBaixaPagCancelar.Show
            
    BaixaPagCancelar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaPagCancelar_Chama_Tela:

    BaixaPagCancelar_Chama_Tela = Err

    Select Case Err
    
        Case 19989
            Unload X.objBaixaPagCancelar
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146707)
        
    End Select
        
    Exit Function

End Function

Function BaixaRecDig_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BaixaRecDig_Chama_Tela

    Load X.objBaixaRecDig
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45159
    
    X.objBaixaRecDig.Show
            
    BaixaRecDig_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaRecDig_Chama_Tela:

    BaixaRecDig_Chama_Tela = Err

    Select Case Err
    
        Case 45159
            Unload X.objBaixaRecDig
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146708)
        
    End Select
        
    Exit Function

End Function

Function BorderoCobrancaCancelar_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoCobrancaCancelar_Chama_Tela

    Load X.objBorderoCobrancaCancelar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45160
    
    X.objBorderoCobrancaCancelar.Show
            
    BorderoCobrancaCancelar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoCobrancaCancelar_Chama_Tela:

    BorderoCobrancaCancelar_Chama_Tela = Err

    Select Case Err
    
        Case 45160
            Unload X.objBorderoCobrancaCancelar
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146709)
        
    End Select
        
    Exit Function

End Function

Function AntecipReceb_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAntecipReceb As ClassAntecipReceb

On Error GoTo Erro_AntecipReceb_Chama_Tela

    Load X.objAntecipReceb
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14420
            
    If IsMissing(vP1) Then
    
        Set objAntecipReceb = Nothing
        
    Else
    
        Set objAntecipReceb = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAntecipReceb.Trata_Parametros(objAntecipReceb)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14421
            
        If lErro <> SUCESSO Then Error 14422
        
    End If
    
    X.objAntecipReceb.Show
            
    AntecipReceb_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AntecipReceb_Chama_Tela:

    AntecipReceb_Chama_Tela = Err

    Select Case Err
    
        Case 14420, 14421
            Unload X.objAntecipReceb
            
        Case 14422
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146710)
        
    End Select
        
    Exit Function

End Function

Function CtaCorrenteInt_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCCI As ClassContasCorrentesInternas

On Error GoTo Erro_CtaCorrenteInt_Chama_Tela

    Load X.objCtaCorrenteInt
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14423
            
    If IsMissing(vP1) Then
    
        Set objCCI = Nothing
        
    Else
    
        Set objCCI = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCtaCorrenteInt.Trata_Parametros(objCCI)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14424
            
        If lErro <> SUCESSO Then Error 14425
        
    End If
    
    X.objCtaCorrenteInt.Show
            
    CtaCorrenteInt_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CtaCorrenteInt_Chama_Tela:

    CtaCorrenteInt_Chama_Tela = Err

    Select Case Err
    
        Case 14423, 14424
            Unload X.objCtaCorrenteInt
            
        Case 14425
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146711)
        
    End Select
        
    Exit Function

End Function

Function Transportadora_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTransportadora As ClassTransportadora

On Error GoTo Erro_Transportadora_Chama_Tela

    Load X.objTransportadora
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24253
            
    If IsMissing(vP1) Then
    
        Set objTransportadora = Nothing
        
    Else
    
        Set objTransportadora = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTransportadora.Trata_Parametros(objTransportadora)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24254
            
        If lErro <> SUCESSO Then Error 24255
        
    End If
    
    X.objTransportadora.Show
            
    Transportadora_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Transportadora_Chama_Tela:

    Transportadora_Chama_Tela = Err

    Select Case Err
    
        Case 24253, 24254
            Unload X.objTransportadora
            
        Case 24255
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146712)
        
    End Select
        
    Exit Function

End Function

Function TipoAplicacao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoAplicacao As ClassTiposDeAplicacao

On Error GoTo Erro_TipoAplicacao_Chama_Tela

    Load X.objTipoAplicacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14435
            
    If IsMissing(vP1) Then
    
        Set objTipoAplicacao = Nothing
        
    Else
    
        Set objTipoAplicacao = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoAplicacao.Trata_Parametros(objTipoAplicacao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14436
            
        If lErro <> SUCESSO Then Error 14437
        
    End If
    
    X.objTipoAplicacao.Show
            
    TipoAplicacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoAplicacao_Chama_Tela:

    TipoAplicacao_Chama_Tela = Err

    Select Case Err
    
        Case 14435, 14436
            Unload X.objTipoAplicacao
            
        Case 14437
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146713)
        
    End Select
        
    Exit Function

End Function

Function Transferencia_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Transferencia_Chama_Tela

    Load X.objTransferencia
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14438
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTransferencia.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14439
            
        If lErro <> SUCESSO Then Error 14440
        
    End If
    
    X.objTransferencia.Show
            
    Transferencia_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Transferencia_Chama_Tela:

    Transferencia_Chama_Tela = Err

    Select Case Err
    
        Case 14438, 14439
            Unload X.objTransferencia
            
        Case 14440
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146714)
        
    End Select
        
    Exit Function

End Function

Function Aplicacao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAplicacao As ClassAplicacao

On Error GoTo Erro_Aplicacao_Chama_Tela

    Load X.objAplicacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14441
            
    If IsMissing(vP1) Then
    
        Set objAplicacao = Nothing
        
    Else
    
        Set objAplicacao = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAplicacao.Trata_Parametros(objAplicacao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14442
            
        If lErro <> SUCESSO Then Error 14443
        
    End If
    
    X.objAplicacao.Show
            
    Aplicacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Aplicacao_Chama_Tela:

    Aplicacao_Chama_Tela = Err

    Select Case Err
    
        Case 14441, 14442
            Unload X.objAplicacao
            
        Case 14443
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146715)
        
    End Select
        
    Exit Function

End Function

Function Resgate_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objResgate As ClassResgate

On Error GoTo Erro_Resgate_Chama_Tela

    Load X.objResgate
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14444
            
    If IsMissing(vP1) Then
    
        Set objResgate = Nothing
        
    Else
    
        Set objResgate = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objResgate.Trata_Parametros(objResgate)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14445
            
        If lErro <> SUCESSO Then Error 14446
        
    End If
    
    X.objResgate.Show
            
    Resgate_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Resgate_Chama_Tela:

    Resgate_Chama_Tela = Err

    Select Case Err
    
        Case 14444, 14445
            Unload X.objResgate
            
        Case 14446
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146716)
        
    End Select
        
    Exit Function

End Function

Function TipoAplicacaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoAplicacao1 As ClassTiposDeAplicacao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoAplicacaoLista_Chama_Tela

    Load X.objTipoAplicacaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14475
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTipoAplicacao1 = Nothing
        
    Else
    
        Set objTipoAplicacao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTipoAplicacaoLista.Trata_Parametros(colSelecao, objTipoAplicacao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14476
            
    If lErro <> SUCESSO Then Error 14477
        
    X.objTipoAplicacaoLista.Show
            
    TipoAplicacaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoAplicacaoLista_Chama_Tela:

    TipoAplicacaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 14475, 14476
            Unload X.objTipoAplicacaoLista
            
        Case 14477
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146717)
        
    End Select
        
    Exit Function

End Function

Function TipoClienteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoCliente1 As ClassTipoCliente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoClienteLista_Chama_Tela

    Load X.objTipoClienteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14495
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTipoCliente1 = Nothing
        
    Else
    
        Set objTipoCliente1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTipoClienteLista.Trata_Parametros(colSelecao, objTipoCliente1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14496
            
    If lErro <> SUCESSO Then Error 14497
        
    X.objTipoClienteLista.Show
            
    TipoClienteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoClienteLista_Chama_Tela:

    TipoClienteLista_Chama_Tela = Err

    Select Case Err
    
        Case 14495, 14496
            Unload X.objTipoClienteLista
            
        Case 14497
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146718)
        
    End Select
        
    Exit Function

End Function

Function BancoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objBanco1 As ClassBanco
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_BancoLista_Chama_Tela

    Load X.objBancoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14543
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objBanco1 = Nothing
        
    Else
    
        Set objBanco1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objBancoLista.Trata_Parametros(colSelecao, objBanco1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14544
            
    If lErro <> SUCESSO Then Error 14545
        
    X.objBancoLista.Show
            
    BancoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BancoLista_Chama_Tela:

    BancoLista_Chama_Tela = Err

    Select Case Err
    
        Case 14543, 14544
            Unload X.objBancoLista
            
        Case 14545
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146719)
        
    End Select
        
    Exit Function

End Function

Function PadraoCobrancaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPadraoCobranca1 As ClassPadraoCobranca
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PadraoCobrancaLista_Chama_Tela

    Load X.objPadraoCobrancaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24926
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPadraoCobranca1 = Nothing
        
    Else
    
        Set objPadraoCobranca1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPadraoCobrancaLista.Trata_Parametros(colSelecao, objPadraoCobranca1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24927
            
    If lErro <> SUCESSO Then Error 24928
        
    X.objPadraoCobrancaLista.Show
            
    PadraoCobrancaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadraoCobrancaLista_Chama_Tela:

    PadraoCobrancaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24926, 24927
            Unload X.objPadraoCobrancaLista
            
        Case 24928
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146720)
        
    End Select
        
    Exit Function

End Function

Function VendedorLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objVendedor1 As ClassVendedor
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_VendedorLista_Chama_Tela

    Load X.objVendedorLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14547
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objVendedor1 = Nothing
        
    Else
    
        Set objVendedor1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objVendedorLista.Trata_Parametros(colSelecao, objVendedor1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14548
            
    If lErro <> SUCESSO Then Error 14549
        
    X.objVendedorLista.Show
            
    VendedorLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_VendedorLista_Chama_Tela:

    VendedorLista_Chama_Tela = Err

    Select Case Err
    
        Case 14547, 14548
            Unload X.objVendedorLista
            
        Case 14549
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146721)
        
    End Select
        
    Exit Function

End Function

Function ExtratoBancarioCNAB2_Chama_Tela(X As Object, vP1 As Integer, vP2 As String) As Long

Dim lErro As Long
Dim iBanco As Integer
Dim sNomeArq As String

On Error GoTo Erro_ExtratoBancarioCNAB2_Chama_Tela

    Load X.objExtratoBancarioCNAB2
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45522
            
    iBanco = vP1
    sNomeArq = vP2
        
    lErro = X.objExtratoBancarioCNAB2.Trata_Parametros(iBanco, sNomeArq)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45523
            
    If lErro <> SUCESSO Then Error 45524
        
    X.objExtratoBancarioCNAB2.Show
            
    ExtratoBancarioCNAB2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExtratoBancarioCNAB2_Chama_Tela:

    ExtratoBancarioCNAB2_Chama_Tela = Err

    Select Case Err
    
        Case 45522, 45523
            Unload X.objExtratoBancarioCNAB2
            
        Case 45524
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146722)
        
    End Select
        
    Exit Function

End Function

Function TransportadoraLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTransportadora1 As ClassTransportadora
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TransportadoraLista_Chama_Tela

    Load X.objTransportadoraLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29547
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTransportadora1 = Nothing
        
    Else
    
        Set objTransportadora1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTransportadoraLista.Trata_Parametros(colSelecao, objTransportadora1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29548
            
    If lErro <> SUCESSO Then Error 29549
        
    X.objTransportadoraLista.Show
            
    TransportadoraLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TransportadoraLista_Chama_Tela:

    TransportadoraLista_Chama_Tela = Err

    Select Case Err
    
        Case 29547, 29548
            Unload X.objTransportadoraLista
            
        Case 29549
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146723)
        
    End Select
        
    Exit Function

End Function

Function CtaCorrenteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCtaCorrente1 As ClassContasCorrentesInternas
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CtaCorrenteLista_Chama_Tela

    Load X.objCtaCorrenteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14553
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCtaCorrente1 = Nothing
        
    Else
    
        Set objCtaCorrente1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCtaCorrenteLista.Trata_Parametros(colSelecao, objCtaCorrente1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14554
            
    If lErro <> SUCESSO Then Error 14555
        
    X.objCtaCorrenteLista.Show
            
    CtaCorrenteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CtaCorrenteLista_Chama_Tela:

    CtaCorrenteLista_Chama_Tela = Err

    Select Case Err
    
        Case 14553, 14554
            Unload X.objCtaCorrenteLista
            
        Case 14555
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146724)
        
    End Select
        
    Exit Function

End Function

Function CtaCorrenteTodasLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCtaCorrente1 As ClassContasCorrentesInternas
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CtaCorrenteTodasLista_Chama_Tela

    Load X.objCtaCorrenteTodasLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19786
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCtaCorrente1 = Nothing
        
    Else
    
        Set objCtaCorrente1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCtaCorrenteTodasLista.Trata_Parametros(colSelecao, objCtaCorrente1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19787
            
    If lErro <> SUCESSO Then Error 19788
        
    X.objCtaCorrenteTodasLista.Show
            
    CtaCorrenteTodasLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CtaCorrenteTodasLista_Chama_Tela:

    CtaCorrenteTodasLista_Chama_Tela = Err

    Select Case Err
    
        Case 19786, 19787
            Unload X.objCtaCorrenteTodasLista
            
        Case 19788
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146725)
        
    End Select
        
    Exit Function

End Function

Function SaqueLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objSaque1 As ClassMovContaCorrente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_SaqueLista_Chama_Tela

    Load X.objSaqueLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14574
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objSaque1 = Nothing
        
    Else
    
        Set objSaque1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objSaqueLista.Trata_Parametros(colSelecao, objSaque1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14575
            
    If lErro <> SUCESSO Then Error 14576
        
    X.objSaqueLista.Show
            
    SaqueLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SaqueLista_Chama_Tela:

    SaqueLista_Chama_Tela = Err

    Select Case Err
    
        Case 14574, 14575
            Unload X.objSaqueLista
            
        Case 14576
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146726)
        
    End Select
        
    Exit Function

End Function

Function ResgateLista_Aplicacao_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objResgate1 As ClassResgate
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ResgateLista_Aplicacao_Chama_Tela

    Load X.objResgateLista_Aplicacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29789
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objResgate1 = Nothing
        
    Else
    
        Set objResgate1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objResgateLista_Aplicacao.Trata_Parametros(colSelecao, objResgate1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29790
            
    If lErro <> SUCESSO Then Error 29791
        
    X.objResgateLista_Aplicacao.Show
            
    ResgateLista_Aplicacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ResgateLista_Aplicacao_Chama_Tela:

    ResgateLista_Aplicacao_Chama_Tela = Err

    Select Case Err
    
        Case 29789, 29790
            Unload X.objResgateLista_Aplicacao
            
        Case 29791
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146727)
        
    End Select
        
    Exit Function

End Function

Function PagtoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPagto1 As ClassMovContaCorrente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PagtoLista_Chama_Tela

    Load X.objPagtoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24184
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPagto1 = Nothing
        
    Else
    
        Set objPagto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPagtoLista.Trata_Parametros(colSelecao, objPagto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24185
            
    If lErro <> SUCESSO Then Error 24186
        
    X.objPagtoLista.Show
            
    PagtoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PagtoLista_Chama_Tela:

    PagtoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24184, 24185
            Unload X.objPagtoLista
            
        Case 24186
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146728)
        
    End Select
        
    Exit Function

End Function

Function PagtoLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPagto1 As ClassMovContaCorrente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PagtoLista_Consulta_Chama_Tela

    Load X.objPagtoLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24196
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPagto1 = Nothing
        
    Else
    
        Set objPagto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPagtoLista_Consulta.Trata_Parametros(colSelecao, objPagto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24197
            
    If lErro <> SUCESSO Then Error 24198
        
    X.objPagtoLista_Consulta.Show
            
    PagtoLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PagtoLista_Consulta_Chama_Tela:

    PagtoLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24196, 24197
            Unload X.objPagtoLista_Consulta
            
        Case 24198
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146729)
        
    End Select
        
    Exit Function

End Function

Function DepositoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objDeposito1 As ClassMovContaCorrente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_DepositoLista_Chama_Tela

    Load X.objDepositoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14577
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objDeposito1 = Nothing
        
    Else
    
        Set objDeposito1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objDepositoLista.Trata_Parametros(colSelecao, objDeposito1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14578
            
    If lErro <> SUCESSO Then Error 14579
        
    X.objDepositoLista.Show
            
    DepositoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DepositoLista_Chama_Tela:

    DepositoLista_Chama_Tela = Err

    Select Case Err
    
        Case 14577, 14578
            Unload X.objDepositoLista
            
        Case 14579
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146730)
        
    End Select
        
    Exit Function

End Function

Function DepositoLista_Conta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objDeposito1 As ClassMovContaCorrente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_DepositoLista_Conta_Chama_Tela

    Load X.objDepositoLista_Conta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14874
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objDeposito1 = Nothing
        
    Else
    
        Set objDeposito1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objDepositoLista_Conta.Trata_Parametros(colSelecao, objDeposito1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14875
            
    If lErro <> SUCESSO Then Error 14876
        
    X.objDepositoLista_Conta.Show
            
    DepositoLista_Conta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DepositoLista_Conta_Chama_Tela:

    DepositoLista_Conta_Chama_Tela = Err

    Select Case Err
    
        Case 14874, 14875
            Unload X.objDepositoLista_Conta
            
        Case 14876
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146731)
        
    End Select
        
    Exit Function

End Function

Function SaqueLista_Conta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objSaque1 As ClassMovContaCorrente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_SaqueLista_Conta_Chama_Tela

    Load X.objSaqueLista_Conta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14877
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objSaque1 = Nothing
        
    Else
    
        Set objSaque1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objSaqueLista_Conta.Trata_Parametros(colSelecao, objSaque1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14878
            
    If lErro <> SUCESSO Then Error 14879
        
    X.objSaqueLista_Conta.Show
            
    SaqueLista_Conta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SaqueLista_Conta_Chama_Tela:

    SaqueLista_Conta_Chama_Tela = Err

    Select Case Err
    
        Case 14877, 14878
            Unload X.objSaqueLista_Conta
            
        Case 14879
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146732)
        
    End Select
        
    Exit Function

End Function

Function FornecedorLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objFornecedor1 As ClassFornecedor
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_FornecedorLista_Chama_Tela

    Load X.objFornecedorLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14604
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objFornecedor1 = Nothing
        
    Else
    
        Set objFornecedor1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objFornecedorLista.Trata_Parametros(colSelecao, objFornecedor1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14605
            
    If lErro <> SUCESSO Then Error 14606
        
    X.objFornecedorLista.Show
            
    FornecedorLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FornecedorLista_Chama_Tela:

    FornecedorLista_Chama_Tela = Err

    Select Case Err
    
        Case 14604, 14605
            Unload X.objFornecedorLista
            
        Case 14606
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146733)
        
    End Select
        
    Exit Function

End Function

Function TipoFornecedorLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoFornecedor1 As ClassTipoFornecedor
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoFornecedorLista_Chama_Tela

    Load X.objTipoFornecedorLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14608
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTipoFornecedor1 = Nothing
        
    Else
    
        Set objTipoFornecedor1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTipoFornecedorLista.Trata_Parametros(colSelecao, objTipoFornecedor1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14609
            
    If lErro <> SUCESSO Then Error 14610
        
    X.objTipoFornecedorLista.Show
            
    TipoFornecedorLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoFornecedorLista_Chama_Tela:

    TipoFornecedorLista_Chama_Tela = Err

    Select Case Err
    
        Case 14608, 14609
            Unload X.objTipoFornecedorLista
            
        Case 14610
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146734)
        
    End Select
        
    Exit Function

End Function

Function PlanoContaCRLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPlanoContaCR1 As ClassPlanoConta
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PlanoContaCRLista_Chama_Tela

    Load X.objPlanoContaCRLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14689
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPlanoContaCR1 = Nothing
        
    Else
    
        Set objPlanoContaCR1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPlanoContaCRLista.Trata_Parametros(colSelecao, objPlanoContaCR1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14690
            
    If lErro <> SUCESSO Then Error 14691
        
    X.objPlanoContaCRLista.Show
            
    PlanoContaCRLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PlanoContaCRLista_Chama_Tela:

    PlanoContaCRLista_Chama_Tela = Err

    Select Case Err
    
        Case 14689, 14690
            Unload X.objPlanoContaCRLista
            
        Case 14691
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146735)
        
    End Select
        
    Exit Function

End Function

Function PlanoContaCPLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPlanoContaCP1 As ClassPlanoConta
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PlanoContaCPLista_Chama_Tela

    Load X.objPlanoContaCPLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14692
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPlanoContaCP1 = Nothing
        
    Else
    
        Set objPlanoContaCP1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPlanoContaCPLista.Trata_Parametros(colSelecao, objPlanoContaCP1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14693
            
    If lErro <> SUCESSO Then Error 14694
        
    X.objPlanoContaCPLista.Show
            
    PlanoContaCPLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PlanoContaCPLista_Chama_Tela:

    PlanoContaCPLista_Chama_Tela = Err

    Select Case Err
    
        Case 14692, 14693
            Unload X.objPlanoContaCPLista
            
        Case 14694
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146736)
        
    End Select
        
    Exit Function

End Function

Function CondicaoPagtoCRLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCondicaoPagtoCR1 As ClassCondicaoPagto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CondicaoPagtoCRLista_Chama_Tela

    Load X.objCondicaoPagtoCRLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14711
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCondicaoPagtoCR1 = Nothing
        
    Else
    
        Set objCondicaoPagtoCR1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCondicaoPagtoCRLista.Trata_Parametros(colSelecao, objCondicaoPagtoCR1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14712
            
    If lErro <> SUCESSO Then Error 14713
        
    X.objCondicaoPagtoCRLista.Show
            
    CondicaoPagtoCRLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CondicaoPagtoCRLista_Chama_Tela:

    CondicaoPagtoCRLista_Chama_Tela = Err

    Select Case Err
    
        Case 14711, 14712
            Unload X.objCondicaoPagtoCRLista
            
        Case 14713
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146737)
        
    End Select
        
    Exit Function

End Function

Function CondicaoPagtoCPLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCondicaoPagtoCP1 As ClassCondicaoPagto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CondicaoPagtoCPLista_Chama_Tela

    Load X.objCondicaoPagtoCPLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14714
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCondicaoPagtoCP1 = Nothing
        
    Else
    
        Set objCondicaoPagtoCP1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCondicaoPagtoCPLista.Trata_Parametros(colSelecao, objCondicaoPagtoCP1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14715
            
    If lErro <> SUCESSO Then Error 14716
        
    X.objCondicaoPagtoCPLista.Show
            
    CondicaoPagtoCPLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CondicaoPagtoCPLista_Chama_Tela:

    CondicaoPagtoCPLista_Chama_Tela = Err

    Select Case Err
    
        Case 14714, 14715
            Unload X.objCondicaoPagtoCPLista
            
        Case 14716
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146738)
        
    End Select
        
    Exit Function

End Function

Function ClientesLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objClientes1 As ClassCliente
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ClientesLista_Chama_Tela

    Load X.objClientesLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14755
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objClientes1 = Nothing
        
    Else
    
        Set objClientes1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objClientesLista.Trata_Parametros(colSelecao, objClientes1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14756
            
    If lErro <> SUCESSO Then Error 14757
        
    X.objClientesLista.Show
            
    ClientesLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ClientesLista_Chama_Tela:

    ClientesLista_Chama_Tela = Err

    Select Case Err
    
        Case 14755, 14756
            Unload X.objClientesLista
            
        Case 14757
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146739)
        
    End Select
        
    Exit Function

End Function

Function NFPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFPag1 As ClassNFsPag
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFPagLista_Chama_Tela

    Load X.objNFPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14758
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objNFPag1 = Nothing
        
    Else
    
        Set objNFPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objNFPagLista.Trata_Parametros(colSelecao, objNFPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14759
            
    If lErro <> SUCESSO Then Error 14760
        
    X.objNFPagLista.Show
            
    NFPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFPagLista_Chama_Tela:

    NFPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 14758, 14759
            Unload X.objNFPagLista
            
        Case 14760
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146740)
        
    End Select
        
    Exit Function

End Function

Function OutrosPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOutrosPag1 As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OutrosPagLista_Chama_Tela

    Load X.objOutrosPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24036
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objOutrosPag1 = Nothing
        
    Else
    
        Set objOutrosPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objOutrosPagLista.Trata_Parametros(colSelecao, objOutrosPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24037
            
    If lErro <> SUCESSO Then Error 24038
        
    X.objOutrosPagLista.Show
            
    OutrosPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OutrosPagLista_Chama_Tela:

    OutrosPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 24036, 24037
            Unload X.objOutrosPagLista
            
        Case 24038
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146741)
        
    End Select
        
    Exit Function

End Function

Function OutrosPagLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOutrosPag1 As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OutrosPagLista_Consulta_Chama_Tela

    Load X.objOutrosPagLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24076
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objOutrosPag1 = Nothing
        
    Else
    
        Set objOutrosPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objOutrosPagLista_Consulta.Trata_Parametros(colSelecao, objOutrosPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24077
            
    If lErro <> SUCESSO Then Error 24078
        
    X.objOutrosPagLista_Consulta.Show
            
    OutrosPagLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OutrosPagLista_Consulta_Chama_Tela:

    OutrosPagLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24076, 24077
            Unload X.objOutrosPagLista_Consulta
            
        Case 24078
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146742)
        
    End Select
        
    Exit Function

End Function

Function CreditosPagarLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCreditosPagar1 As ClassCreditoPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CreditosPagarLista_Chama_Tela

    Load X.objCreditosPagarLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24056
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCreditosPagar1 = Nothing
        
    Else
    
        Set objCreditosPagar1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCreditosPagarLista.Trata_Parametros(colSelecao, objCreditosPagar1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24057
            
    If lErro <> SUCESSO Then Error 24058
        
    X.objCreditosPagarLista.Show
            
    CreditosPagarLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CreditosPagarLista_Chama_Tela:

    CreditosPagarLista_Chama_Tela = Err

    Select Case Err
    
        Case 24056, 24057
            Unload X.objCreditosPagarLista
            
        Case 24058
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146743)
        
    End Select
        
    Exit Function

End Function

Function CarteirasCobrancaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCarteirasCobranca1 As ClassCarteiraCobranca
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CarteirasCobrancaLista_Chama_Tela

    Load X.objCarteirasCobrancaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24145
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCarteirasCobranca1 = Nothing
        
    Else
    
        Set objCarteirasCobranca1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCarteirasCobrancaLista.Trata_Parametros(colSelecao, objCarteirasCobranca1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24146
            
    If lErro <> SUCESSO Then Error 24147
        
    X.objCarteirasCobrancaLista.Show
            
    CarteirasCobrancaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CarteirasCobrancaLista_Chama_Tela:

    CarteirasCobrancaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24145, 24146
            Unload X.objCarteirasCobrancaLista
            
        Case 24147
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146744)
        
    End Select
        
    Exit Function

End Function

Function BorderoCobrancaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objBorderoCobranca1 As ClassBorderoCobranca
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_BorderoCobrancaLista_Chama_Tela

    Load X.objBorderoCobrancaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45170
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objBorderoCobranca1 = Nothing
        
    Else
    
        Set objBorderoCobranca1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objBorderoCobrancaLista.Trata_Parametros(colSelecao, objBorderoCobranca1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45171
            
    If lErro <> SUCESSO Then Error 45172
        
    X.objBorderoCobrancaLista.Show
            
    BorderoCobrancaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoCobrancaLista_Chama_Tela:

    BorderoCobrancaLista_Chama_Tela = Err

    Select Case Err
    
        Case 45170, 45171
            Unload X.objBorderoCobrancaLista
            
        Case 45172
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146745)
        
    End Select
        
    Exit Function

End Function

Function TitulosPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTitulosPag As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TitulosPagLista_Chama_Tela

    Load X.objTitulosPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45303
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTitulosPag = Nothing
        
    Else
    
        Set objTitulosPag = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTitulosPagLista.Trata_Parametros(colSelecao, objTitulosPag, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45304
            
    If lErro <> SUCESSO Then Error 45305
        
    X.objTitulosPagLista.Show
            
    TitulosPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TitulosPagLista_Chama_Tela:

    TitulosPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 45303, 45304
            Unload X.objTitulosPagLista
            
        Case 45305
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146746)
        
    End Select
        
    Exit Function

End Function

Function TitulosPagComBaixaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTitulosPag As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TitulosPagComBaixaLista_Chama_Tela

    Load X.objTitulosPagComBaixaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45479
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTitulosPag = Nothing
        
    Else
    
        Set objTitulosPag = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTitulosPagComBaixaLista.Trata_Parametros(colSelecao, objTitulosPag, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45480
            
    If lErro <> SUCESSO Then Error 45481
        
    X.objTitulosPagComBaixaLista.Show
            
    TitulosPagComBaixaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TitulosPagComBaixaLista_Chama_Tela:

    TitulosPagComBaixaLista_Chama_Tela = Err

    Select Case Err
    
        Case 45479, 45480
            Unload X.objTitulosPagComBaixaLista
            
        Case 45481
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146747)
        
    End Select
        
    Exit Function

End Function

Function BaixasParcPagSeqLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objBaixaParcPagar As ClassBaixaParcPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_BaixasParcPagSeqLista_Chama_Tela

    Load X.objBaixasParcPagSeqLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45491
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objBaixaParcPagar = Nothing
        
    Else
    
        Set objBaixaParcPagar = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objBaixasParcPagSeqLista.Trata_Parametros(colSelecao, objBaixaParcPagar, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45492
            
    If lErro <> SUCESSO Then Error 45493
        
    X.objBaixasParcPagSeqLista.Show
            
    BaixasParcPagSeqLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixasParcPagSeqLista_Chama_Tela:

    BaixasParcPagSeqLista_Chama_Tela = Err

    Select Case Err
    
        Case 45491, 45492
            Unload X.objBaixasParcPagSeqLista
            
        Case 45493
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146748)
        
    End Select
        
    Exit Function

End Function

Function BaixasParcRecSeqLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objBaixaParcRec As ClassBaixaParcRec
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_BaixasParcRecSeqLista_Chama_Tela

    Load X.objBaixasParcRecSeqLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45516
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objBaixaParcRec = Nothing
        
    Else
    
        Set objBaixaParcRec = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objBaixasParcRecSeqLista.Trata_Parametros(colSelecao, objBaixaParcRec, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45517
            
    If lErro <> SUCESSO Then Error 45518
        
    X.objBaixasParcRecSeqLista.Show
            
    BaixasParcRecSeqLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixasParcRecSeqLista_Chama_Tela:

    BaixasParcRecSeqLista_Chama_Tela = Err

    Select Case Err
    
        Case 45516, 45517
            Unload X.objBaixasParcRecSeqLista
            
        Case 45518
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146749)
        
    End Select
        
    Exit Function

End Function

Function CarteirasCobradorLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCarteirasCobrador1 As ClassCarteiraCobrador
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CarteirasCobradorLista_Chama_Tela

    Load X.objCarteirasCobradorLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24265
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCarteirasCobrador1 = Nothing
        
    Else
    
        Set objCarteirasCobrador1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCarteirasCobradorLista.Trata_Parametros(colSelecao, objCarteirasCobrador1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24266
            
    If lErro <> SUCESSO Then Error 24267
        
    X.objCarteirasCobradorLista.Show
            
    CarteirasCobradorLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CarteirasCobradorLista_Chama_Tela:

    CarteirasCobradorLista_Chama_Tela = Err

    Select Case Err
    
        Case 24265, 24266
            Unload X.objCarteirasCobradorLista
            
        Case 24267
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146750)
        
    End Select
        
    Exit Function

End Function

Function CreditosPagarLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCreditosPagar1 As ClassCreditoPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CreditosPagarLista_Consulta_Chama_Tela

    Load X.objCreditosPagarLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24109
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCreditosPagar1 = Nothing
        
    Else
    
        Set objCreditosPagar1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCreditosPagarLista_Consulta.Trata_Parametros(colSelecao, objCreditosPagar1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24110
            
    If lErro <> SUCESSO Then Error 24111
        
    X.objCreditosPagarLista_Consulta.Show
            
    CreditosPagarLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CreditosPagarLista_Consulta_Chama_Tela:

    CreditosPagarLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24109, 24110
            Unload X.objCreditosPagarLista_Consulta
            
        Case 24111
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146751)
        
    End Select
        
    Exit Function

End Function

Function ChequePreLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objChequePre1 As ClassChequePre
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ChequePreLista_Chama_Tela

    Load X.objChequePreLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24221
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objChequePre1 = Nothing
        
    Else
    
        Set objChequePre1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objChequePreLista.Trata_Parametros(colSelecao, objChequePre1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24222
            
    If lErro <> SUCESSO Then Error 24223
        
    X.objChequePreLista.Show
            
    ChequePreLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePreLista_Chama_Tela:

    ChequePreLista_Chama_Tela = Err

    Select Case Err
    
        Case 24221, 24222
            Unload X.objChequePreLista
            
        Case 24223
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146752)
        
    End Select
        
    Exit Function

End Function

Function ChequePreLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objChequePre1 As ClassChequePre
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ChequePreLista_Consulta_Chama_Tela

    Load X.objChequePreLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24224
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objChequePre1 = Nothing
        
    Else
    
        Set objChequePre1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objChequePreLista_Consulta.Trata_Parametros(colSelecao, objChequePre1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24225
            
    If lErro <> SUCESSO Then Error 24226
        
    X.objChequePreLista_Consulta.Show
            
    ChequePreLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePreLista_Consulta_Chama_Tela:

    ChequePreLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24224, 24225
            Unload X.objChequePreLista_Consulta
            
        Case 24226
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146753)
        
    End Select
        
    Exit Function

End Function

Function DebitosRecebLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objDebitosReceb1 As ClassDebitoRecCli
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_DebitosRecebLista_Chama_Tela

    Load X.objDebitosRecebLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24064
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objDebitosReceb1 = Nothing
        
    Else
    
        Set objDebitosReceb1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objDebitosRecebLista.Trata_Parametros(colSelecao, objDebitosReceb1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24065
            
    If lErro <> SUCESSO Then Error 24066
        
    X.objDebitosRecebLista.Show
            
    DebitosRecebLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DebitosRecebLista_Chama_Tela:

    DebitosRecebLista_Chama_Tela = Err

    Select Case Err
    
        Case 24064, 24065
            Unload X.objDebitosRecebLista
            
        Case 24066
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146754)
        
    End Select
        
    Exit Function

End Function

Function DebitosRecebLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objDebitosReceb1 As ClassDebitoRecCli
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_DebitosRecebLista_Consulta_Chama_Tela

    Load X.objDebitosRecebLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24112
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objDebitosReceb1 = Nothing
        
    Else
    
        Set objDebitosReceb1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objDebitosRecebLista_Consulta.Trata_Parametros(colSelecao, objDebitosReceb1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24113
            
    If lErro <> SUCESSO Then Error 24114
        
    X.objDebitosRecebLista_Consulta.Show
            
    DebitosRecebLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DebitosRecebLista_Consulta_Chama_Tela:

    DebitosRecebLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24112, 24113
            Unload X.objDebitosRecebLista_Consulta
            
        Case 24114
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146755)
        
    End Select
        
    Exit Function

End Function

Function NFPagLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFPag1 As ClassNFsPag
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFPagLista_Consulta_Chama_Tela

    Load X.objNFPagLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14917
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objNFPag1 = Nothing
        
    Else
    
        Set objNFPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objNFPagLista_Consulta.Trata_Parametros(colSelecao, objNFPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14918
            
    If lErro <> SUCESSO Then Error 14919
        
    X.objNFPagLista_Consulta.Show
            
    NFPagLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFPagLista_Consulta_Chama_Tela:

    NFPagLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 14917, 14918
            Unload X.objNFPagLista_Consulta
            
        Case 14919
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146756)
        
    End Select
        
    Exit Function

End Function

Function AntecipPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAntecipPag1 As ClassAntecipPag
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AntecipPagLista_Chama_Tela

    Load X.objAntecipPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14890
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objAntecipPag1 = Nothing
        
    Else
    
        Set objAntecipPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objAntecipPagLista.Trata_Parametros(colSelecao, objAntecipPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14891
            
    If lErro <> SUCESSO Then Error 14892
        
    X.objAntecipPagLista.Show
            
    AntecipPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AntecipPagLista_Chama_Tela:

    AntecipPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 14890, 14891
            Unload X.objAntecipPagLista
            
        Case 14892
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146757)
        
    End Select
        
    Exit Function

End Function

Function AntecipPagLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAntecipPag1 As ClassAntecipPag
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AntecipPagLista_Consulta_Chama_Tela

    Load X.objAntecipPagLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24770
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objAntecipPag1 = Nothing
        
    Else
    
        Set objAntecipPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objAntecipPagLista_Consulta.Trata_Parametros(colSelecao, objAntecipPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24771
            
    If lErro <> SUCESSO Then Error 24772
        
    X.objAntecipPagLista_Consulta.Show
            
    AntecipPagLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AntecipPagLista_Consulta_Chama_Tela:

    AntecipPagLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24770, 24771
            Unload X.objAntecipPagLista_Consulta
            
        Case 24772
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146758)
        
    End Select
        
    Exit Function

End Function

Function AntecipRecebLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAntecipReceb1 As ClassAntecipReceb
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AntecipRecebLista_Chama_Tela

    Load X.objAntecipRecebLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14904
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objAntecipReceb1 = Nothing
        
    Else
    
        Set objAntecipReceb1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objAntecipRecebLista.Trata_Parametros(colSelecao, objAntecipReceb1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14905
            
    If lErro <> SUCESSO Then Error 14906
        
    X.objAntecipRecebLista.Show
            
    AntecipRecebLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AntecipRecebLista_Chama_Tela:

    AntecipRecebLista_Chama_Tela = Err

    Select Case Err
    
        Case 14904, 14905
            Unload X.objAntecipRecebLista
            
        Case 14906
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146759)
        
    End Select
        
    Exit Function

End Function

Function AntecipRecebLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAntecipReceb1 As ClassAntecipReceb
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AntecipRecebLista_Consulta_Chama_Tela

    Load X.objAntecipRecebLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24773
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objAntecipReceb1 = Nothing
        
    Else
    
        Set objAntecipReceb1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objAntecipRecebLista_Consulta.Trata_Parametros(colSelecao, objAntecipReceb1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24774
            
    If lErro <> SUCESSO Then Error 24775
        
    X.objAntecipRecebLista_Consulta.Show
            
    AntecipRecebLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AntecipRecebLista_Consulta_Chama_Tela:

    AntecipRecebLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24773, 24774
            Unload X.objAntecipRecebLista_Consulta
            
        Case 24775
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146760)
        
    End Select
        
    Exit Function

End Function

Function AplicacaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAplicacao1 As ClassAplicacao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AplicacaoLista_Chama_Tela

    Load X.objAplicacaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24332
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objAplicacao1 = Nothing
        
    Else
    
        Set objAplicacao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objAplicacaoLista.Trata_Parametros(colSelecao, objAplicacao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24333
            
    If lErro <> SUCESSO Then Error 24334
        
    X.objAplicacaoLista.Show
            
    AplicacaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AplicacaoLista_Chama_Tela:

    AplicacaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24332, 24333
            Unload X.objAplicacaoLista
            
        Case 24334
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146761)
        
    End Select
        
    Exit Function

End Function

Function NFFatPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFFatPag1 As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFFatPagLista_Chama_Tela

    Load X.objNFFatPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14930
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objNFFatPag1 = Nothing
        
    Else
    
        Set objNFFatPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objNFFatPagLista.Trata_Parametros(colSelecao, objNFFatPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14931
            
    If lErro <> SUCESSO Then Error 14932
        
    X.objNFFatPagLista.Show
            
    NFFatPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFFatPagLista_Chama_Tela:

    NFFatPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 14930, 14931
            Unload X.objNFFatPagLista
            
        Case 14932
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146762)
        
    End Select
        
    Exit Function

End Function

Function NFFatPagLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFFatPag1 As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFFatPagLista_Consulta_Chama_Tela

    Load X.objNFFatPagLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24106
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objNFFatPag1 = Nothing
        
    Else
    
        Set objNFFatPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objNFFatPagLista_Consulta.Trata_Parametros(colSelecao, objNFFatPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24107
            
    If lErro <> SUCESSO Then Error 24108
        
    X.objNFFatPagLista_Consulta.Show
            
    NFFatPagLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFFatPagLista_Consulta_Chama_Tela:

    NFFatPagLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24106, 24107
            Unload X.objNFFatPagLista_Consulta
            
        Case 24108
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146763)
        
    End Select
        
    Exit Function

End Function

Function FaturasPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objFaturasPag1 As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_FaturasPagLista_Chama_Tela

    Load X.objFaturasPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14994
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objFaturasPag1 = Nothing
        
    Else
    
        Set objFaturasPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objFaturasPagLista.Trata_Parametros(colSelecao, objFaturasPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14995
            
    If lErro <> SUCESSO Then Error 14996
        
    X.objFaturasPagLista.Show
            
    FaturasPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FaturasPagLista_Chama_Tela:

    FaturasPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 14994, 14995
            Unload X.objFaturasPagLista
            
        Case 14996
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146764)
        
    End Select
        
    Exit Function

End Function

Function BorderosPagtoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objBorderoPagto As ClassBorderoPagto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_BorderosPagtoLista_Chama_Tela

    Load X.objBorderosPagtoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45552
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objBorderoPagto = Nothing
        
    Else
    
        Set objBorderoPagto = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objBorderosPagtoLista.Trata_Parametros(colSelecao, objBorderoPagto, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45553
            
    If lErro <> SUCESSO Then Error 45554
        
    X.objBorderosPagtoLista.Show
            
    BorderosPagtoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderosPagtoLista_Chama_Tela:

    BorderosPagtoLista_Chama_Tela = Err

    Select Case Err
    
        Case 45552, 45553
            Unload X.objBorderosPagtoLista
            
        Case 45554
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146765)
        
    End Select
        
    Exit Function

End Function

Function ParcelasRecLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objParcelasRec As ClassParcelaReceber
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ParcelasRecLista_Chama_Tela

    Load X.objParcelasRecLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45084
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objParcelasRec = Nothing
        
    Else
    
        Set objParcelasRec = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objParcelasRecLista.Trata_Parametros(colSelecao, objParcelasRec, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45085
            
    If lErro <> SUCESSO Then Error 45086
        
    X.objParcelasRecLista.Show
            
    ParcelasRecLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ParcelasRecLista_Chama_Tela:

    ParcelasRecLista_Chama_Tela = Err

    Select Case Err
    
        Case 45084, 45085
            Unload X.objParcelasRecLista
            
        Case 45086
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146766)
        
    End Select
        
    Exit Function

End Function

Function ParcelasRecTodasLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objParcelasRec As ClassParcelaReceber
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ParcelasRecTodasLista_Chama_Tela

    Load X.objParcelasRecTodasLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45439
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objParcelasRec = Nothing
        
    Else
    
        Set objParcelasRec = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objParcelasRecTodasLista.Trata_Parametros(colSelecao, objParcelasRec, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45440
            
    If lErro <> SUCESSO Then Error 45441
        
    X.objParcelasRecTodasLista.Show
            
    ParcelasRecTodasLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ParcelasRecTodasLista_Chama_Tela:

    ParcelasRecTodasLista_Chama_Tela = Err

    Select Case Err
    
        Case 45439, 45440
            Unload X.objParcelasRecTodasLista
            
        Case 45441
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146767)
        
    End Select
        
    Exit Function

End Function

Function TitulosRecTodosLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTitulosRec As ClassTituloReceber
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TitulosRecTodosLista_Chama_Tela

    Load X.objTitulosRecTodosLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45442
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTitulosRec = Nothing
        
    Else
    
        Set objTitulosRec = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTitulosRecTodosLista.Trata_Parametros(colSelecao, objTitulosRec, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45443
            
    If lErro <> SUCESSO Then Error 45444
        
    X.objTitulosRecTodosLista.Show
            
    TitulosRecTodosLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TitulosRecTodosLista_Chama_Tela:

    TitulosRecTodosLista_Chama_Tela = Err

    Select Case Err
    
        Case 45442, 45443
            Unload X.objTitulosRecTodosLista
            
        Case 45444
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146768)
        
    End Select
        
    Exit Function

End Function

Function InstrRemParcRecLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objInstrRemParcRec As ClassOcorrRemParcRec
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_InstrRemParcRecLista_Chama_Tela

    Load X.objInstrRemParcRecLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45336
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objInstrRemParcRec = Nothing
        
    Else
    
        Set objInstrRemParcRec = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objInstrRemParcRecLista.Trata_Parametros(colSelecao, objInstrRemParcRec, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45337
            
    If lErro <> SUCESSO Then Error 45338
        
    X.objInstrRemParcRecLista.Show
            
    InstrRemParcRecLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_InstrRemParcRecLista_Chama_Tela:

    InstrRemParcRecLista_Chama_Tela = Err

    Select Case Err
    
        Case 45336, 45337
            Unload X.objInstrRemParcRecLista
            
        Case 45338
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146769)
        
    End Select
        
    Exit Function

End Function

Function FaturasPagLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objFaturasPag1 As ClassTituloPagar
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_FaturasPagLista_Consulta_Chama_Tela

    Load X.objFaturasPagLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19874
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objFaturasPag1 = Nothing
        
    Else
    
        Set objFaturasPag1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objFaturasPagLista_Consulta.Trata_Parametros(colSelecao, objFaturasPag1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19875
            
    If lErro <> SUCESSO Then Error 19876
        
    X.objFaturasPagLista_Consulta.Show
            
    FaturasPagLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FaturasPagLista_Consulta_Chama_Tela:

    FaturasPagLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 19874, 19875
            Unload X.objFaturasPagLista_Consulta
            
        Case 19876
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146770)
        
    End Select
        
    Exit Function

End Function

Function TipoVendedorLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoVendedor1 As ClassTipoVendedor
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoVendedorLista_Chama_Tela

    Load X.objTipoVendedorLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14997
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTipoVendedor1 = Nothing
        
    Else
    
        Set objTipoVendedor1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTipoVendedorLista.Trata_Parametros(colSelecao, objTipoVendedor1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14998
            
    If lErro <> SUCESSO Then Error 14999
        
    X.objTipoVendedorLista.Show
            
    TipoVendedorLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoVendedorLista_Chama_Tela:

    TipoVendedorLista_Chama_Tela = Err

    Select Case Err
    
        Case 14997, 14998
            Unload X.objTipoVendedorLista
            
        Case 14999
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146771)
        
    End Select
        
    Exit Function

End Function

Function TituloReceberLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTituloReceber1 As ClassTituloReceber
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TituloReceberLista_Chama_Tela

    Load X.objTituloReceberLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24368
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTituloReceber1 = Nothing
        
    Else
    
        Set objTituloReceber1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTituloReceberLista.Trata_Parametros(colSelecao, objTituloReceber1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24369
            
    If lErro <> SUCESSO Then Error 24370
        
    X.objTituloReceberLista.Show
            
    TituloReceberLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TituloReceberLista_Chama_Tela:

    TituloReceberLista_Chama_Tela = Err

    Select Case Err
    
        Case 24368, 24369
            Unload X.objTituloReceberLista
            
        Case 24370
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146772)
        
    End Select
        
    Exit Function

End Function

Function TituloReceberLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTituloReceber1 As ClassTituloReceber
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TituloReceberLista_Consulta_Chama_Tela

    Load X.objTituloReceberLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24388
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTituloReceber1 = Nothing
        
    Else
    
        Set objTituloReceber1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTituloReceberLista_Consulta.Trata_Parametros(colSelecao, objTituloReceber1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24389
            
    If lErro <> SUCESSO Then Error 24390
        
    X.objTituloReceberLista_Consulta.Show
            
    TituloReceberLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TituloReceberLista_Consulta_Chama_Tela:

    TituloReceberLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24388, 24389
            Unload X.objTituloReceberLista_Consulta
            
        Case 24390
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146773)
        
    End Select
        
    Exit Function

End Function

Function TabelasPreco_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTabelasPreco As ClassTabelaPreco

On Error GoTo Erro_TabelasPreco_Chama_Tela

    Load X.objTabelasPreco
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14506
            
    If IsMissing(vP1) Then
    
        Set objTabelasPreco = Nothing
        
    Else
    
        Set objTabelasPreco = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTabelasPreco.Trata_Parametros(objTabelasPreco)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14507
            
        If lErro <> SUCESSO Then Error 14508
        
    End If
    
    X.objTabelasPreco.Show
            
    TabelasPreco_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TabelasPreco_Chama_Tela:

    TabelasPreco_Chama_Tela = Err

    Select Case Err
    
        Case 14506, 14507
            Unload X.objTabelasPreco
            
        Case 14508
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146774)
        
    End Select
        
    Exit Function

End Function

Function Deposito_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objDeposito As ClassMovContaCorrente

On Error GoTo Erro_Deposito_Chama_Tela

    Load X.objDeposito
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14536
            
    If IsMissing(vP1) Then
    
        Set objDeposito = Nothing
        
    Else
    
        Set objDeposito = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objDeposito.Trata_Parametros(objDeposito)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14537
            
        If lErro <> SUCESSO Then Error 14538
        
    End If
    
    X.objDeposito.Show
            
    Deposito_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Deposito_Chama_Tela:

    Deposito_Chama_Tela = Err

    Select Case Err
    
        Case 14536, 14537
            Unload X.objDeposito
            
        Case 14538
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146775)
        
    End Select
        
    Exit Function

End Function

Function Saque_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objSaque As ClassMovContaCorrente

On Error GoTo Erro_Saque_Chama_Tela

    Load X.objSaque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14539
            
    If IsMissing(vP1) Then
    
        Set objSaque = Nothing
        
    Else
    
        Set objSaque = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSaque.Trata_Parametros(objSaque)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14540
            
        If lErro <> SUCESSO Then Error 14541
        
    End If
    
    X.objSaque.Show
            
    Saque_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Saque_Chama_Tela:

    Saque_Chama_Tela = Err

    Select Case Err
    
        Case 14539, 14540
            Unload X.objSaque
            
        Case 14541
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146776)
        
    End Select
        
    Exit Function

End Function

Function NFPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFPag As ClassNFsPag

On Error GoTo Erro_NFPag_Chama_Tela

    Load X.objNFPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14733
            
    If IsMissing(vP1) Then
    
        Set objNFPag = Nothing
        
    Else
    
        Set objNFPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFPag.Trata_Parametros(objNFPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14734
            
        If lErro <> SUCESSO Then Error 14735
        
    End If
    
    X.objNFPag.Show
            
    NFPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFPag_Chama_Tela:

    NFPag_Chama_Tela = Err

    Select Case Err
    
        Case 14733, 14734
            Unload X.objNFPag
            
        Case 14735
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146777)
        
    End Select
        
    Exit Function

End Function

Function NFFatPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFFatPag As ClassTituloPagar

On Error GoTo Erro_NFFatPag_Chama_Tela

    Load X.objNFFatPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24116
            
    If IsMissing(vP1) Then
    
        Set objNFFatPag = Nothing
        
    Else
    
        Set objNFFatPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFFatPag.Trata_Parametros(objNFFatPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24117
            
        If lErro <> SUCESSO Then Error 24118
        
    End If
    
    X.objNFFatPag.Show
            
    NFFatPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFFatPag_Chama_Tela:

    NFFatPag_Chama_Tela = Err

    Select Case Err
    
        Case 24116, 24117
            Unload X.objNFFatPag
            
        Case 24118
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146778)
        
    End Select
        
    Exit Function

End Function

Function OutrosPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objOutrosPag As ClassTituloPagar

On Error GoTo Erro_OutrosPag_Chama_Tela

    Load X.objOutrosPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24061
            
    If IsMissing(vP1) Then
    
        Set objOutrosPag = Nothing
        
    Else
    
        Set objOutrosPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objOutrosPag.Trata_Parametros(objOutrosPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24062
            
        If lErro <> SUCESSO Then Error 24063
        
    End If
    
    X.objOutrosPag.Show
            
    OutrosPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OutrosPag_Chama_Tela:

    OutrosPag_Chama_Tela = Err

    Select Case Err
    
        Case 24061, 24062
            Unload X.objOutrosPag
            
        Case 24063
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146779)
        
    End Select
        
    Exit Function

End Function

Function Portadores_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPortadores As ClassPortador

On Error GoTo Erro_Portadores_Chama_Tela

    Load X.objPortadores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14762
            
    If IsMissing(vP1) Then
    
        Set objPortadores = Nothing
        
    Else
    
        Set objPortadores = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPortadores.Trata_Parametros(objPortadores)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14763
            
        If lErro <> SUCESSO Then Error 14764
        
    End If
    
    X.objPortadores.Show
            
    Portadores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Portadores_Chama_Tela:

    Portadores_Chama_Tela = Err

    Select Case Err
    
        Case 14762, 14763
            Unload X.objPortadores
            
        Case 14764
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146780)
        
    End Select
        
    Exit Function

End Function

Function BorderoPag1_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoPag1 As ClassBorderoPagEmissao

On Error GoTo Erro_BorderoPag1_Chama_Tela

    Load X.objBorderoPag1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14810
            
    If IsMissing(vP1) Then
    
        Set objBorderoPag1 = Nothing
        
    Else
    
        Set objBorderoPag1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoPag1.Trata_Parametros(objBorderoPag1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14811
            
        If lErro <> SUCESSO Then Error 14812
        
    End If
    
    X.objBorderoPag1.Show
            
    BorderoPag1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoPag1_Chama_Tela:

    BorderoPag1_Chama_Tela = Err

    Select Case Err
    
        Case 14810, 14811
            Unload X.objBorderoPag1
            
        Case 14812
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146781)
        
    End Select
        
    Exit Function

End Function

Function BorderoPag2_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoPag2 As ClassBorderoPagEmissao

On Error GoTo Erro_BorderoPag2_Chama_Tela

    Load X.objBorderoPag2
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14813
            
    If IsMissing(vP1) Then
    
        Set objBorderoPag2 = Nothing
        
    Else
    
        Set objBorderoPag2 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoPag2.Trata_Parametros(objBorderoPag2)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14814
            
        If lErro <> SUCESSO Then Error 14815
        
    End If
    
    X.objBorderoPag2.Show
            
    BorderoPag2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoPag2_Chama_Tela:

    BorderoPag2_Chama_Tela = Err

    Select Case Err
    
        Case 14813, 14814
            Unload X.objBorderoPag2
            
        Case 14815
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146782)
        
    End Select
        
    Exit Function

End Function

Function BorderoPag3_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoPag3 As ClassBorderoPagEmissao

On Error GoTo Erro_BorderoPag3_Chama_Tela

    Load X.objBorderoPag3
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14816
            
    If IsMissing(vP1) Then
    
        Set objBorderoPag3 = Nothing
        
    Else
    
        Set objBorderoPag3 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoPag3.Trata_Parametros(objBorderoPag3)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14817
            
        If lErro <> SUCESSO Then Error 14818
        
    End If
    
    X.objBorderoPag3.Show
            
    BorderoPag3_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoPag3_Chama_Tela:

    BorderoPag3_Chama_Tela = Err

    Select Case Err
    
        Case 14816, 14817
            Unload X.objBorderoPag3
            
        Case 14818
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146783)
        
    End Select
        
    Exit Function

End Function

Function BorderoPag4_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoPag4 As ClassBorderoPagEmissao

On Error GoTo Erro_BorderoPag4_Chama_Tela

    Load X.objBorderoPag4
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14819
            
    If IsMissing(vP1) Then
    
        Set objBorderoPag4 = Nothing
        
    Else
    
        Set objBorderoPag4 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoPag4.Trata_Parametros(objBorderoPag4)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14820
            
        If lErro <> SUCESSO Then Error 14821
        
    End If
    
    X.objBorderoPag4.Show
            
    BorderoPag4_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoPag4_Chama_Tela:

    BorderoPag4_Chama_Tela = Err

    Select Case Err
    
        Case 14819, 14820
            Unload X.objBorderoPag4
            
        Case 14821
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146784)
        
    End Select
        
    Exit Function

End Function

Function BorderoPag5_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoPag5 As ClassBorderoPagEmissao

On Error GoTo Erro_BorderoPag5_Chama_Tela

    Load X.objBorderoPag5
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14822
            
    If IsMissing(vP1) Then
    
        Set objBorderoPag5 = Nothing
        
    Else
    
        Set objBorderoPag5 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoPag5.Trata_Parametros(objBorderoPag5)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14823
            
        If lErro <> SUCESSO Then Error 14824
        
    End If
    
    X.objBorderoPag5.Show
            
    BorderoPag5_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoPag5_Chama_Tela:

    BorderoPag5_Chama_Tela = Err

    Select Case Err
    
        Case 14822, 14823
            Unload X.objBorderoPag5
            
        Case 14824
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146785)
        
    End Select
        
    Exit Function

End Function

Function BorderoCobranca_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoCobranca As ClassBorderoCobrancaEmissao

On Error GoTo Erro_BorderoCobranca_Chama_Tela

    Load X.objBorderoCobranca
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24003
            
    If IsMissing(vP1) Then
    
        Set objBorderoCobranca = Nothing
        
    Else
    
        Set objBorderoCobranca = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoCobranca.Trata_Parametros(objBorderoCobranca)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24004
            
        If lErro <> SUCESSO Then Error 24005
        
    End If
    
    X.objBorderoCobranca.Show
            
    BorderoCobranca_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoCobranca_Chama_Tela:

    BorderoCobranca_Chama_Tela = Err

    Select Case Err
    
        Case 24003, 24004
            Unload X.objBorderoCobranca
            
        Case 24005
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146786)
        
    End Select
        
    Exit Function

End Function

Function BorderoCobranca2_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objBorderoCobranca2 As ClassBorderoCobrancaEmissao

On Error GoTo Erro_BorderoCobranca2_Chama_Tela

    Load X.objBorderoCobranca2
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24006
            
    If IsMissing(vP1) Then
    
        Set objBorderoCobranca2 = Nothing
        
    Else
    
        Set objBorderoCobranca2 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objBorderoCobranca2.Trata_Parametros(objBorderoCobranca2)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24007
            
        If lErro <> SUCESSO Then Error 24008
        
    End If
    
    X.objBorderoCobranca2.Show
            
    BorderoCobranca2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoCobranca2_Chama_Tela:

    BorderoCobranca2_Chama_Tela = Err

    Select Case Err
    
        Case 24006, 24007
            Unload X.objBorderoCobranca2
            
        Case 24008
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146787)
        
    End Select
        
    Exit Function

End Function

Function ChequePagAvulso1_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequePagAvulso1 As ClassChequesPagAvulso

On Error GoTo Erro_ChequePagAvulso1_Chama_Tela

    Load X.objChequePagAvulso1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14825
            
    If IsMissing(vP1) Then
    
        Set objChequePagAvulso1 = Nothing
        
    Else
    
        Set objChequePagAvulso1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequePagAvulso1.Trata_Parametros(objChequePagAvulso1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14826
            
        If lErro <> SUCESSO Then Error 14827
        
    End If
    
    X.objChequePagAvulso1.Show
            
    ChequePagAvulso1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePagAvulso1_Chama_Tela:

    ChequePagAvulso1_Chama_Tela = Err

    Select Case Err
    
        Case 14825, 14826
            Unload X.objChequePagAvulso1
            
        Case 14827
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146788)
        
    End Select
        
    Exit Function

End Function

Function ChequePagAvulso2_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequePagAvulso2 As ClassChequesPagAvulso

On Error GoTo Erro_ChequePagAvulso2_Chama_Tela

    Load X.objChequePagAvulso2
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14828
            
    If IsMissing(vP1) Then
    
        Set objChequePagAvulso2 = Nothing
        
    Else
    
        Set objChequePagAvulso2 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequePagAvulso2.Trata_Parametros(objChequePagAvulso2)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14829
            
        If lErro <> SUCESSO Then Error 14830
        
    End If
    
    X.objChequePagAvulso2.Show
            
    ChequePagAvulso2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePagAvulso2_Chama_Tela:

    ChequePagAvulso2_Chama_Tela = Err

    Select Case Err
    
        Case 14828, 14829
            Unload X.objChequePagAvulso2
            
        Case 14830
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146789)
        
    End Select
        
    Exit Function

End Function

Function ChequePagAvulso3_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequePagAvulso3 As ClassChequesPagAvulso

On Error GoTo Erro_ChequePagAvulso3_Chama_Tela

    Load X.objChequePagAvulso3
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14831
            
    If IsMissing(vP1) Then
    
        Set objChequePagAvulso3 = Nothing
        
    Else
    
        Set objChequePagAvulso3 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequePagAvulso3.Trata_Parametros(objChequePagAvulso3)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14832
            
        If lErro <> SUCESSO Then Error 14833
        
    End If
    
    X.objChequePagAvulso3.Show
            
    ChequePagAvulso3_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePagAvulso3_Chama_Tela:

    ChequePagAvulso3_Chama_Tela = Err

    Select Case Err
    
        Case 14831, 14832
            Unload X.objChequePagAvulso3
            
        Case 14833
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146790)
        
    End Select
        
    Exit Function

End Function

Function ChequePagAvulso4_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequePagAvulso4 As ClassChequesPagAvulso

On Error GoTo Erro_ChequePagAvulso4_Chama_Tela

    Load X.objChequePagAvulso4
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14834
            
    If IsMissing(vP1) Then
    
        Set objChequePagAvulso4 = Nothing
        
    Else
    
        Set objChequePagAvulso4 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequePagAvulso4.Trata_Parametros(objChequePagAvulso4)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14835
            
        If lErro <> SUCESSO Then Error 14836
        
    End If
    
    X.objChequePagAvulso4.Show
            
    ChequePagAvulso4_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequePagAvulso4_Chama_Tela:

    ChequePagAvulso4_Chama_Tela = Err

    Select Case Err
    
        Case 14834, 14835
            Unload X.objChequePagAvulso4
            
        Case 14836
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146791)
        
    End Select
        
    Exit Function

End Function

Function ChequesPag_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequesPag As ClassChequesPag

On Error GoTo Erro_ChequesPag_Chama_Tela

    Load X.objChequesPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14837
            
    If IsMissing(vP1) Then
    
        Set objChequesPag = Nothing
        
    Else
    
        Set objChequesPag = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequesPag.Trata_Parametros(objChequesPag)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14838
            
        If lErro <> SUCESSO Then Error 14839
        
    End If
    
    X.objChequesPag.Show
            
    ChequesPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequesPag_Chama_Tela:

    ChequesPag_Chama_Tela = Err

    Select Case Err
    
        Case 14837, 14838
            Unload X.objChequesPag
            
        Case 14839
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146792)
        
    End Select
        
    Exit Function

End Function

Function ChequesPag2_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequesPag2 As ClassChequesPag

On Error GoTo Erro_ChequesPag2_Chama_Tela

    Load X.objChequesPag2
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14840
            
    If IsMissing(vP1) Then
    
        Set objChequesPag2 = Nothing
        
    Else
    
        Set objChequesPag2 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequesPag2.Trata_Parametros(objChequesPag2)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14841
            
        If lErro <> SUCESSO Then Error 14842
        
    End If
    
    X.objChequesPag2.Show
            
    ChequesPag2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequesPag2_Chama_Tela:

    ChequesPag2_Chama_Tela = Err

    Select Case Err
    
        Case 14840, 14841
            Unload X.objChequesPag2
            
        Case 14842
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146793)
        
    End Select
        
    Exit Function

End Function

Function ChequesPag3_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequesPag3 As ClassChequesPag

On Error GoTo Erro_ChequesPag3_Chama_Tela

    Load X.objChequesPag3
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14843
            
    If IsMissing(vP1) Then
    
        Set objChequesPag3 = Nothing
        
    Else
    
        Set objChequesPag3 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequesPag3.Trata_Parametros(objChequesPag3)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14844
            
        If lErro <> SUCESSO Then Error 14845
        
    End If
    
    X.objChequesPag3.Show
            
    ChequesPag3_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequesPag3_Chama_Tela:

    ChequesPag3_Chama_Tela = Err

    Select Case Err
    
        Case 14843, 14844
            Unload X.objChequesPag3
            
        Case 14845
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146794)
        
    End Select
        
    Exit Function

End Function

Function ChequesPag4_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objChequesPag4 As ClassChequesPag

On Error GoTo Erro_ChequesPag4_Chama_Tela

    Load X.objChequesPag4
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14846
            
    If IsMissing(vP1) Then
    
        Set objChequesPag4 = Nothing
        
    Else
    
        Set objChequesPag4 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objChequesPag4.Trata_Parametros(objChequesPag4)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14847
            
        If lErro <> SUCESSO Then Error 14848
        
    End If
    
    X.objChequesPag4.Show
            
    ChequesPag4_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ChequesPag4_Chama_Tela:

    ChequesPag4_Chama_Tela = Err

    Select Case Err
    
        Case 14846, 14847
            Unload X.objChequesPag4
            
        Case 14848
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146795)
        
    End Select
        
    Exit Function

End Function

Function CategoriaCliente_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCategoriaCliente As ClassCategoriaCliente

On Error GoTo Erro_CategoriaCliente_Chama_Tela

    Load X.objCategoriaCliente
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24815
            
    If IsMissing(vP1) Then
    
        Set objCategoriaCliente = Nothing
        
    Else
    
        Set objCategoriaCliente = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCategoriaCliente.Trata_Parametros(objCategoriaCliente)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24816
            
        If lErro <> SUCESSO Then Error 24817
        
    End If
    
    X.objCategoriaCliente.Show
            
    CategoriaCliente_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CategoriaCliente_Chama_Tela:

    CategoriaCliente_Chama_Tela = Err

    Select Case Err
    
        Case 24815, 24816
            Unload X.objCategoriaCliente
            
        Case 24817
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146796)
        
    End Select
        
    Exit Function

End Function

Function RelOpComisVend_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpComisVend_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14626
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpComisVend.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14627
    
    Load X.objRelOpComisVend
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14628
            
    X.objRelOpComisVend.Show
            
    RelOpComisVend_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpComisVend_Chama_Tela:

    RelOpComisVend_Chama_Tela = Err

    Select Case Err
    
        Case 14627
    
        Case 14628
            Unload X.objRelOpComisVend
        
        Case 14626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146797)
        
    End Select
        
    Exit Function

End Function

Function RelOpDevedores_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDevedores_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14629
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDevedores.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14630
    
    Load X.objRelOpDevedores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14631
            
    X.objRelOpDevedores.Show
            
    RelOpDevedores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDevedores_Chama_Tela:

    RelOpDevedores_Chama_Tela = Err

    Select Case Err
    
        Case 14630
    
        Case 14631
            Unload X.objRelOpDevedores
        
        Case 14629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146798)
        
    End Select
        
    Exit Function

End Function

Function RelOpDiarioCP_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDiarioCP_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14632
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDiarioCP.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14633
    
    Load X.objRelOpDiarioCP
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14634
            
    X.objRelOpDiarioCP.Show
            
    RelOpDiarioCP_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDiarioCP_Chama_Tela:

    RelOpDiarioCP_Chama_Tela = Err

    Select Case Err
    
        Case 14633
    
        Case 14634
            Unload X.objRelOpDiarioCP
        
        Case 14632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146799)
        
    End Select
        
    Exit Function

End Function

Function RelOpDiarioCR_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDiarioCR_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14635
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
            
    End If
    
    lErro = X.objRelOpDiarioCR.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14636
    
    Load X.objRelOpDiarioCR
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14637
            
    X.objRelOpDiarioCR.Show
            
    RelOpDiarioCR_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDiarioCR_Chama_Tela:

    RelOpDiarioCR_Chama_Tela = Err

    Select Case Err
    
        Case 14636
    
        Case 14637
            Unload X.objRelOpDiarioCR
        
        Case 14635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146800)
        
    End Select
        
    Exit Function

End Function

Function RelOpHistAplic_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpHistAplic_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14638
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpHistAplic.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14639
    
    Load X.objRelOpHistAplic
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14640
            
    X.objRelOpHistAplic.Show
            
    RelOpHistAplic_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpHistAplic_Chama_Tela:

    RelOpHistAplic_Chama_Tela = Err

    Select Case Err
    
        Case 14639
    
        Case 14640
            Unload X.objRelOpHistAplic
        
        Case 14638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146801)
        
    End Select
        
    Exit Function

End Function

Function RelOpMovFin_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpMovFin_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14641
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpMovFin.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14642
    
    Load X.objRelOpMovFin
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14643
            
    X.objRelOpMovFin.Show
            
    RelOpMovFin_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpMovFin_Chama_Tela:

    RelOpMovFin_Chama_Tela = Err

    Select Case Err
    
        Case 14642
    
        Case 14643
            Unload X.objRelOpMovFin
        
        Case 14641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146802)
        
    End Select
        
    Exit Function

End Function

Function RelOpMovFinDet_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpMovFinDet_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14644
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
    
    End If
    
    lErro = X.objRelOpMovFinDet.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14645
    
    Load X.objRelOpMovFinDet
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14646
            
    X.objRelOpMovFinDet.Show
            
    RelOpMovFinDet_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpMovFinDet_Chama_Tela:

    RelOpMovFinDet_Chama_Tela = Err

    Select Case Err
    
        Case 14645
    
        Case 14646
            Unload X.objRelOpMovFinDet
        
        Case 14644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146803)
        
    End Select
        
    Exit Function

End Function

Function RelOpPosAplic_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPosAplic_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14647
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
            
    End If
    
    lErro = X.objRelOpPosAplic.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14648
    
    Load X.objRelOpPosAplic
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14649
            
    X.objRelOpPosAplic.Show
            
    RelOpPosAplic_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPosAplic_Chama_Tela:

    RelOpPosAplic_Chama_Tela = Err

    Select Case Err
    
        Case 14648
    
        Case 14649
            Unload X.objRelOpPosAplic
        
        Case 14647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146804)
        
    End Select
        
    Exit Function

End Function

Function RelOpPosCli_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPosCli_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14650
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPosCli.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14651
    
    Load X.objRelOpPosCli
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14652
            
    X.objRelOpPosCli.Show
            
    RelOpPosCli_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPosCli_Chama_Tela:

    RelOpPosCli_Chama_Tela = Err

    Select Case Err
    
        Case 14651
    
        Case 14652
            Unload X.objRelOpPosCli
        
        Case 14650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146805)
        
    End Select
        
    Exit Function

End Function

Function RelOpRazaoCP_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRazaoCP_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14653
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRazaoCP.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14654
    
    Load X.objRelOpRazaoCP
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14655
            
    X.objRelOpRazaoCP.Show
            
    RelOpRazaoCP_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRazaoCP_Chama_Tela:

    RelOpRazaoCP_Chama_Tela = Err

    Select Case Err
    
        Case 14654
    
        Case 14655
            Unload X.objRelOpRazaoCP
        
        Case 14653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146806)
        
    End Select
        
    Exit Function

End Function

Function RelOpRazaoCR_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRazaoCR_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14656
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRazaoCR.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14657
    
    Load X.objRelOpRazaoCR
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14658
            
    X.objRelOpRazaoCR.Show
            
    RelOpRazaoCR_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRazaoCR_Chama_Tela:

    RelOpRazaoCR_Chama_Tela = Err

    Select Case Err
    
        Case 14657
    
        Case 14658
            Unload X.objRelOpRazaoCR
        
        Case 14656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146807)
        
    End Select
        
    Exit Function

End Function

Function RelOpResTitRec_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpResTitRec_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14659
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpResTitRec.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14660
    
    Load X.objRelOpResTitRec
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14661
            
    X.objRelOpResTitRec.Show
            
    RelOpResTitRec_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpResTitRec_Chama_Tela:

    RelOpResTitRec_Chama_Tela = Err

    Select Case Err
    
        Case 14660
    
        Case 14661
            Unload X.objRelOpResTitRec
        
        Case 14659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146808)
        
    End Select
        
    Exit Function

End Function

Function RelOpResumoComis_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpResumoComis_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14662
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpResumoComis.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14663
    
    Load X.objRelOpResumoComis
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14664
            
    X.objRelOpResumoComis.Show
            
    RelOpResumoComis_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpResumoComis_Chama_Tela:

    RelOpResumoComis_Chama_Tela = Err

    Select Case Err
    
        Case 14663
    
        Case 14664
            Unload X.objRelOpResumoComis
        
        Case 14662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146809)
        
    End Select
        
    Exit Function

End Function

Function RelOpTitRecMala_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpTitRecMala_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14665
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpTitRecMala.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14666
    
    Load X.objRelOpTitRecMala
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14667
            
    X.objRelOpTitRecMala.Show
            
    RelOpTitRecMala_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpTitRecMala_Chama_Tela:

    RelOpTitRecMala_Chama_Tela = Err

    Select Case Err
    
        Case 14666
    
        Case 14667
            Unload X.objRelOpTitRecMala
        
        Case 14665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146810)
        
    End Select
        
    Exit Function

End Function

Function RelOpTitRecTel_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpTitRecTel_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14668
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpTitRecTel.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14669
    
    Load X.objRelOpTitRecTel
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14670
            
    X.objRelOpTitRecTel.Show
            
    RelOpTitRecTel_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpTitRecTel_Chama_Tela:

    RelOpTitRecTel_Chama_Tela = Err

    Select Case Err
    
        Case 14669
    
        Case 14670
            Unload X.objRelOpTitRecTel
        
        Case 14668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146811)
        
    End Select
        
    Exit Function

End Function

Function RelOpCadCli_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpCadCli_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45179
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpCadCli.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45180
    
    Load X.objRelOpCadCli
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45181
            
    X.objRelOpCadCli.Show
            
    RelOpCadCli_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpCadCli_Chama_Tela:

    RelOpCadCli_Chama_Tela = Err

    Select Case Err
    
        Case 45179
    
        Case 45180
            Unload X.objRelOpCadCli
        
        Case 45181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146812)
        
    End Select
        
    Exit Function

End Function

Function RelOpCadForn_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpCadForn_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45182
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpCadForn.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45183
    
    Load X.objRelOpCadForn
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45184
            
    X.objRelOpCadForn.Show
            
    RelOpCadForn_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpCadForn_Chama_Tela:

    RelOpCadForn_Chama_Tela = Err

    Select Case Err
    
        Case 45182
    
        Case 45183
            Unload X.objRelOpCadForn
        
        Case 45184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146813)
        
    End Select
        
    Exit Function

End Function

Function RelOpCadVend_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpCadVend_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45185
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpCadVend.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45186
    
    Load X.objRelOpCadVend
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45187
            
    X.objRelOpCadVend.Show
            
    RelOpCadVend_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpCadVend_Chama_Tela:

    RelOpCadVend_Chama_Tela = Err

    Select Case Err
    
        Case 45185
    
        Case 45186
            Unload X.objRelOpCadVend
        
        Case 45187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146814)
        
    End Select
        
    Exit Function

End Function

Function RelOpTitRec_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpTitRec_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45188
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpTitRec.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45189
    
    Load X.objRelOpTitRec
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45190
            
    X.objRelOpTitRec.Show
            
    RelOpTitRec_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpTitRec_Chama_Tela:

    RelOpTitRec_Chama_Tela = Err

    Select Case Err
    
        Case 45188
    
        Case 45189
            Unload X.objRelOpTitRec
        
        Case 45190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146815)
        
    End Select
        
    Exit Function

End Function

Function RelOpTitPag_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpTitPag_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45191
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpTitPag.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45192
    
    Load X.objRelOpTitPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45193
            
    X.objRelOpTitPag.Show
            
    RelOpTitPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpTitPag_Chama_Tela:

    RelOpTitPag_Chama_Tela = Err

    Select Case Err
    
        Case 45191
    
        Case 45192
            Unload X.objRelOpTitPag
        
        Case 45193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146816)
        
    End Select
        
    Exit Function

End Function

Function RelOpJurosRec_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpJurosRec_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45194
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpJurosRec.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45195
    
    Load X.objRelOpJurosRec
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45196
            
    X.objRelOpJurosRec.Show
            
    RelOpJurosRec_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpJurosRec_Chama_Tela:

    RelOpJurosRec_Chama_Tela = Err

    Select Case Err
    
        Case 45194
    
        Case 45195
            Unload X.objRelOpJurosRec
        
        Case 45196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146817)
        
    End Select
        
    Exit Function

End Function

Function RelOpBaixasCP_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBaixasCP_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45225
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpBaixasCP.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45226
    
    Load X.objRelOpBaixasCP
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45227
            
    X.objRelOpBaixasCP.Show
            
    RelOpBaixasCP_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBaixasCP_Chama_Tela:

    RelOpBaixasCP_Chama_Tela = Err

    Select Case Err
    
        Case 45225
    
        Case 45226
            Unload X.objRelOpBaixasCP
        
        Case 45227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146818)
        
    End Select
        
    Exit Function

End Function

Function RelOpPagtosCancelados_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPagtosCancelados_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45228
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPagtosCancelados.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45229
    
    Load X.objRelOpPagtosCancelados
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45230
            
    X.objRelOpPagtosCancelados.Show
            
    RelOpPagtosCancelados_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPagtosCancelados_Chama_Tela:

    RelOpPagtosCancelados_Chama_Tela = Err

    Select Case Err
    
        Case 45228
    
        Case 45229
            Unload X.objRelOpPagtosCancelados
        
        Case 45230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146819)
        
    End Select
        
    Exit Function

End Function

Function RelOpRelCheques_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRelCheques_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45231
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRelCheques.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45232
    
    Load X.objRelOpRelCheques
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45233
            
    X.objRelOpRelCheques.Show
            
    RelOpRelCheques_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRelCheques_Chama_Tela:

    RelOpRelCheques_Chama_Tela = Err

    Select Case Err
    
        Case 45231
    
        Case 45232
            Unload X.objRelOpRelCheques
        
        Case 45233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146820)
        
    End Select
        
    Exit Function

End Function

Function RelOpPosForn_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPosForn_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45256
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPosForn.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45257
    
    Load X.objRelOpPosForn
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45258
            
    X.objRelOpPosForn.Show
            
    RelOpPosForn_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPosForn_Chama_Tela:

    RelOpPosForn_Chama_Tela = Err

    Select Case Err
    
        Case 45256
    
        Case 45257
            Unload X.objRelOpPosForn
        
        Case 45258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146821)
        
    End Select
        
    Exit Function

End Function

Function RelOpNFPagSemFat_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpNFPagSemFat_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45259
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpNFPagSemFat.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45260
    
    Load X.objRelOpNFPagSemFat
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45261
            
    X.objRelOpNFPagSemFat.Show
            
    RelOpNFPagSemFat_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpNFPagSemFat_Chama_Tela:

    RelOpNFPagSemFat_Chama_Tela = Err

    Select Case Err
    
        Case 45259
    
        Case 45260
            Unload X.objRelOpNFPagSemFat
        
        Case 45261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146822)
        
    End Select
        
    Exit Function

End Function

Function RelOpBorderoPag_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBorderoPag_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45262
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpBorderoPag.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45263
    
    Load X.objRelOpBorderoPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45264
            
    X.objRelOpBorderoPag.Show
            
    RelOpBorderoPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBorderoPag_Chama_Tela:

    RelOpBorderoPag_Chama_Tela = Err

    Select Case Err
    
        Case 45262
    
        Case 45263
            Unload X.objRelOpBorderoPag
        
        Case 45264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146823)
        
    End Select
        
    Exit Function

End Function

Function RelOpCliAtraso_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpCliAtraso_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45265
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpCliAtraso.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45266
    
    Load X.objRelOpCliAtraso
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45267
            
    X.objRelOpCliAtraso.Show
            
    RelOpCliAtraso_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpCliAtraso_Chama_Tela:

    RelOpCliAtraso_Chama_Tela = Err

    Select Case Err
    
        Case 45265
    
        Case 45266
            Unload X.objRelOpCliAtraso
        
        Case 45267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146824)
        
    End Select
        
    Exit Function

End Function

Function RelOpDIRF_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDIRF_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45315
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDIRF.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45316
    
    Load X.objRelOpDIRF
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45317
            
    X.objRelOpDIRF.Show
            
    RelOpDIRF_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDIRF_Chama_Tela:

    RelOpDIRF_Chama_Tela = Err

    Select Case Err
    
        Case 45315
    
        Case 45316
            Unload X.objRelOpDIRF
        
        Case 45317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146825)
        
    End Select
        
    Exit Function

End Function

Function RelOpICMS_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpICMS_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45318
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpICMS.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45319
    
    Load X.objRelOpICMS
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45320
            
    X.objRelOpICMS.Show
            
    RelOpICMS_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpICMS_Chama_Tela:

    RelOpICMS_Chama_Tela = Err

    Select Case Err
    
        Case 45318
    
        Case 45319
            Unload X.objRelOpICMS
        
        Case 45320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146826)
        
    End Select
        
    Exit Function

End Function

Function RelOpINSSRet_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpINSSRet_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45321
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpINSSRet.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45322
    
    Load X.objRelOpINSSRet
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45323
            
    X.objRelOpINSSRet.Show
            
    RelOpINSSRet_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpINSSRet_Chama_Tela:

    RelOpINSSRet_Chama_Tela = Err

    Select Case Err
    
        Case 45321
    
        Case 45322
            Unload X.objRelOpINSSRet
        
        Case 45323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146827)
        
    End Select
        
    Exit Function

End Function

Function RelOpIRRF_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpIRRF_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45324
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpIRRF.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45325
    
    Load X.objRelOpIRRF
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45326
            
    X.objRelOpIRRF.Show
            
    RelOpIRRF_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpIRRF_Chama_Tela:

    RelOpIRRF_Chama_Tela = Err

    Select Case Err
    
        Case 45324
    
        Case 45325
            Unload X.objRelOpIRRF
        
        Case 45326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146828)
        
    End Select
        
    Exit Function

End Function

Function ConciliacaoBancaria_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ConciliacaoBancaria_Chama_Tela

    Load X.objConciliacaoBancaria
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 10903
            
    X.objConciliacaoBancaria.Show
            
    ConciliacaoBancaria_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ConciliacaoBancaria_Chama_Tela:

    ConciliacaoBancaria_Chama_Tela = Err

    Select Case Err
    
        Case 10903
            Unload X.objConciliacaoBancaria
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146829)
        
    End Select
        
    Exit Function

End Function

Function FluxoDeCaixa_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_FluxoDeCaixa_Chama_Tela

    Load X.objFluxoDeCaixa
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20266
            
    X.objFluxoDeCaixa.Show
            
    FluxoDeCaixa_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoDeCaixa_Chama_Tela:

    FluxoDeCaixa_Chama_Tela = Err

    Select Case Err
    
        Case 20266
            Unload X.objFluxoDeCaixa
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146830)
        
    End Select
        
    Exit Function

End Function

Function FluxoPag_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoPag_Chama_Tela

    Load X.objFluxoPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20267
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoPag.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20268
            
    If lErro <> SUCESSO Then Error 20269
        
    X.objFluxoPag.Show
            
    FluxoPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoPag_Chama_Tela:

    FluxoPag_Chama_Tela = Err

    Select Case Err
    
        Case 20267, 20268
            Unload X.objFluxoPag
            
        Case 20269
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146831)
        
    End Select
        
    Exit Function

End Function

Function FluxoReceb_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoReceb_Chama_Tela

    Load X.objFluxoReceb
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20270
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoReceb.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20271
            
    If lErro <> SUCESSO Then Error 20272
        
    X.objFluxoReceb.Show
            
    FluxoReceb_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoReceb_Chama_Tela:

    FluxoReceb_Chama_Tela = Err

    Select Case Err
    
        Case 20270, 20271
            Unload X.objFluxoReceb
            
        Case 20272
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146832)
        
    End Select
        
    Exit Function

End Function

Function FluxoPagForn_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoPagForn_Chama_Tela

    Load X.objFluxoPagForn
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20273
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoPagForn.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20274
            
    If lErro <> SUCESSO Then Error 20275
        
    X.objFluxoPagForn.Show
            
    FluxoPagForn_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoPagForn_Chama_Tela:

    FluxoPagForn_Chama_Tela = Err

    Select Case Err
    
        Case 20273, 20274
            Unload X.objFluxoPagForn
            
        Case 20275
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146833)
        
    End Select
        
    Exit Function

End Function

Function FluxoRecebCli_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoRecebCli_Chama_Tela

    Load X.objFluxoRecebCli
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20310
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoRecebCli.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20311
            
    If lErro <> SUCESSO Then Error 20312
        
    X.objFluxoRecebCli.Show
            
    FluxoRecebCli_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoRecebCli_Chama_Tela:

    FluxoRecebCli_Chama_Tela = Err

    Select Case Err
    
        Case 20310, 20311
            Unload X.objFluxoRecebCli
            
        Case 20312
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146834)
        
    End Select
        
    Exit Function

End Function

Function FluxoTipoForn_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoTipoForn_Chama_Tela

    Load X.objFluxoTipoForn
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20313
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoTipoForn.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20314
            
    If lErro <> SUCESSO Then Error 20315
        
    X.objFluxoTipoForn.Show
            
    FluxoTipoForn_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoTipoForn_Chama_Tela:

    FluxoTipoForn_Chama_Tela = Err

    Select Case Err
    
        Case 20313, 20314
            Unload X.objFluxoTipoForn
            
        Case 20315
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146835)
        
    End Select
        
    Exit Function

End Function

Function FluxoTipoCli_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoTipoCli_Chama_Tela

    Load X.objFluxoTipoCli
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20316
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoTipoCli.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20317
            
    If lErro <> SUCESSO Then Error 20318
        
    X.objFluxoTipoCli.Show
            
    FluxoTipoCli_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoTipoCli_Chama_Tela:

    FluxoTipoCli_Chama_Tela = Err

    Select Case Err
    
        Case 20316, 20317
            Unload X.objFluxoTipoCli
            
        Case 20318
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146836)
        
    End Select
        
    Exit Function

End Function

Function FluxoSintProj_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoSintProj_Chama_Tela

    Load X.objFluxoSintProj
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20326
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoSintProj.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20327
            
    If lErro <> SUCESSO Then Error 20328
        
    X.objFluxoSintProj.Show
            
    FluxoSintProj_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoSintProj_Chama_Tela:

    FluxoSintProj_Chama_Tela = Err

    Select Case Err
    
        Case 20326, 20327
            Unload X.objFluxoSintProj
            
        Case 20328
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146837)
        
    End Select
        
    Exit Function

End Function

Function FluxoSintRev_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoSintRev_Chama_Tela

    Load X.objFluxoSintRev
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20329
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoSintRev.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20330
            
    If lErro <> SUCESSO Then Error 20331
        
    X.objFluxoSintRev.Show
            
    FluxoSintRev_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoSintRev_Chama_Tela:

    FluxoSintRev_Chama_Tela = Err

    Select Case Err
    
        Case 20329, 20330
            Unload X.objFluxoSintRev
            
        Case 20331
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146838)
        
    End Select
        
    Exit Function

End Function

Function FluxoAplic_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoAplic_Chama_Tela

    Load X.objFluxoAplic
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24166
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoAplic.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24167
            
    If lErro <> SUCESSO Then Error 24168
        
    X.objFluxoAplic.Show
            
    FluxoAplic_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoAplic_Chama_Tela:

    FluxoAplic_Chama_Tela = Err

    Select Case Err
    
        Case 24166, 24167
            Unload X.objFluxoAplic
            
        Case 24168
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146839)
        
    End Select
        
    Exit Function

End Function

Function FluxoTipoAplic_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoTipoAplic_Chama_Tela

    Load X.objFluxoTipoAplic
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24169
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoTipoAplic.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24170
            
    If lErro <> SUCESSO Then Error 24171
        
    X.objFluxoTipoAplic.Show
            
    FluxoTipoAplic_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoTipoAplic_Chama_Tela:

    FluxoTipoAplic_Chama_Tela = Err

    Select Case Err
    
        Case 24169, 24170
            Unload X.objFluxoTipoAplic
            
        Case 24171
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146840)
        
    End Select
        
    Exit Function

End Function

Function FluxoSaldoInicial_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objFluxo As ClassFluxo

On Error GoTo Erro_FluxoSaldoInicial_Chama_Tela

    Load X.objFluxoSaldoInicial
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24172
            
    Set objFluxo = vP1
    
    lErro = X.objFluxoSaldoInicial.Trata_Parametros(objFluxo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24173
            
    If lErro <> SUCESSO Then Error 24174
        
    X.objFluxoSaldoInicial.Show
            
    FluxoSaldoInicial_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FluxoSaldoInicial_Chama_Tela:

    FluxoSaldoInicial_Chama_Tela = Err

    Select Case Err
    
        Case 24172, 24173
            Unload X.objFluxoSaldoInicial
            
        Case 24174
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146841)
        
    End Select
        
    Exit Function

End Function

Function GeracaoPlanoConta_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoPlanoConta_Chama_Tela

    Load X.objGeracaoPlanoConta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24523
    
    X.objGeracaoPlanoConta.Show
            
    GeracaoPlanoConta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoPlanoConta_Chama_Tela:

    GeracaoPlanoConta_Chama_Tela = Err

    Select Case Err
    
        Case 24523
            Unload X.objGeracaoPlanoConta
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146842)
        
    End Select
        
    Exit Function

End Function

Function ComissoesPag_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ComissoesPag_Chama_Tela

    Load X.objComissoesPag
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19771
    
    X.objComissoesPag.Show
            
    ComissoesPag_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ComissoesPag_Chama_Tela:

    ComissoesPag_Chama_Tela = Err

    Select Case Err
    
        Case 19771
            Unload X.objComissoesPag
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146843)
        
    End Select
        
    Exit Function

End Function

Function BorderoChequesPre_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoChequesPre_Chama_Tela

    Load X.objBorderoChequesPre
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19772
    
    X.objBorderoChequesPre.Show
            
    BorderoChequesPre_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BorderoChequesPre_Chama_Tela:

    BorderoChequesPre_Chama_Tela = Err

    Select Case Err
    
        Case 19772
            Unload X.objBorderoChequesPre
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146844)
        
    End Select
        
    Exit Function

End Function

Function CobradorLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCobrador As ClassCobrador
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CobradorLista_Chama_Tela

    Load X.objCobradorLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19637
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
                   
    If IsMissing(vP2) Then
    
        Set objCobrador = Nothing
        
    Else
    
        Set objCobrador = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objCobradorLista.Trata_Parametros(colSelecao, objCobrador, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19638
            
    If lErro <> SUCESSO Then Error 19639
        
    X.objCobradorLista.Show
            
    CobradorLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CobradorLista_Chama_Tela:

    CobradorLista_Chama_Tela = Err

    Select Case Err
    
        Case 19637, 19638
            Unload X.objCobradorLista
            
        Case 19639
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146845)
        
    End Select
        
    Exit Function

End Function

Function MensagemLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objMensagem As ClassMensagem
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_MensagemLista_Chama_Tela

    Load X.objMensagemLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19640
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
                  
    If IsMissing(vP2) Then
    
        Set objMensagem = Nothing
        
    Else
    
        Set objMensagem = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objMensagemLista.Trata_Parametros(colSelecao, objMensagem, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19641
            
    If lErro <> SUCESSO Then Error 19642
        
    X.objMensagemLista.Show
            
    MensagemLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MensagemLista_Chama_Tela:

    MensagemLista_Chama_Tela = Err

    Select Case Err
    
        Case 19640, 19641
            Unload X.objMensagemLista
            
        Case 19642
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146846)
        
    End Select
        
    Exit Function

End Function

Function FilialEmpresaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objFilialEmpresa As AdmFiliais
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_FilialEmpresaLista_Chama_Tela

    Load X.objFilialEmpresaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45289
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
                  
    If IsMissing(vP2) Then
    
        Set objFilialEmpresa = Nothing
        
    Else
    
        Set objFilialEmpresa = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objFilialEmpresaLista.Trata_Parametros(colSelecao, objFilialEmpresa, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45290
            
    If lErro <> SUCESSO Then Error 45291
        
    X.objFilialEmpresaLista.Show
            
    FilialEmpresaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FilialEmpresaLista_Chama_Tela:

    FilialEmpresaLista_Chama_Tela = Err

    Select Case Err
    
        Case 45289, 45290
            Unload X.objFilialEmpresaLista
            
        Case 45291
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146847)
        
    End Select
        
    Exit Function

End Function

Function ConfiguraTES_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ConfiguraTES_Chama_Tela

    Load X.objConfiguraTES
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44697
            
    lErro = X.objConfiguraTES.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44698
            
    If lErro <> SUCESSO Then Error 44699
        
    X.objConfiguraTES.Show
            
    ConfiguraTES_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ConfiguraTES_Chama_Tela:

    ConfiguraTES_Chama_Tela = Err

    Select Case Err
    
        Case 44697
            
        Case 44698, 44699
            Unload X.objConfiguraTES
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146848)
        
    End Select
        
    Exit Function

End Function

