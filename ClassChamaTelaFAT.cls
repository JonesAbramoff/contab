VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassChamaTelaFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function NaturezaOperacao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNaturezaOperacao As ClassNaturezaOp

On Error GoTo Erro_NaturezaOperacao_Chama_Tela

    Load X.objNaturezaOperacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24216
            
    If IsMissing(vP1) Then
    
        Set objNaturezaOperacao = Nothing
        
    Else
    
        Set objNaturezaOperacao = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNaturezaOperacao.Trata_Parametros(objNaturezaOperacao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24217
            
        If lErro <> SUCESSO Then Error 24218
        
    End If
    
    X.objNaturezaOperacao.Show
            
    NaturezaOperacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NaturezaOperacao_Chama_Tela:

    NaturezaOperacao_Chama_Tela = Err

    Select Case Err
    
        Case 24216, 24217
            Unload X.objNaturezaOperacao
            
        Case 24218
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146967)
        
    End Select
        
    Exit Function

End Function

Function PrevVenda_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPrevVenda As ClassPrevVenda

On Error GoTo Erro_PrevVenda_Chama_Tela

    Load X.objPrevVenda
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29641
            
    If IsMissing(vP1) Then
    
        Set objPrevVenda = Nothing
        
    Else
    
        Set objPrevVenda = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPrevVenda.Trata_Parametros(objPrevVenda)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29642
            
        If lErro <> SUCESSO Then Error 29643
        
    End If
    
    X.objPrevVenda.Show
            
    PrevVenda_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PrevVenda_Chama_Tela:

    PrevVenda_Chama_Tela = Err

    Select Case Err
    
        Case 29641, 29642
            Unload X.objPrevVenda
            
        Case 29643
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146968)
        
    End Select
        
    Exit Function

End Function

Function NFiscal_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_NFiscal_Chama_Tela

    Load X.objNFiscal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16859
            
    If IsMissing(vP1) Then
    
        Set objNFiscal = Nothing
        
    Else
    
        Set objNFiscal = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscal.Trata_Parametros(objNFiscal)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16860
            
        If lErro <> SUCESSO Then Error 16861
        
    End If
    
    X.objNFiscal.Show
            
    NFiscal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscal_Chama_Tela:

    NFiscal_Chama_Tela = Err

    Select Case Err
    
        Case 16859, 16860
            Unload X.objNFiscal
            
        Case 16861
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146969)
        
    End Select
        
    Exit Function

End Function

Function NFiscalDev_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalDev As ClassNFiscal

On Error GoTo Erro_NFiscalDev_Chama_Tela

    Load X.objNFiscalDev
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19595
            
    If IsMissing(vP1) Then
    
        Set objNFiscalDev = Nothing
        
    Else
    
        Set objNFiscalDev = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalDev.Trata_Parametros(objNFiscalDev)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19596
            
        If lErro <> SUCESSO Then Error 19597
        
    End If
    
    X.objNFiscalDev.Show
            
    NFiscalDev_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalDev_Chama_Tela:

    NFiscalDev_Chama_Tela = Err

    Select Case Err
    
        Case 19595, 19596
            Unload X.objNFiscalDev
            
        Case 19597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146970)
        
    End Select
        
    Exit Function

End Function

Function NFiscalRem_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalRem As ClassNFiscal

On Error GoTo Erro_NFiscalRem_Chama_Tela

    Load X.objNFiscalRem
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19749
            
    If IsMissing(vP1) Then
    
        Set objNFiscalRem = Nothing
        
    Else
    
        Set objNFiscalRem = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalRem.Trata_Parametros(objNFiscalRem)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19750
            
        If lErro <> SUCESSO Then Error 19751
        
    End If
    
    X.objNFiscalRem.Show
            
    NFiscalRem_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalRem_Chama_Tela:

    NFiscalRem_Chama_Tela = Err

    Select Case Err
    
        Case 19749, 19750
            Unload X.objNFiscalRem
            
        Case 19751
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146971)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFatura_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFatura As ClassNFiscal

On Error GoTo Erro_NFiscalFatura_Chama_Tela

    Load X.objNFiscalFatura
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19774
            
    If IsMissing(vP1) Then
    
        Set objNFiscalFatura = Nothing
        
    Else
    
        Set objNFiscalFatura = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalFatura.Trata_Parametros(objNFiscalFatura)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19775
            
        If lErro <> SUCESSO Then Error 19776
        
    End If
    
    X.objNFiscalFatura.Show
            
    NFiscalFatura_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFatura_Chama_Tela:

    NFiscalFatura_Chama_Tela = Err

    Select Case Err
    
        Case 19774, 19775
            Unload X.objNFiscalFatura
            
        Case 19776
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146972)
        
    End Select
        
    Exit Function

End Function

Function ExcecoesICMS_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objExcecoesICMS As ClassICMSExcecao

On Error GoTo Erro_ExcecoesICMS_Chama_Tela

    Load X.objExcecoesICMS
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24501
            
    If IsMissing(vP1) Then
    
        Set objExcecoesICMS = Nothing
        
    Else
    
        Set objExcecoesICMS = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objExcecoesICMS.Trata_Parametros(objExcecoesICMS)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24502
            
        If lErro <> SUCESSO Then Error 24615
        
    End If
    
    X.objExcecoesICMS.Show
            
    ExcecoesICMS_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExcecoesICMS_Chama_Tela:

    ExcecoesICMS_Chama_Tela = Err

    Select Case Err
    
        Case 24501, 24502
            Unload X.objExcecoesICMS
            
        Case 24615
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146973)
        
    End Select
        
    Exit Function

End Function

Function ExcecoesIPI_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objExcecoesIPI As ClassIPIExcecao

On Error GoTo Erro_ExcecoesIPI_Chama_Tela

    Load X.objExcecoesIPI
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24616
            
    If IsMissing(vP1) Then
    
        Set objExcecoesIPI = Nothing
        
    Else
    
        Set objExcecoesIPI = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objExcecoesIPI.Trata_Parametros(objExcecoesIPI)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24617
            
        If lErro <> SUCESSO Then Error 24618
        
    End If
    
    X.objExcecoesIPI.Show
            
    ExcecoesIPI_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExcecoesIPI_Chama_Tela:

    ExcecoesIPI_Chama_Tela = Err

    Select Case Err
    
        Case 24616, 24617
            Unload X.objExcecoesIPI
            
        Case 24618
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146974)
        
    End Select
        
    Exit Function

End Function

Function LiberaBloqueio_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLiberaBloqueio As ClassPedidoDeVenda

On Error GoTo Erro_LiberaBloqueio_Chama_Tela

    Load X.objLiberaBloqueio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24399
            
    If IsMissing(vP1) Then
    
        Set objLiberaBloqueio = Nothing
        
    Else
    
        Set objLiberaBloqueio = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLiberaBloqueio.Trata_Parametros(objLiberaBloqueio)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24400
            
        If lErro <> SUCESSO Then Error 24401
        
    End If
    
    X.objLiberaBloqueio.Show
            
    LiberaBloqueio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LiberaBloqueio_Chama_Tela:

    LiberaBloqueio_Chama_Tela = Err

    Select Case Err
    
        Case 24399, 24400
            Unload X.objLiberaBloqueio
            
        Case 24401
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146975)
        
    End Select
        
    Exit Function

End Function

Function TabelaPrecoAlteracao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTabelaPrecoAlteracao As ClassTabelaPreco

On Error GoTo Erro_TabelaPrecoAlteracao_Chama_Tela

    Load X.objTabelaPrecoAlteracao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24230
            
    If IsMissing(vP1) Then
    
        Set objTabelaPrecoAlteracao = Nothing
        
    Else
    
        Set objTabelaPrecoAlteracao = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTabelaPrecoAlteracao.Trata_Parametros(objTabelaPrecoAlteracao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24231
            
        If lErro <> SUCESSO Then Error 24232
        
    End If
    
    X.objTabelaPrecoAlteracao.Show vbModal
            
    TabelaPrecoAlteracao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TabelaPrecoAlteracao_Chama_Tela:

    TabelaPrecoAlteracao_Chama_Tela = Err

    Select Case Err
    
        Case 24230, 24231
            Unload X.objTabelaPrecoAlteracao
            
        Case 24232
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146976)
        
    End Select
        
    Exit Function

End Function

Function TabelaPrecoCriacao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTabelaPrecoCriacao As ClassTabelaPreco

On Error GoTo Erro_TabelaPrecoCriacao_Chama_Tela

    Load X.objTabelaPrecoCriacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24233
            
    If IsMissing(vP1) Then
    
        Set objTabelaPrecoCriacao = Nothing
        
    Else
    
        Set objTabelaPrecoCriacao = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTabelaPrecoCriacao.Trata_Parametros(objTabelaPrecoCriacao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24234
            
        If lErro <> SUCESSO Then Error 24235
        
    End If
    
    X.objTabelaPrecoCriacao.Show vbModal
            
    TabelaPrecoCriacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TabelaPrecoCriacao_Chama_Tela:

    TabelaPrecoCriacao_Chama_Tela = Err

    Select Case Err
    
        Case 24233, 24234
            Unload X.objTabelaPrecoCriacao
            
        Case 24235
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146977)
        
    End Select
        
    Exit Function

End Function

Function TabelaPrecoItem_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As ClassTabelaPrecoItem

On Error GoTo Erro_TabelaPrecoItem_Chama_Tela

    Load X.objTabelaPrecoItem
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24236

    If IsMissing(vP1) Then

        Set objTabelaPrecoItem = Nothing

    Else

        Set objTabelaPrecoItem = vP1

    End If

    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then

        lErro = X.objTabelaPrecoItem.Trata_Parametros(objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24237

        If lErro <> SUCESSO Then Error 24238

    End If

    X.objTabelaPrecoItem.Show

    TabelaPrecoItem_Chama_Tela = SUCESSO

    Exit Function

Erro_TabelaPrecoItem_Chama_Tela:

    TabelaPrecoItem_Chama_Tela = Err

    Select Case Err

        Case 24236, 24237
            Unload X.objTabelaPrecoItem

        Case 24238

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146978)

    End Select

    Exit Function

End Function

Function TabelaPrecoItemEmpresaToda_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTabelaPrecoItemEmpresaToda As ClassTabelaPrecoItem

On Error GoTo Erro_TabelaPrecoItemEmpresaToda_Chama_Tela

    Load X.objTabelaPrecoItemEmpresaToda
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24359

    If IsMissing(vP1) Then

        Set objTabelaPrecoItemEmpresaToda = Nothing

    Else

        Set objTabelaPrecoItemEmpresaToda = vP1

    End If

    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then

        lErro = X.objTabelaPrecoItemEmpresaToda.Trata_Parametros(objTabelaPrecoItemEmpresaToda)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24360

        If lErro <> SUCESSO Then Error 24361

    End If

    X.objTabelaPrecoItemEmpresaToda.Show

    TabelaPrecoItemEmpresaToda_Chama_Tela = SUCESSO

    Exit Function

Erro_TabelaPrecoItemEmpresaToda_Chama_Tela:

    TabelaPrecoItemEmpresaToda_Chama_Tela = Err

    Select Case Err

        Case 24359, 24360
            Unload X.objTabelaPrecoItemEmpresaToda

        Case 24361

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146979)

    End Select

    Exit Function

End Function

Function CanalDeVenda_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCanalDeVenda As ClassCanalVenda

On Error GoTo Erro_CanalDeVenda_Chama_Tela

    Load X.objCanalDeVenda
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24242
            
    If IsMissing(vP1) Then
    
        Set objCanalDeVenda = Nothing
        
    Else
    
        Set objCanalDeVenda = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCanalDeVenda.Trata_Parametros(objCanalDeVenda)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24243
            
        If lErro <> SUCESSO Then Error 24244
        
    End If
    
    X.objCanalDeVenda.Show
            
    CanalDeVenda_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CanalDeVenda_Chama_Tela:

    CanalDeVenda_Chama_Tela = Err

    Select Case Err
    
        Case 24242, 24243
            Unload X.objCanalDeVenda
            
        Case 24244
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146980)
        
    End Select
        
    Exit Function

End Function

Function LiberacaoCreditoConfig_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLiberacaoCreditoConfig As ClassLiberacaoCredito

On Error GoTo Erro_LiberacaoCreditoConfig_Chama_Tela

    Load X.objLiberacaoCreditoConfig
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45021
            
    If IsMissing(vP1) Then
    
        Set objLiberacaoCreditoConfig = Nothing
        
    Else
    
        Set objLiberacaoCreditoConfig = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLiberacaoCreditoConfig.Trata_Parametros(objLiberacaoCreditoConfig)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45022
            
        If lErro <> SUCESSO Then Error 45023
        
    End If
    
    X.objLiberacaoCreditoConfig.Show
            
    LiberacaoCreditoConfig_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LiberacaoCreditoConfig_Chama_Tela:

    LiberacaoCreditoConfig_Chama_Tela = Err

    Select Case Err
    
        Case 45021, 45022
            Unload X.objLiberacaoCreditoConfig
            
        Case 45023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146981)
        
    End Select
        
    Exit Function

End Function

Function SerieNFiscal_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objSerieNFiscal As ClassSerie

On Error GoTo Erro_SerieNFiscal_Chama_Tela

    Load X.objSerieNFiscal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24245
            
    If IsMissing(vP1) Then
    
        Set objSerieNFiscal = Nothing
        
    Else
    
        Set objSerieNFiscal = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSerieNFiscal.Trata_Parametros(objSerieNFiscal)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24246
            
        If lErro <> SUCESSO Then Error 24247
        
    End If
    
    X.objSerieNFiscal.Show
            
    SerieNFiscal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SerieNFiscal_Chama_Tela:

    SerieNFiscal_Chama_Tela = Err

    Select Case Err
    
        Case 24245, 24246
            Unload X.objSerieNFiscal
            
        Case 24247
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146982)
        
    End Select
        
    Exit Function

End Function

Function TipoDeBloqueio_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoDeBloqueio As ClassTipoDeBloqueio

On Error GoTo Erro_TipoDeBloqueio_Chama_Tela

    Load X.objTipoDeBloqueio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24248
            
    If IsMissing(vP1) Then
    
        Set objTipoDeBloqueio = Nothing
        
    Else
    
        Set objTipoDeBloqueio = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoDeBloqueio.Trata_Parametros(objTipoDeBloqueio)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24249
            
        If lErro <> SUCESSO Then Error 24250
        
    End If
    
    X.objTipoDeBloqueio.Show
            
    TipoDeBloqueio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDeBloqueio_Chama_Tela:

    TipoDeBloqueio_Chama_Tela = Err

    Select Case Err
    
        Case 24248, 24249
            Unload X.objTipoDeBloqueio
            
        Case 24250
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146983)
        
    End Select
        
    Exit Function

End Function

Function PedidoVenda_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPedidoVenda As ClassPedidoDeVenda

On Error GoTo Erro_PedidoVenda_Chama_Tela

    Load X.objPedidoVenda
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24350
            
    If IsMissing(vP1) Then
    
        Set objPedidoVenda = Nothing
        
    Else
    
        Set objPedidoVenda = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPedidoVenda.Trata_Parametros(objPedidoVenda)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24351
            
        If lErro <> SUCESSO Then Error 24352
        
    End If
    
    X.objPedidoVenda.Show
            
    PedidoVenda_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoVenda_Chama_Tela:

    PedidoVenda_Chama_Tela = Err

    Select Case Err
    
        Case 24350, 24351
            Unload X.objPedidoVenda
            
        Case 24352
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146984)
        
    End Select
        
    Exit Function

End Function

Function Estados_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objEstados As ClassEstado

On Error GoTo Erro_Estados_Chama_Tela

    Load X.objEstados
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24503
            
    If IsMissing(vP1) Then
    
        Set objEstados = Nothing
        
    Else
    
        Set objEstados = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objEstados.Trata_Parametros(objEstados)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24504
            
        If lErro <> SUCESSO Then Error 24505
        
    End If
    
    X.objEstados.Show
            
    Estados_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Estados_Chama_Tela:

    Estados_Chama_Tela = Err

    Select Case Err
    
        Case 24503, 24504
            Unload X.objEstados
            
        Case 24505
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146985)
        
    End Select
        
    Exit Function

End Function

Function SubstProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objSubstProduto As ClassSubstProduto

On Error GoTo Erro_SubstProduto_Chama_Tela

    Load X.objSubstProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24403
            
    If IsMissing(vP1) Then
    
        Set objSubstProduto = Nothing
        
    Else
    
        Set objSubstProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSubstProduto.Trata_Parametros(objSubstProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24404
            
        If lErro <> SUCESSO Then Error 24405
        
    End If
    
    X.objSubstProduto.Show vbModal
            
    SubstProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SubstProduto_Chama_Tela:

    SubstProduto_Chama_Tela = Err

    Select Case Err
    
        Case 24403, 24404
            Unload X.objSubstProduto
            
        Case 24405
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146986)
        
    End Select
        
    Exit Function

End Function

Function SubstProdutoNF_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objSubstProdutoNF As ClassSubstProdutoNF

On Error GoTo Erro_SubstProdutoNF_Chama_Tela

    Load X.objSubstProdutoNF
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19601
            
    If IsMissing(vP1) Then
    
        Set objSubstProdutoNF = Nothing
        
    Else
    
        Set objSubstProdutoNF = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSubstProdutoNF.Trata_Parametros(objSubstProdutoNF)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19602
            
        If lErro <> SUCESSO Then Error 19603
        
    End If
    
    X.objSubstProdutoNF.Show vbModal
            
    SubstProdutoNF_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SubstProdutoNF_Chama_Tela:

    SubstProdutoNF_Chama_Tela = Err

    Select Case Err
    
        Case 19601, 19602
            Unload X.objSubstProdutoNF
            
        Case 19603
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146987)
        
    End Select
        
    Exit Function

End Function

Function AlocacaoProduto1_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAlocacaoProduto1 As ClassItemPedido

On Error GoTo Erro_AlocacaoProduto1_Chama_Tela

    Load X.objAlocacaoProduto1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24337
            
    If IsMissing(vP1) Then
    
        Set objAlocacaoProduto1 = Nothing
        
    Else
    
        Set objAlocacaoProduto1 = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAlocacaoProduto1.Trata_Parametros(objAlocacaoProduto1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24338
            
        If lErro <> SUCESSO Then Error 24339
        
    End If
    
    X.objAlocacaoProduto1.Show vbModal
            
    AlocacaoProduto1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlocacaoProduto1_Chama_Tela:

    AlocacaoProduto1_Chama_Tela = Err

    Select Case Err
    
        Case 24337, 24338
            Unload X.objAlocacaoProduto1
            
        Case 24339
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146988)
        
    End Select
        
    Exit Function

End Function

Function AlocacaoProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAlocacaoProduto As colItemPedido

On Error GoTo Erro_AlocacaoProduto_Chama_Tela

    Load X.objAlocacaoProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24392
            
    If IsMissing(vP1) Then
    
        Set objAlocacaoProduto = Nothing
        
    Else
    
        Set objAlocacaoProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAlocacaoProduto.Trata_Parametros(objAlocacaoProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24393
            
        If lErro <> SUCESSO Then Error 24394
        
    End If
    
    X.objAlocacaoProduto.Show vbModal
            
    AlocacaoProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlocacaoProduto_Chama_Tela:

    AlocacaoProduto_Chama_Tela = Err

    Select Case Err
    
        Case 24392, 24393
            Unload X.objAlocacaoProduto
            
        Case 24394
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146989)
        
    End Select
        
    Exit Function

End Function

Function AlocacaoProdutoSaida1_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objGenerico As AdmGenerico
'Dim objAlocacaoProdutoSaida1 As ClassItemPedido
'Dim objForm As Form

On Error GoTo Erro_AlocacaoProdutoSaida1_Chama_Tela

    Load X.objAlocacaoProdutoSaida1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24340
            
    If IsMissing(vP1) Then
    
        Set objGenerico = Nothing
        
    Else
    
        Set objGenerico = vP1
        
    End If
           
    lErro = X.objAlocacaoProdutoSaida1.Trata_Parametros(objGenerico)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24341
            
    If lErro <> SUCESSO Then Error 24342
    
    X.objAlocacaoProdutoSaida1.Show vbModal
            
    AlocacaoProdutoSaida1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlocacaoProdutoSaida1_Chama_Tela:

    AlocacaoProdutoSaida1_Chama_Tela = Err

    Select Case Err
    
        Case 24340, 24341
            Unload X.objAlocacaoProdutoSaida1
            
        Case 24342
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146990)
        
    End Select
        
    Exit Function

End Function

Function AlocacaoProdutoSaida_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objGenerico As AdmGenerico

'Dim objAlocacaoProdutoSaida As colItemPedido
'Dim objForm As Form
'Dim objGrid As AdmGrid
'Dim iListIndex As Integer

On Error GoTo Erro_AlocacaoProdutoSaida_Chama_Tela

    Load X.objAlocacaoProdutoSaida
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24396
            
    If IsMissing(vP1) Then
    
        Set objGenerico = Nothing
        
    Else
    
        Set objGenerico = vP1
        
    End If
           
    lErro = X.objAlocacaoProdutoSaida.Trata_Parametros(objGenerico)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24397
            
    If lErro <> SUCESSO Then Error 24398
    
    X.objAlocacaoProdutoSaida.Show vbModal
            
    AlocacaoProdutoSaida_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlocacaoProdutoSaida_Chama_Tela:

    AlocacaoProdutoSaida_Chama_Tela = Err

    Select Case Err
    
        Case 24396, 24397
            Unload X.objAlocacaoProdutoSaida
            
        Case 24398
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146991)
        
    End Select
        
    Exit Function

End Function

Function TabelaPrecoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTabelaPreco1 As ClassTabelaPreco
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TabelaPrecoLista_Chama_Tela

    Load X.objTabelaPrecoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24302
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objTabelaPreco1 = Nothing
        
    Else
    
        Set objTabelaPreco1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objTabelaPrecoLista.Trata_Parametros(colSelecao, objTabelaPreco1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24303
            
    If lErro <> SUCESSO Then Error 24304
        
    X.objTabelaPrecoLista.Show
            
    TabelaPrecoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TabelaPrecoLista_Chama_Tela:

    TabelaPrecoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24302, 24303
            Unload X.objTabelaPrecoLista
            
        Case 24304
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146992)
        
    End Select
        
    Exit Function

End Function

Function TabelaPrecoItemLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTabelaPrecoItem1 As ClassTabelaPrecoItem
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TabelaPrecoItemLista_Chama_Tela

    Load X.objTabelaPrecoItemLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24785
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTabelaPrecoItem1 = Nothing
        
    Else
    
        Set objTabelaPrecoItem1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTabelaPrecoItemLista.Trata_Parametros(colSelecao, objTabelaPrecoItem1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24786
            
    If lErro <> SUCESSO Then Error 24787
        
    X.objTabelaPrecoItemLista.Show
            
    TabelaPrecoItemLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TabelaPrecoItemLista_Chama_Tela:

    TabelaPrecoItemLista_Chama_Tela = Err

    Select Case Err
    
        Case 24785, 24786
            Unload X.objTabelaPrecoItemLista
            
        Case 24787
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146993)
        
    End Select
        
    Exit Function

End Function

Function PedidoVendaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoVenda1 As ClassPedidoDeVenda
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoVendaLista_Chama_Tela

    Load X.objPedidoVendaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24447
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPedidoVenda1 = Nothing
        
    Else
    
        Set objPedidoVenda1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedidoVendaLista.Trata_Parametros(colSelecao, objPedidoVenda1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24448
            
    If lErro <> SUCESSO Then Error 24449
        
    X.objPedidoVendaLista.Show
            
    PedidoVendaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoVendaLista_Chama_Tela:

    PedidoVendaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24447, 24448
            Unload X.objPedidoVendaLista
            
        Case 24449
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146994)
        
    End Select
        
    Exit Function

End Function

Function NFsRecLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long
'??? revisar
''Dim lErro As Long
''Dim objNFsRec1 As ClassNFsRec
''Dim objEvento As AdmEvento
''Dim colSelecao As Collection
''
''On Error GoTo Erro_NFsRecLista_Chama_Tela
''
''    Load X.objNFsRecLista
''    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24450
''
''    If IsMissing(vP1) Then
''
''        Set colSelecao = Nothing
''
''    Else
''
''        Set colSelecao = vP1
''
''    End If
''
''
''    If IsMissing(vP2) Then
''
''        Set objNFsRec1 = Nothing
''
''    Else
''
''        Set objNFsRec1 = vP2
''
''    End If
''
''    If IsMissing(vP3) Then
''
''        Set objEvento = Nothing
''
''
''    Else
''
''        Set objEvento = vP3
''
''    End If
''
''
''    lErro = X.objNFsRecLista.Trata_Parametros(colSelecao, objNFsRec1, objEvento)
''    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24451
''
''    If lErro <> SUCESSO Then Error 24452
''
''    X.objNFsRecLista.Show
''
''    NFsRecLista_Chama_Tela = SUCESSO
''
''    Exit Function
''
''Erro_NFsRecLista_Chama_Tela:
''
''    NFsRecLista_Chama_Tela = Err
''
''    Select Case Err
''
''        Case 24450, 24451
''            Unload X.objNFsRecLista
''
''        Case 24452
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146995)
''
''    End Select
''
''    Exit Function
''
End Function

Function NaturezaOperacaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNaturezaOperacao1 As ClassNaturezaOp
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NaturezaOperacaoLista_Chama_Tela

    Load X.objNaturezaOperacaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24515
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNaturezaOperacao1 = Nothing
        
    Else
    
        Set objNaturezaOperacao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNaturezaOperacaoLista.Trata_Parametros(colSelecao, objNaturezaOperacao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24516
            
    If lErro <> SUCESSO Then Error 24517
        
    X.objNaturezaOperacaoLista.Show
            
    NaturezaOperacaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NaturezaOperacaoLista_Chama_Tela:

    NaturezaOperacaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24515, 24516
            Unload X.objNaturezaOperacaoLista
            
        Case 24517
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146996)
        
    End Select
        
    Exit Function

End Function

Function NaturezaOpLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNaturezaOp1 As ClassNaturezaOp
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NaturezaOpLista_Chama_Tela

    Load X.objNaturezaOpLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29704
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNaturezaOp1 = Nothing
        
    Else
    
        Set objNaturezaOp1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNaturezaOpLista.Trata_Parametros(colSelecao, objNaturezaOp1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29705
            
    If lErro <> SUCESSO Then Error 29706
        
    X.objNaturezaOpLista.Show
            
    NaturezaOpLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NaturezaOpLista_Chama_Tela:

    NaturezaOpLista_Chama_Tela = Err

    Select Case Err
    
        Case 29704, 29705
            Unload X.objNaturezaOpLista
            
        Case 29706
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146997)
        
    End Select
        
    Exit Function

End Function

Function PrevVendaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPrevVenda1 As ClassPrevVenda
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PrevVendaLista_Chama_Tela

    Load X.objPrevVendaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29692
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPrevVenda1 = Nothing
        
    Else
    
        Set objPrevVenda1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objPrevVendaLista.Trata_Parametros(colSelecao, objPrevVenda1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29693
            
    If lErro <> SUCESSO Then Error 29694
        
    X.objPrevVendaLista.Show
            
    PrevVendaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PrevVendaLista_Chama_Tela:

    PrevVendaLista_Chama_Tela = Err

    Select Case Err
    
        Case 29692, 29693
            Unload X.objPrevVendaLista
            
        Case 29694
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146998)
        
    End Select
        
    Exit Function

End Function

Function SerieLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objSerie1 As ClassSerie
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_SerieLista_Chama_Tela

    Load X.objSerieLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24855
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objSerie1 = Nothing
        
    Else
    
        Set objSerie1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objSerieLista.Trata_Parametros(colSelecao, objSerie1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24865
            
    If lErro <> SUCESSO Then Error 24857
        
    X.objSerieLista.Show
            
    SerieLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SerieLista_Chama_Tela:

    SerieLista_Chama_Tela = Err

    Select Case Err
    
        Case 24855, 24856
            Unload X.objSerieLista
            
        Case 24857
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146999)
        
    End Select
        
    Exit Function

End Function

Function SerieListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objSerie1 As ClassSerie
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_SerieListaModal_Chama_Tela

    Load X.objSerieListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45109
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objSerie1 = Nothing
        
    Else
    
        Set objSerie1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objSerieListaModal.Trata_Parametros(colSelecao, objSerie1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45110
            
    If lErro <> SUCESSO Then Error 45111
        
    X.objSerieListaModal.Show vbModal
            
    SerieListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SerieListaModal_Chama_Tela:

    SerieListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 45109, 45110
            Unload X.objSerieListaModal
            
        Case 45111
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147000)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProdutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProduto1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueProdutoLista_Chama_Tela

    Load X.objEstoqueProdutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24531
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProduto1 = Nothing
        
    Else
    
        Set objEstoqueProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueProdutoLista.Trata_Parametros(colSelecao, objEstoqueProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24532
            
    If lErro <> SUCESSO Then Error 24533
        
    X.objEstoqueProdutoLista.Show
            
    EstoqueProdutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProdutoLista_Chama_Tela:

    EstoqueProdutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24531, 24532
            Unload X.objEstoqueProdutoLista
            
        Case 24533
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147001)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProdutoFilialLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProdutoFilial1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueProdutoFilialLista_Chama_Tela

    Load X.objEstoqueProdutoFilialLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16835
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProdutoFilial1 = Nothing
        
    Else
    
        Set objEstoqueProdutoFilial1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueProdutoFilialLista.Trata_Parametros(colSelecao, objEstoqueProdutoFilial1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16836
            
    If lErro <> SUCESSO Then Error 16837
        
    X.objEstoqueProdutoFilialLista.Show
            
    EstoqueProdutoFilialLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProdutoFilialLista_Chama_Tela:

    EstoqueProdutoFilialLista_Chama_Tela = Err

    Select Case Err
    
        Case 16835, 16836
            Unload X.objEstoqueProdutoFilialLista
            
        Case 16837
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147002)
        
    End Select
        
    Exit Function

End Function

Function NFiscalPedidoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalPedido As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalPedidoLista_Chama_Tela

    Load X.objNFiscalPedidoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45417
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalPedido = Nothing
        
    Else
    
        Set objNFiscalPedido = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalPedidoLista.Trata_Parametros(colSelecao, objNFiscalPedido, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45418
            
    If lErro <> SUCESSO Then Error 45419
        
    X.objNFiscalPedidoLista.Show
            
    NFiscalPedidoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalPedidoLista_Chama_Tela:

    NFiscalPedidoLista_Chama_Tela = Err

    Select Case Err
    
        Case 45417, 45418
            Unload X.objNFiscalPedidoLista
        
        Case 45419
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147003)
        
    End Select
        
    Exit Function

End Function

Function UsuarioLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objUsuarios As ClassUsuarios
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_UsuarioLista_Chama_Tela

    Load X.objUsuarioLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45636
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objUsuarios = Nothing
        
    Else
    
        Set objUsuarios = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objUsuarioLista.Trata_Parametros(colSelecao, objUsuarios, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45637
            
    If lErro <> SUCESSO Then Error 45638
        
    X.objUsuarioLista.Show
            
    UsuarioLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_UsuarioLista_Chama_Tela:

    UsuarioLista_Chama_Tela = Err

    Select Case Err
    
        Case 45636, 45637
            Unload X.objUsuarioLista
            
        Case 45638
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147004)
        
    End Select
        
    Exit Function

End Function

Function ItemPVProdutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItemPedido As ClassItemPedido
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ItemPVProdutoLista_Chama_Tela

    Load X.objItemPVProdutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45793
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objItemPedido = Nothing
        
    Else
    
        Set objItemPedido = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objItemPVProdutoLista.Trata_Parametros(colSelecao, objItemPedido, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45794
            
    If lErro <> SUCESSO Then Error 45795
        
    X.objItemPVProdutoLista.Show
            
    ItemPVProdutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ItemPVProdutoLista_Chama_Tela:

    ItemPVProdutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 45793, 45794
            Unload X.objItemPVProdutoLista
            
        Case 45795
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147005)
        
    End Select
        
    Exit Function

End Function

Function PedidoNFiscalLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoNFiscalLista_Chama_Tela

    Load X.objPedidoNFiscalLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45648
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedidoDeVenda = Nothing
        
    Else
    
        Set objPedidoDeVenda = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objPedidoNFiscalLista.Trata_Parametros(colSelecao, objPedidoDeVenda, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45649
            
    If lErro <> SUCESSO Then Error 45650
        
    X.objPedidoNFiscalLista.Show
            
    PedidoNFiscalLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoNFiscalLista_Chama_Tela:

    PedidoNFiscalLista_Chama_Tela = Err

    Select Case Err
    
        Case 45648, 45649
            Unload X.objPedidoNFiscalLista
            
        Case 45650
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147006)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFaturaPedidoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFaturaPedido As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalFaturaPedidoLista_Chama_Tela

    Load X.objNFiscalFaturaPedidoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45535
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalFaturaPedido = Nothing
        
    Else
    
        Set objNFiscalFaturaPedido = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalFaturaPedidoLista.Trata_Parametros(colSelecao, objNFiscalFaturaPedido, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45536
            
    If lErro <> SUCESSO Then Error 45537
        
    X.objNFiscalFaturaPedidoLista.Show
            
    NFiscalFaturaPedidoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFaturaPedidoLista_Chama_Tela:

    NFiscalFaturaPedidoLista_Chama_Tela = Err

    Select Case Err
    
        Case 45535, 45536
            Unload X.objNFiscalFaturaPedidoLista
            
        Case 45537
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147007)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProdutoFilialLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProdutoFilial1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueProdutoFilialLista_Consulta_Chama_Tela

    Load X.objEstoqueProdutoFilialLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16839
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProdutoFilial1 = Nothing
        
    Else
    
        Set objEstoqueProdutoFilial1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueProdutoFilialLista_Consulta.Trata_Parametros(colSelecao, objEstoqueProdutoFilial1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16840
            
    If lErro <> SUCESSO Then Error 16841
        
    X.objEstoqueProdutoFilialLista_Consulta.Show
            
    EstoqueProdutoFilialLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProdutoFilialLista_Consulta_Chama_Tela:

    EstoqueProdutoFilialLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 16839, 16840
            Unload X.objEstoqueProdutoFilialLista_Consulta
            
        Case 16841
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147008)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProdutoTerceirosLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProdutoTerceiros1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueProdutoTerceirosLista_Chama_Tela

    Load X.objEstoqueProdutoTerceirosLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16926
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProdutoTerceiros1 = Nothing
        
    Else
    
        Set objEstoqueProdutoTerceiros1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueProdutoTerceirosLista.Trata_Parametros(colSelecao, objEstoqueProdutoTerceiros1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16927
            
    If lErro <> SUCESSO Then Error 16928
        
    X.objEstoqueProdutoTerceirosLista.Show
            
    EstoqueProdutoTerceirosLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProdutoTerceirosLista_Chama_Tela:

    EstoqueProdutoTerceirosLista_Chama_Tela = Err

    Select Case Err
    
        Case 16926, 16927
            Unload X.objEstoqueProdutoTerceirosLista
            
        Case 16928
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147009)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProdutoTerceirosLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProdutoTerceiros1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueProdutoTerceirosLista_Consulta_Chama_Tela

    Load X.objEstoqueProdutoTerceirosLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16929
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProdutoTerceiros1 = Nothing
        
    Else
    
        Set objEstoqueProdutoTerceiros1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueProdutoTerceirosLista_Consulta.Trata_Parametros(colSelecao, objEstoqueProdutoTerceiros1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16930
            
    If lErro <> SUCESSO Then Error 16931
        
    X.objEstoqueProdutoTerceirosLista_Consulta.Show
            
    EstoqueProdutoTerceirosLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProdutoTerceirosLista_Consulta_Chama_Tela:

    EstoqueProdutoTerceirosLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 16929, 16930
            Unload X.objEstoqueProdutoTerceirosLista_Consulta
            
        Case 16931
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147010)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProdutoLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProduto1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueProdutoLista_Consulta_Chama_Tela

    Load X.objEstoqueProdutoLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24751
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProduto1 = Nothing
        
    Else
    
        Set objEstoqueProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueProdutoLista_Consulta.Trata_Parametros(colSelecao, objEstoqueProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24752
            
    If lErro <> SUCESSO Then Error 24753
        
    X.objEstoqueProdutoLista_Consulta.Show
            
    EstoqueProdutoLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProdutoLista_Consulta_Chama_Tela:

    EstoqueProdutoLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24751, 24752
            Unload X.objEstoqueProdutoLista_Consulta
            
        Case 24753
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147011)
        
    End Select
        
    Exit Function

End Function

Function ExcecoesICMSLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objExcecoesICMS1 As ClassICMSExcecao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ExcecoesICMSLista_Chama_Tela

    Load X.objExcecoesICMSLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24551
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objExcecoesICMS1 = Nothing
        
    Else
    
        Set objExcecoesICMS1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objExcecoesICMSLista.Trata_Parametros(colSelecao, objExcecoesICMS1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24552
            
    If lErro <> SUCESSO Then Error 24553
        
    X.objExcecoesICMSLista.Show
            
    ExcecoesICMSLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExcecoesICMSLista_Chama_Tela:

    ExcecoesICMSLista_Chama_Tela = Err

    Select Case Err
    
        Case 24551, 24552
            Unload X.objExcecoesICMSLista
            
        Case 24553
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147012)
        
    End Select
        
    Exit Function

End Function

Function ExcecoesIPILista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objExcecoesIPI1 As ClassIPIExcecao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ExcecoesIPILista_Chama_Tela

    Load X.objExcecoesIPILista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24554
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objExcecoesIPI1 = Nothing
        
    Else
    
        Set objExcecoesIPI1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objExcecoesIPILista.Trata_Parametros(colSelecao, objExcecoesIPI1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24555
            
    If lErro <> SUCESSO Then Error 24556
        
    X.objExcecoesIPILista.Show
            
    ExcecoesIPILista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExcecoesIPILista_Chama_Tela:

    ExcecoesIPILista_Chama_Tela = Err

    Select Case Err
    
        Case 24554, 24555
            Unload X.objExcecoesIPILista
            
        Case 24556
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147013)
        
    End Select
        
    Exit Function

End Function

Function AtualizacaoPreco_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizacaoPreco_Chama_Tela

    Load X.objAtualizacaoPreco
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24330
    
    X.objAtualizacaoPreco.Show
            
    AtualizacaoPreco_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AtualizacaoPreco_Chama_Tela:

    AtualizacaoPreco_Chama_Tela = Err

    Select Case Err
    
        Case 24330
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147014)
        
    End Select
        
    Exit Function

End Function

Function ConfiguraFAT_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ConfiguraFAT_Chama_Tela

    Load X.objConfiguraFAT
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29615
    
    X.objConfiguraFAT.Show
            
    ConfiguraFAT_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ConfiguraFAT_Chama_Tela:

    ConfiguraFAT_Chama_Tela = Err

    Select Case Err
    
        Case 29615
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147015)
        
    End Select
        
    Exit Function

End Function

Function BaixaPedido_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BaixaPedido_Chama_Tela

    Load X.objBaixaPedido
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29682
    
    X.objBaixaPedido.Show
            
    BaixaPedido_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaPedido_Chama_Tela:

    BaixaPedido_Chama_Tela = Err

    Select Case Err
    
        Case 29682
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147016)
        
    End Select
        
    Exit Function

End Function

Function CancelaNFiscal_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CancelaNFiscal_Chama_Tela

    Load X.objCancelaNFiscal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24353
    
    X.objCancelaNFiscal.Show
            
    CancelaNFiscal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CancelaNFiscal_Chama_Tela:

    CancelaNFiscal_Chama_Tela = Err

    Select Case Err
    
        Case 24353
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147017)
        
    End Select
        
    Exit Function

End Function

Function GeracaoFatura_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoFatura_Chama_Tela

    Load X.objGeracaoFatura
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24355
    
    X.objGeracaoFatura.Show
            
    GeracaoFatura_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoFatura_Chama_Tela:

    GeracaoFatura_Chama_Tela = Err

    Select Case Err
    
        Case 24355
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147018)
        
    End Select
        
    Exit Function

End Function

Function GeracaoNFiscal_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoNFiscal_Chama_Tela

    Load X.objGeracaoNFiscal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24356
    
    X.objGeracaoNFiscal.Show
            
    GeracaoNFiscal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoNFiscal_Chama_Tela:

    GeracaoNFiscal_Chama_Tela = Err

    Select Case Err
    
        Case 24356
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147019)
        
    End Select
        
    Exit Function

End Function

Function Comissoes_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Comissoes_Chama_Tela

    Load X.objComissoes
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24357
    
    X.objComissoes.Show
            
    Comissoes_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Comissoes_Chama_Tela:

    Comissoes_Chama_Tela = Err

    Select Case Err
    
        Case 24357
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147020)
        
    End Select
        
    Exit Function

End Function

Function RelOpContasCorrentes_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpContasCorrentes_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29906
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpContasCorrentes.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29907
    
    Load X.objRelOpContasCorrentes
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29908
            
    X.objRelOpContasCorrentes.Show
            
    RelOpContasCorrentes_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpContasCorrentes_Chama_Tela:

    RelOpContasCorrentes_Chama_Tela = Err

    Select Case Err
    
        Case 29906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29907
    
        Case 29908
            Unload X.objRelOpContasCorrentes
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147021)
        
    End Select
        
    Exit Function

End Function

Function RelOpDespacho_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDespacho_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29909
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDespacho.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29910
    
    Load X.objRelOpDespacho
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29911
            
    X.objRelOpDespacho.Show
            
    RelOpDespacho_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDespacho_Chama_Tela:

    RelOpDespacho_Chama_Tela = Err

    Select Case Err
    
        Case 29909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29910
    
        Case 29911
            Unload X.objRelOpDespacho
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147022)
        
    End Select
        
    Exit Function

End Function

Function RelOpEmissaoDupl_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpEmissaoDupl_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29912
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpEmissaoDupl.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29913
    
    Load X.objRelOpEmissaoDupl
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29914
            
    X.objRelOpEmissaoDupl.Show
            
    RelOpEmissaoDupl_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpEmissaoDupl_Chama_Tela:

    RelOpEmissaoDupl_Chama_Tela = Err

    Select Case Err
    
        Case 29912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29913
    
        Case 29914
            Unload X.objRelOpEmissaoDupl
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147023)
        
    End Select
        
    Exit Function

End Function

Function RelOpEmissaoFatura_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpEmissaoFatura_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29915
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpEmissaoFatura.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29916
    
    Load X.objRelOpEmissaoFatura
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29917
            
    X.objRelOpEmissaoFatura.Show
            
    RelOpEmissaoFatura_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpEmissaoFatura_Chama_Tela:

    RelOpEmissaoFatura_Chama_Tela = Err

    Select Case Err
    
        Case 29915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29916
    
        Case 29917
            Unload X.objRelOpEmissaoFatura
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147024)
        
    End Select
        
    Exit Function

End Function

Function RelOpEmissaoNFiscal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpEmissaoNFiscal_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29918
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpEmissaoNFiscal.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29919
    
    Load X.objRelOpEmissaoNFiscal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29920
            
    X.objRelOpEmissaoNFiscal.Show
            
    RelOpEmissaoNFiscal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpEmissaoNFiscal_Chama_Tela:

    RelOpEmissaoNFiscal_Chama_Tela = Err

    Select Case Err
    
        Case 29918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29919
    
        Case 29920
            Unload X.objRelOpEmissaoNFiscal
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147025)
        
    End Select
        
    Exit Function

End Function

Function RelOpEstoqueVendas_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpEstoqueVendas_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29921
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpEstoqueVendas.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29922
    
    Load X.objRelOpEstoqueVendas
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29923
            
    X.objRelOpEstoqueVendas.Show
            
    RelOpEstoqueVendas_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpEstoqueVendas_Chama_Tela:

    RelOpEstoqueVendas_Chama_Tela = Err

    Select Case Err
    
        Case 29921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29922
    
        Case 29923
            Unload X.objRelOpEstoqueVendas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147026)
        
    End Select
        
    Exit Function

End Function

Function RelOpPrecos_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPrecos_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29599
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPrecos.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29600
    
    Load X.objRelOpPrecos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29601
            
    X.objRelOpPrecos.Show
            
    RelOpPrecos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPrecos_Chama_Tela:

    RelOpPrecos_Chama_Tela = Err

    Select Case Err
    
        Case 29599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29600
    
        Case 29601
            Unload X.objRelOpPrecos
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147027)
        
    End Select
        
    Exit Function

End Function

Function RelOpFatCliente_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpFatCliente_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29924
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpFatCliente.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29925
    
    Load X.objRelOpFatCliente
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29926
            
    X.objRelOpFatCliente.Show
            
    RelOpFatCliente_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpFatCliente_Chama_Tela:

    RelOpFatCliente_Chama_Tela = Err

    Select Case Err
    
        Case 29924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29925
    
        Case 29926
            Unload X.objRelOpFatCliente
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147028)
        
    End Select
        
    Exit Function

End Function

Function RelOpFatClienteProd_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpFatClienteProd_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29928
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpFatClienteProd.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29929
    
    Load X.objRelOpFatClienteProd
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29930
            
    X.objRelOpFatClienteProd.Show
            
    RelOpFatClienteProd_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpFatClienteProd_Chama_Tela:

    RelOpFatClienteProd_Chama_Tela = Err

    Select Case Err
    
        Case 29928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29929
    
        Case 29930
            Unload X.objRelOpFatClienteProd
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147029)
        
    End Select
        
    Exit Function

End Function

Function RelOpFatComissoes_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpFatComissoes_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29931
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpFatComissoes.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29932
    
    Load X.objRelOpFatComissoes
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29933
            
    X.objRelOpFatComissoes.Show
            
    RelOpFatComissoes_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpFatComissoes_Chama_Tela:

    RelOpFatComissoes_Chama_Tela = Err

    Select Case Err
    
        Case 29931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29932
    
        Case 29933
            Unload X.objRelOpFatComissoes
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147030)
        
    End Select
        
    Exit Function

End Function

Function RelOpFatVendedor_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpFatVendedor_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29934
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpFatVendedor.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29935
    
    Load X.objRelOpFatVendedor
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29936
            
    X.objRelOpFatVendedor.Show
            
    RelOpFatVendedor_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpFatVendedor_Chama_Tela:

    RelOpFatVendedor_Chama_Tela = Err

    Select Case Err
    
        Case 29934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29935
    
        Case 29936
            Unload X.objRelOpFatVendedor
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147031)
        
    End Select
        
    Exit Function

End Function

Function RelOpNFiscalDev_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpNFiscalDev_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29937
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpNFiscalDev.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29938
    
    Load X.objRelOpNFiscalDev
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29939
            
    X.objRelOpNFiscalDev.Show
            
    RelOpNFiscalDev_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpNFiscalDev_Chama_Tela:

    RelOpNFiscalDev_Chama_Tela = Err

    Select Case Err
    
        Case 29937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29938
    
        Case 29939
            Unload X.objRelOpNFiscalDev
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147032)
        
    End Select
        
    Exit Function

End Function

Function RelOpNFiscalTransp_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpNFiscalTransp_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29940
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpNFiscalTransp.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29941
    
    Load X.objRelOpNFiscalTransp
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29942
            
    X.objRelOpNFiscalTransp.Show
            
    RelOpNFiscalTransp_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpNFiscalTransp_Chama_Tela:

    RelOpNFiscalTransp_Chama_Tela = Err

    Select Case Err
    
        Case 29940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29941
    
        Case 29942
            Unload X.objRelOpNFiscalTransp
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147033)
        
    End Select
        
    Exit Function

End Function

Function RelOpNotasFiscais_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpNotasFiscais_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29943
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
            
    End If
    
    lErro = X.objRelOpNotasFiscais.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29944
    
    Load X.objRelOpNotasFiscais
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29945
            
    X.objRelOpNotasFiscais.Show
            
    RelOpNotasFiscais_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpNotasFiscais_Chama_Tela:

    RelOpNotasFiscais_Chama_Tela = Err

    Select Case Err
    
        Case 29943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29944
    
        Case 29945
            Unload X.objRelOpNotasFiscais
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147034)
        
    End Select
        
    Exit Function

End Function

Function RelOpPedAptosFat_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPedAptosFat_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29946
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPedAptosFat.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29947
    
    Load X.objRelOpPedAptosFat
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29948
            
    X.objRelOpPedAptosFat.Show
            
    RelOpPedAptosFat_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPedAptosFat_Chama_Tela:

    RelOpPedAptosFat_Chama_Tela = Err

    Select Case Err
    
        Case 29946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29947
    
        Case 29948
            Unload X.objRelOpPedAptosFat
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147035)
        
    End Select
        
    Exit Function

End Function

Function RelOpPedNaoEntregue_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPedNaoEntregue_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29949
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPedNaoEntregue.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29950
    
    Load X.objRelOpPedNaoEntregue
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29951
            
    X.objRelOpPedNaoEntregue.Show
            
    RelOpPedNaoEntregue_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPedNaoEntregue_Chama_Tela:

    RelOpPedNaoEntregue_Chama_Tela = Err

    Select Case Err
    
        Case 29949
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29950
    
        Case 29951
            Unload X.objRelOpPedNaoEntregue
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147036)
        
    End Select
        
    Exit Function

End Function

Function RelOpPedProduto_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPedProduto_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29952
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPedProduto.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29953
    
    Load X.objRelOpPedProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29954
            
    X.objRelOpPedProduto.Show
            
    RelOpPedProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPedProduto_Chama_Tela:

    RelOpPedProduto_Chama_Tela = Err

    Select Case Err
    
        Case 29952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29953
    
        Case 29954
            Unload X.objRelOpPedProduto
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147037)
        
    End Select
        
    Exit Function

End Function

Function RelOpPedVendedorCli_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPedVendedorCli_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29955
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPedVendedorCli.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29956
    
    Load X.objRelOpPedVendedorCli
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29957
            
    X.objRelOpPedVendedorCli.Show
            
    RelOpPedVendedorCli_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPedVendedorCli_Chama_Tela:

    RelOpPedVendedorCli_Chama_Tela = Err

    Select Case Err
    
        Case 29955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29956
    
        Case 29957
            Unload X.objRelOpPedVendedorCli
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147038)
        
    End Select
        
    Exit Function

End Function

Function RelOpPedVendedorProd_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPedVendedorProd_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29958
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPedVendedorProd.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29959
    
    Load X.objRelOpPedVendedorProd
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29960
            
    X.objRelOpPedVendedorProd.Show
            
    RelOpPedVendedorProd_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPedVendedorProd_Chama_Tela:

    RelOpPedVendedorProd_Chama_Tela = Err

    Select Case Err
    
        Case 29958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29959
    
        Case 29960
            Unload X.objRelOpPedVendedorProd
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147039)
        
    End Select
        
    Exit Function

End Function

Function RelOpPlanReajuste_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPlanReajuste_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29961
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPlanReajuste.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29962
    
    Load X.objRelOpPlanReajuste
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29963
            
    X.objRelOpPlanReajuste.Show
            
    RelOpPlanReajuste_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPlanReajuste_Chama_Tela:

    RelOpPlanReajuste_Chama_Tela = Err

    Select Case Err
    
        Case 29961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29962
    
        Case 29963
            Unload X.objRelOpPlanReajuste
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147040)
        
    End Select
        
    Exit Function

End Function

Function RelOpPrazoPagto_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPrazoPagto_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29964
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPrazoPagto.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29965
    
    Load X.objRelOpPrazoPagto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29966
            
    X.objRelOpPrazoPagto.Show
            
    RelOpPrazoPagto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPrazoPagto_Chama_Tela:

    RelOpPrazoPagto_Chama_Tela = Err

    Select Case Err
    
        Case 29964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29965
    
        Case 29966
            Unload X.objRelOpPrazoPagto
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147041)
        
    End Select
        
    Exit Function

End Function

Function RelOpPreNota_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPreNota_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29967
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPreNota.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29968
    
    Load X.objRelOpPreNota
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29969
            
    X.objRelOpPreNota.Show
            
    RelOpPreNota_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPreNota_Chama_Tela:

    RelOpPreNota_Chama_Tela = Err

    Select Case Err
    
        Case 29967
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29968
    
        Case 29969
            Unload X.objRelOpPreNota
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147042)
        
    End Select
        
    Exit Function

End Function

Function RelOpRealPrev_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRealPrev_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29970
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRealPrev.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29971
    
    Load X.objRelOpRealPrev
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29972
            
    X.objRelOpRealPrev.Show
            
    RelOpRealPrev_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRealPrev_Chama_Tela:

    RelOpRealPrev_Chama_Tela = Err

    Select Case Err
    
        Case 29970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29971
    
        Case 29972
            Unload X.objRelOpRealPrev
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147043)
        
    End Select
        
    Exit Function

End Function

Function RelOpResVendas_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpResVendas_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29973
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpResVendas.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29974
    
    Load X.objRelOpResVendas
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29975
            
    X.objRelOpResVendas.Show
            
    RelOpResVendas_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpResVendas_Chama_Tela:

    RelOpResVendas_Chama_Tela = Err

    Select Case Err
    
        Case 29973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29974
    
        Case 29975
            Unload X.objRelOpResVendas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147044)
        
    End Select
        
    Exit Function

End Function

Function RelOpTransportadoras_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpTransportadoras_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29976
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpTransportadoras.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29977
    
    Load X.objRelOpTransportadoras
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29978
            
    X.objRelOpTransportadoras.Show
            
    RelOpTransportadoras_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpTransportadoras_Chama_Tela:

    RelOpTransportadoras_Chama_Tela = Err

    Select Case Err
    
        Case 29976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29977
    
        Case 29978
            Unload X.objRelOpTransportadoras
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147045)
        
    End Select
        
    Exit Function

End Function

Function RelOpVendedores_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpVendedores_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29979
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpVendedores.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29980
    
    Load X.objRelOpVendedores
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29981
            
    X.objRelOpVendedores.Show
            
    RelOpVendedores_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpVendedores_Chama_Tela:

    RelOpVendedores_Chama_Tela = Err

    Select Case Err
    
        Case 29979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29980
    
        Case 29981
            Unload X.objRelOpVendedores
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147046)
        
    End Select
        
    Exit Function

End Function

Function RelOpVisitas_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpVisitas_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29982
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpVisitas.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29983
    
    Load X.objRelOpVisitas
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29984
            
    X.objRelOpVisitas.Show
            
    RelOpVisitas_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpVisitas_Chama_Tela:

    RelOpVisitas_Chama_Tela = Err

    Select Case Err
    
        Case 29982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29983
    
        Case 29984
            Unload X.objRelOpVisitas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147047)
        
    End Select
        
    Exit Function

End Function

Function RelOpFatPickList_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpFatPickList_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45216
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpFatPickList.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45217
    
    Load X.objRelOpFatPickList
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45218
            
    X.objRelOpFatPickList.Show
            
    RelOpFatPickList_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpFatPickList_Chama_Tela:

    RelOpFatPickList_Chama_Tela = Err

    Select Case Err
    
        Case 45216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45217
    
        Case 45218
            Unload X.objRelOpFatPickList
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147048)
        
    End Select
        
    Exit Function

End Function

Function RelOpReciboDespacho_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpReciboDespacho_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45219
    
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpReciboDespacho.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45220
    
    Load X.objRelOpReciboDespacho
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45221
            
    X.objRelOpReciboDespacho.Show
            
    RelOpReciboDespacho_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpReciboDespacho_Chama_Tela:

    RelOpReciboDespacho_Chama_Tela = Err

    Select Case Err
    
        Case 45219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45220
    
        Case 45221
            Unload X.objRelOpReciboDespacho
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147049)
        
    End Select
        
    Exit Function

End Function

Function PedidoVendaListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoVenda1 As ClassPedidoDeVenda
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoVendaListaModal_Chama_Tela

    Load X.objPedidoVendaListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16892
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedidoVenda1 = Nothing
        
    Else
    
        Set objPedidoVenda1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objPedidoVendaListaModal.Trata_Parametros(colSelecao, objPedidoVenda1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16893
            
    If lErro <> SUCESSO Then Error 16894
        
    X.objPedidoVendaListaModal.Show vbModal
            
    PedidoVendaListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoVendaListaModal_Chama_Tela:

    PedidoVendaListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 16892, 16893
            Unload X.objPedidoVendaListaModal
            
        Case 16894
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147050)
        
    End Select
        
    Exit Function

End Function

Function NFiscalRemLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalRem1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalRemLista_Chama_Tela

    Load X.objNFiscalRemLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19515
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalRem1 = Nothing
        
    Else
    
        Set objNFiscalRem1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalRemLista.Trata_Parametros(colSelecao, objNFiscalRem1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19516
            
    If lErro <> SUCESSO Then Error 19517
        
    X.objNFiscalRemLista.Show
            
    NFiscalRemLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalRemLista_Chama_Tela:

    NFiscalRemLista_Chama_Tela = Err

    Select Case Err
    
        Case 19515, 19516
            Unload X.objNFiscalRemLista
            
        Case 19517
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147051)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFaturaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFatura As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalFaturaLista_Chama_Tela

    Load X.objNFiscalFaturaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19798
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalFatura = Nothing
        
    Else
    
        Set objNFiscalFatura = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalFaturaLista.Trata_Parametros(colSelecao, objNFiscalFatura, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19799
            
    If lErro <> SUCESSO Then Error 19800
        
    X.objNFiscalFaturaLista.Show
            
    NFiscalFaturaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFaturaLista_Chama_Tela:

    NFiscalFaturaLista_Chama_Tela = Err

    Select Case Err
    
        Case 19798, 19799
            Unload X.objNFiscalFaturaLista
            
        Case 19800
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147052)
        
    End Select
        
    Exit Function

End Function

Function NFiscalDevLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalDev1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalDevLista_Chama_Tela

    Load X.objNFiscalDevLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19518
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalDev1 = Nothing
        
    Else
    
        Set objNFiscalDev1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalDevLista.Trata_Parametros(colSelecao, objNFiscalDev1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19519
            
    If lErro <> SUCESSO Then Error 19520
        
    X.objNFiscalDevLista.Show
            
    NFiscalDevLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalDevLista_Chama_Tela:

    NFiscalDevLista_Chama_Tela = Err

    Select Case Err
    
        Case 19518, 19519
            Unload X.objNFiscalDevLista
            
        Case 19520
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147053)
        
    End Select
        
    Exit Function

End Function

Function NF_NFFaturaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNF_NFFatura As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NF_NFFaturaLista_Chama_Tela

    Load X.objNF_NFFaturaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45467
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNF_NFFatura = Nothing
        
    Else
    
        Set objNF_NFFatura = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNF_NFFaturaLista.Trata_Parametros(colSelecao, objNF_NFFatura, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45468
            
    If lErro <> SUCESSO Then Error 45469
        
    X.objNF_NFFaturaLista.Show
            
    NF_NFFaturaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NF_NFFaturaLista_Chama_Tela:

    NF_NFFaturaLista_Chama_Tela = Err

    Select Case Err
    
        Case 45467, 45468
            Unload X.objNF_NFFaturaLista
            
        Case 45469
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147054)
        
    End Select
        
    Exit Function

End Function

Function TabPrecoItemLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TabPrecoItemLista_Chama_Tela

    Load X.objTabPrecoItemLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45504
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTabelaPrecoItem = Nothing
        
    Else
    
        Set objTabelaPrecoItem = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTabPrecoItemLista.Trata_Parametros(colSelecao, objTabelaPrecoItem, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45505
            
    If lErro <> SUCESSO Then Error 45506
        
    X.objTabPrecoItemLista.Show
            
    TabPrecoItemLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TabPrecoItemLista_Chama_Tela:

    TabPrecoItemLista_Chama_Tela = Err

    Select Case Err
    
        Case 45504, 45505
            Unload X.objTabPrecoItemLista
            
        Case 45506
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147055)
        
    End Select
        
    Exit Function

End Function

Function NFiscalLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscal1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalLista_Chama_Tela

    Load X.objNFiscalLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19556
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscal1 = Nothing
        
    Else
    
        Set objNFiscal1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalLista.Trata_Parametros(colSelecao, objNFiscal1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19557
            
    If lErro <> SUCESSO Then Error 19558
        
    X.objNFiscalLista.Show
            
    NFiscalLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalLista_Chama_Tela:

    NFiscalLista_Chama_Tela = Err

    Select Case Err
    
        Case 19556, 19557
            Unload X.objNFiscalLista
            
        Case 19558
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147056)
        
    End Select
        
    Exit Function

End Function

Function NFiscalNormalLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalNormal1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalNormalLista_Chama_Tela

    Load X.objNFiscalNormalLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19559
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalNormal1 = Nothing
        
    Else
    
        Set objNFiscalNormal1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalNormalLista.Trata_Parametros(colSelecao, objNFiscalNormal1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19560
            
    If lErro <> SUCESSO Then Error 19561
        
    X.objNFiscalNormalLista.Show
            
    NFiscalNormalLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalNormalLista_Chama_Tela:

    NFiscalNormalLista_Chama_Tela = Err

    Select Case Err
    
        Case 19559, 19560
            Unload X.objNFiscalNormalLista
            
        Case 19561
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147057)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFaturaNormalLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFaturaNormal As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalFaturaNormalLista_Chama_Tela

    Load X.objNFiscalFaturaNormalLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19838
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalFaturaNormal = Nothing
        
    Else
    
        Set objNFiscalFaturaNormal = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalFaturaNormalLista.Trata_Parametros(colSelecao, objNFiscalFaturaNormal, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19839
            
    If lErro <> SUCESSO Then Error 19840
        
    X.objNFiscalFaturaNormalLista.Show
            
    NFiscalFaturaNormalLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFaturaNormalLista_Chama_Tela:

    NFiscalFaturaNormalLista_Chama_Tela = Err

    Select Case Err
    
        Case 19838, 19839
            Unload X.objNFiscalFaturaNormalLista
            
        Case 19840
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147058)
        
    End Select
        
    Exit Function

End Function

Function TipoDocumentoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoDocumentoLista_Chama_Tela

    Load X.objTipoDocumentoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19659
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoDocumento = Nothing
        
    Else
    
        Set objTipoDocumento = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTipoDocumentoLista.Trata_Parametros(colSelecao, objTipoDocumento, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19660
            
    If lErro <> SUCESSO Then Error 19661
        
    X.objTipoDocumentoLista.Show
            
    TipoDocumentoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDocumentoLista_Chama_Tela:

    TipoDocumentoLista_Chama_Tela = Err

    Select Case Err
    
        Case 19659, 19660
            Unload X.objTipoDocumentoLista
            
        Case 19661
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147059)
        
    End Select
        
    Exit Function

End Function

Function TipoDocTituloRecLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoDocTituloRecLista_Chama_Tela

    Load X.objTipoDocTituloRecLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45014
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoDocumento = Nothing
        
    Else
    
        Set objTipoDocumento = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTipoDocTituloRecLista.Trata_Parametros(colSelecao, objTipoDocumento, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45015
            
    If lErro <> SUCESSO Then Error 45016
        
    X.objTipoDocTituloRecLista.Show
            
    TipoDocTituloRecLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDocTituloRecLista_Chama_Tela:

    TipoDocTituloRecLista_Chama_Tela = Err

    Select Case Err
    
        Case 45014, 45015
            Unload X.objTipoDocTituloRecLista
            
        Case 45016
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147060)
        
    End Select
        
    Exit Function

End Function

Function TipoDocCreditoPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoDocCreditoPagLista_Chama_Tela

    Load X.objTipoDocCreditoPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45277
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoDocumento = Nothing
        
    Else
    
        Set objTipoDocumento = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTipoDocCreditoPagLista.Trata_Parametros(colSelecao, objTipoDocumento, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45278
            
    If lErro <> SUCESSO Then Error 45279
        
    X.objTipoDocCreditoPagLista.Show
            
    TipoDocCreditoPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDocCreditoPagLista_Chama_Tela:

    TipoDocCreditoPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 45277, 45278
            Unload X.objTipoDocCreditoPagLista
            
        Case 45279
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147061)
        
    End Select
        
    Exit Function

End Function

Function TipoDocOutrosPagLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoDocOutrosPagLista_Chama_Tela

    Load X.objTipoDocOutrosPagLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45017
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoDocumento = Nothing
        
    Else
    
        Set objTipoDocumento = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTipoDocOutrosPagLista.Trata_Parametros(colSelecao, objTipoDocumento, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45018
            
    If lErro <> SUCESSO Then Error 45019
        
    X.objTipoDocOutrosPagLista.Show
            
    TipoDocOutrosPagLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDocOutrosPagLista_Chama_Tela:

    TipoDocOutrosPagLista_Chama_Tela = Err

    Select Case Err
    
        Case 45017, 45018
            Unload X.objTipoDocOutrosPagLista
            
        Case 45019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147062)
        
    End Select
        
    Exit Function

End Function

Function TiposDocBaixaPagCancelarLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TiposDocBaixaPagCancelarLista_Chama_Tela

    Load X.objTiposDocBaixaPagCancelarLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19930
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoDocumento = Nothing
        
    Else
    
        Set objTipoDocumento = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTiposDocBaixaPagCancelarLista.Trata_Parametros(colSelecao, objTipoDocumento, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19931
            
    If lErro <> SUCESSO Then Error 19932
        
    X.objTiposDocBaixaPagCancelarLista.Show
            
    TiposDocBaixaPagCancelarLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TiposDocBaixaPagCancelarLista_Chama_Tela:

    TiposDocBaixaPagCancelarLista_Chama_Tela = Err

    Select Case Err
    
        Case 19930, 19931
            Unload X.objTiposDocBaixaPagCancelarLista
            
        Case 19932
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147063)
        
    End Select
        
    Exit Function

End Function

Function TipoDocInfoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoDocInfo As ClassTipoDocInfo
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoDocInfoLista_Chama_Tela

    Load X.objTipoDocInfoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19662
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoDocInfo = Nothing
        
    Else
    
        Set objTipoDocInfo = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTipoDocInfoLista.Trata_Parametros(colSelecao, objTipoDocInfo, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19663
            
    If lErro <> SUCESSO Then Error 19664
        
    X.objTipoDocInfoLista.Show
            
    TipoDocInfoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDocInfoLista_Chama_Tela:

    TipoDocInfoLista_Chama_Tela = Err

    Select Case Err
    
        Case 19662, 19663
            Unload X.objTipoDocInfoLista
            
        Case 19664
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147064)
        
    End Select
        
    Exit Function

End Function

Function LocalizacaoProduto_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
'Dim objItemNF As ClassItemNF
'Dim dQuantAlocar As Double
'Dim colOutros As Collection
Dim colItensNF As colItensNF

On Error GoTo Erro_LocalizacaoProduto_Chama_Tela

    Load X.objLocalizacaoProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19604
            
    If IsMissing(vP1) Then
    
        Set colItensNF = Nothing
        
    Else
    
        Set colItensNF = vP1
        
    End If
  
    lErro = X.objLocalizacaoProduto.Trata_Parametros(colItensNF)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19605
            
    If lErro <> SUCESSO Then Error 19606
        
    X.objLocalizacaoProduto.Show vbModal
            
    LocalizacaoProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LocalizacaoProduto_Chama_Tela:

    LocalizacaoProduto_Chama_Tela = Err

    Select Case Err
    
        Case 19604, 19605
            Unload X.objLocalizacaoProduto
            
        Case 19606
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147065)
        
    End Select
        
    Exit Function

End Function

Function LocalizacaoProduto1_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim dQuantAlocar As Double
Dim colOutros As Collection

On Error GoTo Erro_LocalizacaoProduto1_Chama_Tela

    Load X.objLocalizacaoProduto1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19607
            
    If IsMissing(vP1) Then
    
        Set objItemNF = Nothing
        
    Else
    
        Set objItemNF = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set colOutros = Nothing
        
    Else
    
        Set colOutros = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        dQuantAlocar = 0
            
    Else
    
        dQuantAlocar = vP3
        
    End If
    
    lErro = X.objLocalizacaoProduto1.Trata_Parametros(objItemNF, colOutros, dQuantAlocar)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19608
            
    If lErro <> SUCESSO Then Error 19609
        
    X.objLocalizacaoProduto1.Show vbModal
            
    LocalizacaoProduto1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LocalizacaoProduto1_Chama_Tela:

    LocalizacaoProduto1_Chama_Tela = Err

    Select Case Err
    
        Case 19607, 19608
            Unload X.objLocalizacaoProduto1
            
        Case 19609
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147066)
        
    End Select
        
    Exit Function

End Function

Function AlcadaFat_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAlcada As ClassAlcada

On Error GoTo Erro_AlcadaFat_Chama_Tela

    Load X.objAlcadaFat
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44447
            
    If IsMissing(vP1) Then
    
        Set objAlcada = Nothing
        
    Else
    
        Set objAlcada = vP1
        
    End If
            
    lErro = X.objAlcadaFat.Trata_Parametros(objAlcada)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44448
            
    If lErro <> SUCESSO Then Error 44449
        
    X.objAlcadaFat.Show
            
    AlcadaFat_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlcadaFat_Chama_Tela:

    AlcadaFat_Chama_Tela = Err

    Select Case Err
    
        Case 44447
        
        Case 44448, 44449
            Unload X.objAlcadaFat
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147067)
        
    End Select
        
    Exit Function

End Function


Function AutorizacaoCredito_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAutorizacaoCredito As ClassAutorizacaoCredito

On Error GoTo Erro_AutorizacaoCredito_Chama_Tela

    Load X.objAutorizacaoCredito
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44450
            
    If IsMissing(vP1) Then
    
        Set objAutorizacaoCredito = Nothing
        
    Else
    
        Set objAutorizacaoCredito = vP1
        
    End If
            
    lErro = X.objAutorizacaoCredito.Trata_Parametros(objAutorizacaoCredito)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44451
            
    If lErro <> SUCESSO Then Error 44452
        
    X.objAutorizacaoCredito.Show vbModal
            
    AutorizacaoCredito_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AutorizacaoCredito_Chama_Tela:

    AutorizacaoCredito_Chama_Tela = Err

    Select Case Err
    
        Case 44450
        
        Case 44451, 44452
            Unload X.objAutorizacaoCredito
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147068)
        
    End Select
        
    Exit Function

End Function

Function NFiscalPedido_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFiscalPedido_Chama_Tela

    Load X.objNFiscalPedido
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44636
    
    lErro = X.objNFiscalPedido.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44637
            
    If lErro <> SUCESSO Then Error 44638
        
    X.objNFiscalPedido.Show
            
    NFiscalPedido_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalPedido_Chama_Tela:

    NFiscalPedido_Chama_Tela = Err

    Select Case Err
    
        Case 44636
        
        Case 44637, 44638
            Unload X.objNFiscalPedido
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147069)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFaturaPedido_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFiscalFaturaPedido_Chama_Tela

    Load X.objNFiscalFaturaPedido
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44640
    
    lErro = X.objNFiscalFaturaPedido.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44641
            
    If lErro <> SUCESSO Then Error 44642
        
    X.objNFiscalFaturaPedido.Show
            
    NFiscalFaturaPedido_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFaturaPedido_Chama_Tela:

    NFiscalFaturaPedido_Chama_Tela = Err

    Select Case Err
    
        Case 44640
        
        Case 44641, 44642
            Unload X.objNFiscalFaturaPedido
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147070)
        
    End Select
        
    Exit Function

End Function


