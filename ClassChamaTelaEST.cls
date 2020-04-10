VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaTelaEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Almoxarifado_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifado_Chama_Tela

    Load X.objAlmoxarifado
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24276
            
    If IsMissing(vP1) Then
    
        Set objAlmoxarifado = Nothing
        
    Else
    
        Set objAlmoxarifado = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAlmoxarifado.Trata_Parametros(objAlmoxarifado)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24277
            
        If lErro <> SUCESSO Then Error 24278
        
    End If
    
    X.objAlmoxarifado.Show
            
    Almoxarifado_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Almoxarifado_Chama_Tela:

    Almoxarifado_Chama_Tela = Err

    Select Case Err
    
        Case 24276, 24277
            Unload X.objAlmoxarifado
            
        Case 24278
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146849)
        
    End Select
        
    Exit Function

End Function

Function MovEstoque_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objMovEstoque As ClassMovEstoque

On Error GoTo Erro_MovEstoque_Chama_Tela

    Load X.objMovEstoque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24843
            
    If IsMissing(vP1) Then
    
        Set objMovEstoque = Nothing
        
    Else
    
        Set objMovEstoque = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objMovEstoque.Trata_Parametros(objMovEstoque)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24844
            
        If lErro <> SUCESSO Then Error 24845
        
    End If
    
    X.objMovEstoque.Show
            
    MovEstoque_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MovEstoque_Chama_Tela:

    MovEstoque_Chama_Tela = Err

    Select Case Err
    
        Case 24843, 24844
            Unload X.objMovEstoque
            
        Case 24845
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146850)
        
    End Select
        
    Exit Function

End Function

Function ReqConsumo_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objReqConsumo As ClassMovEstoque

On Error GoTo Erro_ReqConsumo_Chama_Tela

    Load X.objReqConsumo
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24708
            
    If IsMissing(vP1) Then
    
        Set objReqConsumo = Nothing
        
    Else
    
        Set objReqConsumo = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objReqConsumo.Trata_Parametros(objReqConsumo)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24709
            
        If lErro <> SUCESSO Then Error 24710
        
    End If
    
    X.objReqConsumo.Show
            
    ReqConsumo_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ReqConsumo_Chama_Tela:

    ReqConsumo_Chama_Tela = Err

    Select Case Err
    
        Case 24708, 24709
            Unload X.objReqConsumo
            
        Case 24710
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146851)
        
    End Select
        
    Exit Function

End Function

Function Estoque_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objEstoque As ClassProduto

On Error GoTo Erro_Estoque_Chama_Tela

    Load X.objEstoque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24695
            
    If IsMissing(vP1) Then
    
        Set objEstoque = Nothing
        
    Else
    
        Set objEstoque = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objEstoque.Trata_Parametros(objEstoque)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24696
            
        If lErro <> SUCESSO Then Error 24697
        
    End If
    
    X.objEstoque.Show
            
    Estoque_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Estoque_Chama_Tela:

    Estoque_Chama_Tela = Err

    Select Case Err
    
        Case 24695, 24696
            Unload X.objEstoque
            
        Case 24697
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146852)
        
    End Select
        
    Exit Function

End Function

Function Kit_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objKit As ClassKit

On Error GoTo Erro_Kit_Chama_Tela

    Load X.objKit
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24689
            
    If IsMissing(vP1) Then
    
        Set objKit = Nothing
        
    Else
    
        Set objKit = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objKit.Trata_Parametros(objKit)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24690
            
        If lErro <> SUCESSO Then Error 24691
        
    End If
    
    X.objKit.Show
            
    Kit_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Kit_Chama_Tela:

    Kit_Chama_Tela = Err

    Select Case Err
    
        Case 24689, 24690
            Unload X.objKit
            
        Case 24691
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146853)
        
    End Select
        
    Exit Function

End Function

Function Custos_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCustos As ClassProduto

On Error GoTo Erro_Custos_Chama_Tela

    Load X.objCustos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24680
            
    If IsMissing(vP1) Then
    
        Set objCustos = Nothing
        
    Else
    
        Set objCustos = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCustos.Trata_Parametros(objCustos)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24681
            
        If lErro <> SUCESSO Then Error 24682
        
    End If
    
    X.objCustos.Show
            
    Custos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Custos_Chama_Tela:

    Custos_Chama_Tela = Err

    Select Case Err
    
        Case 24680, 24681
            Unload X.objCustos
            
        Case 24682
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146854)
        
    End Select
        
    Exit Function

End Function

Function OrdemProducao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objOrdemProducao As ClassOrdemDeProducao

On Error GoTo Erro_OrdemProducao_Chama_Tela

    Load X.objOrdemProducao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24668
            
    If IsMissing(vP1) Then
    
        Set objOrdemProducao = Nothing
        
    Else
    
        Set objOrdemProducao = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objOrdemProducao.Trata_Parametros(objOrdemProducao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24669
            
        If lErro <> SUCESSO Then Error 24670
        
    End If
    
    X.objOrdemProducao.Show
            
    OrdemProducao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OrdemProducao_Chama_Tela:

    OrdemProducao_Chama_Tela = Err

    Select Case Err
    
        Case 24668, 24669
            Unload X.objOrdemProducao
            
        Case 24670
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146855)
        
    End Select
        
    Exit Function

End Function

Function RecebMaterialF_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRecebMaterialF As ClassNFiscal

On Error GoTo Erro_RecebMaterialF_Chama_Tela

    Load X.objRecebMaterialF
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24913
            
    If IsMissing(vP1) Then
    
        Set objRecebMaterialF = Nothing
        
    Else
    
        Set objRecebMaterialF = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRecebMaterialF.Trata_Parametros(objRecebMaterialF)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24914
            
        If lErro <> SUCESSO Then Error 24915
        
    End If
    
    X.objRecebMaterialF.Show
            
    RecebMaterialF_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RecebMaterialF_Chama_Tela:

    RecebMaterialF_Chama_Tela = Err

    Select Case Err
    
        Case 24913, 24914
            Unload X.objRecebMaterialF
            
        Case 24915
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146856)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntrada_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntrada As ClassNFiscal

On Error GoTo Erro_NFiscalEntrada_Chama_Tela

    Load X.objNFiscalEntrada
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24992
            
    If IsMissing(vP1) Then
    
        Set objNFiscalEntrada = Nothing
        
    Else
    
        Set objNFiscalEntrada = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalEntrada.Trata_Parametros(objNFiscalEntrada)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24993
            
        If lErro <> SUCESSO Then Error 24994
        
    End If
    
    X.objNFiscalEntrada.Show
            
    NFiscalEntrada_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntrada_Chama_Tela:

    NFiscalEntrada_Chama_Tela = Err

    Select Case Err
    
        Case 24992, 24993
            Unload X.objNFiscalEntrada
            
        Case 24994
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146857)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFatEntrada_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFatEntrada As ClassNFiscal

On Error GoTo Erro_NFiscalFatEntrada_Chama_Tela

    Load X.objNFiscalFatEntrada
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19826
            
    If IsMissing(vP1) Then
    
        Set objNFiscalFatEntrada = Nothing
        
    Else
    
        Set objNFiscalFatEntrada = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalFatEntrada.Trata_Parametros(objNFiscalFatEntrada)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19827
            
        If lErro <> SUCESSO Then Error 19828
        
    End If
    
    X.objNFiscalFatEntrada.Show
            
    NFiscalFatEntrada_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFatEntrada_Chama_Tela:

    NFiscalFatEntrada_Chama_Tela = Err

    Select Case Err
    
        Case 19826, 19827
            Unload X.objNFiscalFatEntrada
            
        Case 19828
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146858)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntDev_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntDev As ClassNFiscal

On Error GoTo Erro_NFiscalEntDev_Chama_Tela

    Load X.objNFiscalEntDev
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29631
            
    If IsMissing(vP1) Then
    
        Set objNFiscalEntDev = Nothing
        
    Else
    
        Set objNFiscalEntDev = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalEntDev.Trata_Parametros(objNFiscalEntDev)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29720
            
        If lErro <> SUCESSO Then Error 29721
        
    End If
    
    X.objNFiscalEntDev.Show
            
    NFiscalEntDev_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntDev_Chama_Tela:

    NFiscalEntDev_Chama_Tela = Err

    Select Case Err
    
        Case 29631, 29720
            Unload X.objNFiscalEntDev
            
        Case 29721
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146859)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntRem_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntRem As ClassNFiscal '???? Este objeto está com o mesmo nome que o objeto que referencia a tela (Mario - 30/10/98)

On Error GoTo Erro_NFiscalEntRem_Chama_Tela

    Load X.objNFiscalEntRem
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29878
            
    If IsMissing(vP1) Then
    
        Set objNFiscalEntRem = Nothing
        
    Else
    
        Set objNFiscalEntRem = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objNFiscalEntRem.Trata_Parametros(objNFiscalEntRem)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29879
            
        If lErro <> SUCESSO Then Error 29880
        
    End If
    
    X.objNFiscalEntRem.Show
            
    NFiscalEntRem_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntRem_Chama_Tela:

    NFiscalEntRem_Chama_Tela = Err

    Select Case Err
    
        Case 29878, 29879
            Unload X.objNFiscalEntRem
            
        Case 29880
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146860)
        
    End Select
        
    Exit Function

End Function

Function Empenho_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objEmpenho As ClassEmpenho

On Error GoTo Erro_Empenho_Chama_Tela

    Load X.objEmpenho
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24673
            
    If IsMissing(vP1) Then
    
        Set objEmpenho = Nothing
        
    Else
    
        Set objEmpenho = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objEmpenho.Trata_Parametros(objEmpenho)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24674
            
        If lErro <> SUCESSO Then Error 24675
        
    End If
    
    X.objEmpenho.Show
            
    Empenho_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Empenho_Chama_Tela:

    Empenho_Chama_Tela = Err

    Select Case Err
    
        Case 24673, 24674
            Unload X.objEmpenho
            
        Case 24675
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146861)
        
    End Select
        
    Exit Function

End Function

Function TipoMovEstoque_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoMovEstoque As ClassTipoMovEst

On Error GoTo Erro_TipoMovEstoque_Chama_Tela

    Load X.objTipoMovEstoque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24651
            
    If IsMissing(vP1) Then
    
        Set objTipoMovEstoque = Nothing
        
    Else
    
        Set objTipoMovEstoque = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoMovEstoque.Trata_Parametros(objTipoMovEstoque)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24652
            
        If lErro <> SUCESSO Then Error 24653
        
    End If
    
    X.objTipoMovEstoque.Show
            
    TipoMovEstoque_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoMovEstoque_Chama_Tela:

    TipoMovEstoque_Chama_Tela = Err

    Select Case Err
    
        Case 24651, 24652
            Unload X.objTipoMovEstoque
            
        Case 24653
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146862)
        
    End Select
        
    Exit Function

End Function

Function ClasseUM_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objClasseUM As ClassClasseUM

On Error GoTo Erro_ClasseUM_Chama_Tela

    Load X.objClasseUM
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24309
            
    If IsMissing(vP1) Then
    
        Set objClasseUM = Nothing
        
    Else
    
        Set objClasseUM = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objClasseUM.Trata_Parametros(objClasseUM)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24310
            
        If lErro <> SUCESSO Then Error 24311
        
    End If
    
    X.objClasseUM.Show
            
    ClasseUM_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ClasseUM_Chama_Tela:

    ClasseUM_Chama_Tela = Err

    Select Case Err
    
        Case 24309, 24310
            Unload X.objClasseUM
            
        Case 24311
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146863)
        
    End Select
        
    Exit Function

End Function

Function CategoriaProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCategoriaProduto As ClassCategoriaProduto

On Error GoTo Erro_CategoriaProduto_Chama_Tela

    Load X.objCategoriaProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24312
            
    If IsMissing(vP1) Then
    
        Set objCategoriaProduto = Nothing
        
    Else
    
        Set objCategoriaProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCategoriaProduto.Trata_Parametros(objCategoriaProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24313
            
        If lErro <> SUCESSO Then Error 24314
        
    End If
    
    X.objCategoriaProduto.Show
            
    CategoriaProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CategoriaProduto_Chama_Tela:

    CategoriaProduto_Chama_Tela = Err

    Select Case Err
    
        Case 24312, 24313
            Unload X.objCategoriaProduto
            
        Case 24314
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146864)
        
    End Select
        
    Exit Function

End Function

Function EstoqueProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto

On Error GoTo Erro_EstoqueProduto_Chama_Tela

    Load X.objEstoqueProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24317
            
    If IsMissing(vP1) Then
    
        Set objEstoqueProduto = Nothing
        
    Else
    
        Set objEstoqueProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objEstoqueProduto.Trata_Parametros(objEstoqueProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24318
            
        If lErro <> SUCESSO Then Error 24319
        
    End If
    
    X.objEstoqueProduto.Show
            
    EstoqueProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueProduto_Chama_Tela:

    EstoqueProduto_Chama_Tela = Err

    Select Case Err
    
        Case 24317, 24318
            Unload X.objEstoqueProduto
            
        Case 24319
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146865)
        
    End Select
        
    Exit Function

End Function

Function Produto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Chama_Tela

    Load X.objProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24467
            
    If IsMissing(vP1) Then
    
        Set objProduto = Nothing
        
    Else
    
        Set objProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objProduto.Trata_Parametros(objProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24468
            
        If lErro <> SUCESSO Then Error 24469
        
    End If
    
    X.objProduto.Show
            
    Produto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Produto_Chama_Tela:

    Produto_Chama_Tela = Err

    Select Case Err
    
        Case 24467, 24468
            Unload X.objProduto
            
        Case 24469
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146866)
        
    End Select
        
    Exit Function

End Function

Function ProducaoEntrada_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objProducaoEntrada As ClassMovEstoque

On Error GoTo Erro_ProducaoEntrada_Chama_Tela

    Load X.objProducaoEntrada
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16851
            
    If IsMissing(vP1) Then
    
        Set objProducaoEntrada = Nothing
        
    Else
    
        Set objProducaoEntrada = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objProducaoEntrada.Trata_Parametros(objProducaoEntrada)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16852
            
        If lErro <> SUCESSO Then Error 16853
        
    End If
    
    X.objProducaoEntrada.Show
            
    ProducaoEntrada_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProducaoEntrada_Chama_Tela:

    ProducaoEntrada_Chama_Tela = Err

    Select Case Err
    
        Case 16851, 16852
            Unload X.objProducaoEntrada
            
        Case 16853
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146867)
        
    End Select
        
    Exit Function

End Function

Function ClassificacaoABC_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objClassificacaoABC As ClassClassificacaoABC

On Error GoTo Erro_ClassificacaoABC_Chama_Tela

    Load X.objClassificacaoABC
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45308
            
    If IsMissing(vP1) Then
    
        Set objClassificacaoABC = Nothing
        
    Else
    
        Set objClassificacaoABC = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objClassificacaoABC.Trata_Parametros(objClassificacaoABC)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45309
            
        If lErro <> SUCESSO Then Error 45310
        
    End If
    
    X.objClassificacaoABC.Show
            
    ClassificacaoABC_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ClassificacaoABC_Chama_Tela:

    ClassificacaoABC_Chama_Tela = Err

    Select Case Err
    
        Case 45308, 45309
            Unload X.objClassificacaoABC
            
        Case 45310
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146868)
        
    End Select
        
    Exit Function

End Function

Function ProducaoSaida_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objProducaoSaida As ClassMovEstoque

On Error GoTo Erro_ProducaoSaida_Chama_Tela

    Load X.objProducaoSaida
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16904
            
    If IsMissing(vP1) Then
    
        Set objProducaoSaida = Nothing
        
    Else
    
        Set objProducaoSaida = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objProducaoSaida.Trata_Parametros(objProducaoSaida)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16905
            
        If lErro <> SUCESSO Then Error 16906
        
    End If
    
    X.objProducaoSaida.Show
            
    ProducaoSaida_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProducaoSaida_Chama_Tela:

    ProducaoSaida_Chama_Tela = Err

    Select Case Err
    
        Case 16904, 16905
            Unload X.objProducaoSaida
            
        Case 16906
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146869)
        
    End Select
        
    Exit Function

End Function

Function Inventario_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objInventario As ClassInventario

On Error GoTo Erro_Inventario_Chama_Tela

    Load X.objInventario
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19588
    
    If IsMissing(vP1) Then
    
        Set objInventario = Nothing
        
    Else
    
        Set objInventario = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objInventario.Trata_Parametros(objInventario)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19589
            
        If lErro <> SUCESSO Then Error 19590
        
    End If
    
    X.objInventario.Show
            
    Inventario_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Inventario_Chama_Tela:

    Inventario_Chama_Tela = Err

    Select Case Err
    
        Case 19588, 19589
            Unload X.objInventario
            
        Case 19590
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146870)
        
    End Select
        
    Exit Function

End Function

Function InventarioLote_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objInventarioLote As ClassInventario

On Error GoTo Erro_InventarioLote_Chama_Tela

    Load X.objInventarioLote
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19591
            
    If IsMissing(vP1) Then
    
        Set objInventarioLote = Nothing
        
    Else
    
        Set objInventarioLote = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objInventarioLote.Trata_Parametros(objInventarioLote)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19592
            
        If lErro <> SUCESSO Then Error 19593
        
    End If
    
    X.objInventarioLote.Show
            
    InventarioLote_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_InventarioLote_Chama_Tela:

    InventarioLote_Chama_Tela = Err

    Select Case Err
    
        Case 19591, 19592
            Unload X.objInventarioLote
            
        Case 19593
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146871)
        
    End Select
        
    Exit Function

End Function

Function TipoProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto

On Error GoTo Erro_TipoProduto_Chama_Tela

    Load X.objTipoProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24470
            
    If IsMissing(vP1) Then
    
        Set objTipoProduto = Nothing
        
    Else
    
        Set objTipoProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoProduto.Trata_Parametros(objTipoProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24471
            
        If lErro <> SUCESSO Then Error 24472
        
    End If
    
    X.objTipoProduto.Show
            
    TipoProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoProduto_Chama_Tela:

    TipoProduto_Chama_Tela = Err

    Select Case Err
    
        Case 24470, 24471
            Unload X.objTipoProduto
            
        Case 24472
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146872)
        
    End Select
        
    Exit Function

End Function

Function Reserva_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objReserva As ClassReserva

On Error GoTo Erro_Reserva_Chama_Tela

    Load X.objReserva
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24558
            
    If IsMissing(vP1) Then
    
        Set objReserva = Nothing
        
    Else
    
        Set objReserva = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objReserva.Trata_Parametros(objReserva)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24559
            
        If lErro <> SUCESSO Then Error 24560
        
    End If
    
    X.objReserva.Show
            
    Reserva_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Reserva_Chama_Tela:

    Reserva_Chama_Tela = Err

    Select Case Err
    
        Case 24558, 24559
            Unload X.objReserva
            
        Case 24560
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146873)
        
    End Select
        
    Exit Function

End Function

Function Transfer_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTransfer As ClassMovEstoque

On Error GoTo Erro_Transfer_Chama_Tela

    Load X.objTransfer
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29671
            
    If IsMissing(vP1) Then
    
        Set objTransfer = Nothing
        
    Else
    
        Set objTransfer = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTransfer.Trata_Parametros(objTransfer)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29672
            
        If lErro <> SUCESSO Then Error 29673
        
    End If
    
    X.objTransfer.Show
            
    Transfer_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Transfer_Chama_Tela:

    Transfer_Chama_Tela = Err

    Select Case Err
    
        Case 29671, 29672
            Unload X.objTransfer
            
        Case 29673
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146874)
        
    End Select
        
    Exit Function

End Function

Function RecebMaterialC_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRecebMaterialC As ClassNFiscal

On Error GoTo Erro_RecebMaterialC_Chama_Tela

    Load X.objRecebMaterialC
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29674
            
    If IsMissing(vP1) Then
    
        Set objRecebMaterialC = Nothing
        
    Else
    
        Set objRecebMaterialC = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRecebMaterialC.Trata_Parametros(objRecebMaterialC)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29675
            
        If lErro <> SUCESSO Then Error 29676
        
    End If
    
    X.objRecebMaterialC.Show
            
    RecebMaterialC_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RecebMaterialC_Chama_Tela:

    RecebMaterialC_Chama_Tela = Err

    Select Case Err
    
        Case 29674, 29675
            Unload X.objRecebMaterialC
            
        Case 29676
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146875)
        
    End Select
        
    Exit Function

End Function

Function FornecedorProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFornecedorProduto As ClassFornecedorProduto

On Error GoTo Erro_FornecedorProduto_Chama_Tela

    Load X.objFornecedorProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24364
            
    If IsMissing(vP1) Then
    
        Set objFornecedorProduto = Nothing
        
    Else
    
        Set objFornecedorProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFornecedorProduto.Trata_Parametros(objFornecedorProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24365
            
        If lErro <> SUCESSO Then Error 24366
        
    End If
    
    X.objFornecedorProduto.Show
            
    FornecedorProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FornecedorProduto_Chama_Tela:

    FornecedorProduto_Chama_Tela = Err

    Select Case Err
    
        Case 24364, 24365
            Unload X.objFornecedorProduto
            
        Case 24366
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146876)
        
    End Select
        
    Exit Function

End Function

Function EstoqueInicial_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_EstoqueInicial_Chama_Tela

    Load X.objEstoqueInicial
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24320
    
    X.objEstoqueInicial.Show
            
    EstoqueInicial_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueInicial_Chama_Tela:

    EstoqueInicial_Chama_Tela = Err

    Select Case Err
    
        Case 24320
            Unload X.objEstoqueInicial
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146877)
        
    End Select
        
    Exit Function

End Function

Function GeracaoOP_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoOP_Chama_Tela

    Load X.objGeracaoOP
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24671
    
    X.objGeracaoOP.Show
            
    GeracaoOP_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoOP_Chama_Tela:

    GeracaoOP_Chama_Tela = Err

    Select Case Err
    
        Case 24671
            Unload X.objGeracaoOP
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146878)
        
    End Select
        
    Exit Function

End Function

Function LoteEst_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objInvLote As ClassInvLote

On Error GoTo Erro_LoteEst_Chama_Tela

    Load X.objLoteEst
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29638
        
    If IsMissing(vP1) Then
    
        Set objInvLote = Nothing
        
    Else
    
        Set objInvLote = vP1
        
    End If

    lErro = X.objLoteEst.Trata_Parametros(objInvLote)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24289
            
    If lErro <> SUCESSO Then Error 45538

    X.objLoteEst.Show
            
    LoteEst_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteEst_Chama_Tela:

    LoteEst_Chama_Tela = Err

    Select Case Err
    
        Case 29638
            Unload X.objLoteEst
            
        Case 45538
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146879)
        
    End Select
        
    Exit Function

End Function

Function LoteEstAtualiza_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_LoteEstAtualiza_Chama_Tela

    Load X.objLoteEstAtualiza
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29639
    
    X.objLoteEstAtualiza.Show
            
    LoteEstAtualiza_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteEstAtualiza_Chama_Tela:

    LoteEstAtualiza_Chama_Tela = Err

    Select Case Err
    
        Case 29639
            Unload X.objLoteEstAtualiza
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146880)
        
    End Select
        
    Exit Function

End Function

Function EstoqueAlmoxarifadoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueAlmoxarifado1 As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueAlmoxarifadoLista_Chama_Tela

    Load X.objEstoqueAlmoxarifadoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24288
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    Set objEstoqueAlmoxarifado1 = Nothing
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objEstoqueAlmoxarifadoLista.Trata_Parametros(colSelecao, objEstoqueAlmoxarifado1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24289
            
    If lErro <> SUCESSO Then Error 24290
        
    X.objEstoqueAlmoxarifadoLista.Show
            
    EstoqueAlmoxarifadoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueAlmoxarifadoLista_Chama_Tela:

    EstoqueAlmoxarifadoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24288, 24289
            Unload X.objEstoqueAlmoxarifadoLista
            
        Case 24290
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146881)
        
    End Select
        
    Exit Function

End Function

Function ClasseUMLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objClasseUM1 As ClassClasseUM
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ClasseUMLista_Chama_Tela

    Load X.objClasseUMLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24454
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objClasseUM1 = Nothing
        
    Else
    
        Set objClasseUM1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objClasseUMLista.Trata_Parametros(colSelecao, objClasseUM1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24455
            
    If lErro <> SUCESSO Then Error 24456
        
    X.objClasseUMLista.Show
            
    ClasseUMLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ClasseUMLista_Chama_Tela:

    ClasseUMLista_Chama_Tela = Err

    Select Case Err
    
        Case 24454, 24455
            Unload X.objClasseUMLista
            
        Case 24456
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146882)
        
    End Select
        
    Exit Function

End Function

Function UnidadeDeMedidaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objUnidadeDeMedida1 As ClassUnidadeDeMedida
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_UnidadeDeMedidaLista_Chama_Tela

    Load X.objUnidadeDeMedidaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24587
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objUnidadeDeMedida1 = Nothing
        
    Else
    
        Set objUnidadeDeMedida1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objUnidadeDeMedidaLista.Trata_Parametros(colSelecao, objUnidadeDeMedida1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24588
            
    If lErro <> SUCESSO Then Error 24589
        
    X.objUnidadeDeMedidaLista.Show
            
    UnidadeDeMedidaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_UnidadeDeMedidaLista_Chama_Tela:

    UnidadeDeMedidaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24587, 24588
            Unload X.objUnidadeDeMedidaLista
            
        Case 24589
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146883)
        
    End Select
        
    Exit Function

End Function

Function TipoProdutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTipoProduto1 As ClassTipoDeProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TipoProdutoLista_Chama_Tela

    Load X.objTipoProdutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24590
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTipoProduto1 = Nothing
        
    Else
    
        Set objTipoProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTipoProdutoLista.Trata_Parametros(colSelecao, objTipoProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24591
            
    If lErro <> SUCESSO Then Error 24592
        
    X.objTipoProdutoLista.Show
            
    TipoProdutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoProdutoLista_Chama_Tela:

    TipoProdutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24590, 24591
            Unload X.objTipoProdutoLista
            
        Case 24592
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146884)
        
    End Select
        
    Exit Function

End Function

Function ProdutoAtivoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProdutoAtivo1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoAtivoLista_Chama_Tela

    Load X.objProdutoAtivoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29752
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProdutoAtivo1 = Nothing
        
    Else
    
        Set objProdutoAtivo1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoAtivoLista.Trata_Parametros(colSelecao, objProdutoAtivo1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29753
            
    If lErro <> SUCESSO Then Error 29754
        
    X.objProdutoAtivoLista.Show
            
    ProdutoAtivoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoAtivoLista_Chama_Tela:

    ProdutoAtivoLista_Chama_Tela = Err

    Select Case Err
    
        Case 29752, 29753
            Unload X.objProdutoAtivoLista
            
        Case 29754
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146885)
        
    End Select
        
    Exit Function

End Function

Function ProdutoEstoqueLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProdutoEstoque1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoEstoqueLista_Chama_Tela

    Load X.objProdutoEstoqueLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29755
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProdutoEstoque1 = Nothing
        
    Else
    
        Set objProdutoEstoque1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoEstoqueLista.Trata_Parametros(colSelecao, objProdutoEstoque1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29756
            
    If lErro <> SUCESSO Then Error 29757
        
    X.objProdutoEstoqueLista.Show
            
    ProdutoEstoqueLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoEstoqueLista_Chama_Tela:

    ProdutoEstoqueLista_Chama_Tela = Err

    Select Case Err
    
        Case 29755, 29756
            Unload X.objProdutoEstoqueLista
            
        Case 29757
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146886)
        
    End Select
        
    Exit Function

End Function

Function RecebMaterialFLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRecebMaterialF1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RecebMaterialFLista_Chama_Tela

    Load X.objRecebMaterialFLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24867
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objRecebMaterialF1 = Nothing
        
    Else
    
        Set objRecebMaterialF1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objRecebMaterialFLista.Trata_Parametros(colSelecao, objRecebMaterialF1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24868
            
    If lErro <> SUCESSO Then Error 24869
        
    X.objRecebMaterialFLista.Show
            
    RecebMaterialFLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RecebMaterialFLista_Chama_Tela:

    RecebMaterialFLista_Chama_Tela = Err

    Select Case Err
    
        Case 24867, 24868
            Unload X.objRecebMaterialFLista
            
        Case 24869
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146887)
        
    End Select
        
    Exit Function

End Function

Function RecebMaterialLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRecebMaterial1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RecebMaterialLista_Chama_Tela

    Load X.objRecebMaterialLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24832
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objRecebMaterial1 = Nothing
        
    Else
    
        Set objRecebMaterial1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objRecebMaterialLista.Trata_Parametros(colSelecao, objRecebMaterial1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24833
            
    If lErro <> SUCESSO Then Error 24835
        
    X.objRecebMaterialLista.Show
            
    RecebMaterialLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RecebMaterialLista_Chama_Tela:

    RecebMaterialLista_Chama_Tela = Err

    Select Case Err
    
        Case 24832, 24833
            Unload X.objRecebMaterialLista
            
        Case 24835
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146888)
        
    End Select
        
    Exit Function

End Function

Function RecebMaterialCLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRecebMaterialC1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RecebMaterialCLista_Chama_Tela

    Load X.objRecebMaterialCLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24961
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objRecebMaterialC1 = Nothing
        
    Else
    
        Set objRecebMaterialC1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objRecebMaterialCLista.Trata_Parametros(colSelecao, objRecebMaterialC1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24962
            
    If lErro <> SUCESSO Then Error 24963
        
    X.objRecebMaterialCLista.Show
            
    RecebMaterialCLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RecebMaterialCLista_Chama_Tela:

    RecebMaterialCLista_Chama_Tela = Err

    Select Case Err
    
        Case 24961, 24962
            Unload X.objRecebMaterialCLista
            
        Case 24963
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146889)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntradaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntrada1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalEntradaLista_Chama_Tela

    Load X.objNFiscalEntradaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29523
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalEntrada1 = Nothing
        
    Else
    
        Set objNFiscalEntrada1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalEntradaLista.Trata_Parametros(colSelecao, objNFiscalEntrada1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29524
            
    If lErro <> SUCESSO Then Error 29525
        
    X.objNFiscalEntradaLista.Show
            
    NFiscalEntradaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntradaLista_Chama_Tela:

    NFiscalEntradaLista_Chama_Tela = Err

    Select Case Err
    
        Case 29523, 29524
            Unload X.objNFiscalEntradaLista
            
        Case 29525
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146890)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntDevLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntDev1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalEntDevLista_Chama_Tela

    Load X.objNFiscalEntDevLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24839
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalEntDev1 = Nothing
        
    Else
    
        Set objNFiscalEntDev1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalEntDevLista.Trata_Parametros(colSelecao, objNFiscalEntDev1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24840
            
    If lErro <> SUCESSO Then Error 24841
        
    X.objNFiscalEntDevLista.Show
            
    NFiscalEntDevLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntDevLista_Chama_Tela:

    NFiscalEntDevLista_Chama_Tela = Err

    Select Case Err
    
        Case 24839, 24840
            Unload X.objNFiscalEntDevLista
            
        Case 24841
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146891)
        
    End Select
        
    Exit Function

End Function

Function NFEntNaoDevLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFEntNaoDev1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFEntNaoDevLista_Chama_Tela

    Load X.objNFEntNaoDevLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19512
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFEntNaoDev1 = Nothing
        
    Else
    
        Set objNFEntNaoDev1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFEntNaoDevLista.Trata_Parametros(colSelecao, objNFEntNaoDev1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19513
            
    If lErro <> SUCESSO Then Error 19514
        
    X.objNFEntNaoDevLista.Show
            
    NFEntNaoDevLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFEntNaoDevLista_Chama_Tela:

    NFEntNaoDevLista_Chama_Tela = Err

    Select Case Err
    
        Case 19512, 19513
            Unload X.objNFEntNaoDevLista
            
        Case 19514
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146892)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntRemLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntRem1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalEntRemLista_Chama_Tela

    Load X.objNFiscalEntRemLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16823
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalEntRem1 = Nothing
        
    Else
    
        Set objNFiscalEntRem1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalEntRemLista.Trata_Parametros(colSelecao, objNFiscalEntRem1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 16824
            
    If lErro <> SUCESSO Then Error 16825
        
    X.objNFiscalEntRemLista.Show
            
    NFiscalEntRemLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntRemLista_Chama_Tela:

    NFiscalEntRemLista_Chama_Tela = Err

    Select Case Err
    
        Case 16823, 16824
            Unload X.objNFiscalEntRemLista
            
        Case 16825
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146893)
        
    End Select
        
    Exit Function

End Function

Function NFiscalEntradaNormalLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalEntradaNormal1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalEntradaNormalLista_Chama_Tela

    Load X.objNFiscalEntradaNormalLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29535
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalEntradaNormal1 = Nothing
        
    Else
    
        Set objNFiscalEntradaNormal1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalEntradaNormalLista.Trata_Parametros(colSelecao, objNFiscalEntradaNormal1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29536
            
    If lErro <> SUCESSO Then Error 29537
        
    X.objNFiscalEntradaNormalLista.Show
            
    NFiscalEntradaNormalLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalEntradaNormalLista_Chama_Tela:

    NFiscalEntradaNormalLista_Chama_Tela = Err

    Select Case Err
    
        Case 29535, 29536
            Unload X.objNFiscalEntradaNormalLista
            
        Case 29537
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146894)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFatEntradaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFatEntrada As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalFatEntradaLista_Chama_Tela

    Load X.objNFiscalFatEntradaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19819
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalFatEntrada = Nothing
        
    Else
    
        Set objNFiscalFatEntrada = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalFatEntradaLista.Trata_Parametros(colSelecao, objNFiscalFatEntrada, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19820
            
    If lErro <> SUCESSO Then Error 19821
        
    X.objNFiscalFatEntradaLista.Show
            
    NFiscalFatEntradaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFatEntradaLista_Chama_Tela:

    NFiscalFatEntradaLista_Chama_Tela = Err

    Select Case Err
    
        Case 19819, 19820
            Unload X.objNFiscalFatEntradaLista
            
        Case 19821
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146895)
        
    End Select
        
    Exit Function

End Function

Function NFiscalFatEntradaNormalLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalFatEntradaNormal As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalFatEntradaNormalLista_Chama_Tela

    Load X.objNFiscalFatEntradaNormalLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19822
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalFatEntradaNormal = Nothing
        
    Else
    
        Set objNFiscalFatEntradaNormal = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalFatEntradaNormalLista.Trata_Parametros(colSelecao, objNFiscalFatEntradaNormal, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19823
            
    If lErro <> SUCESSO Then Error 19824
        
    X.objNFiscalFatEntradaNormalLista.Show
            
    NFiscalFatEntradaNormalLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalFatEntradaNormalLista_Chama_Tela:

    NFiscalFatEntradaNormalLista_Chama_Tela = Err

    Select Case Err
    
        Case 19822, 19823
            Unload X.objNFiscalFatEntradaNormalLista
            
        Case 19824
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146896)
        
    End Select
        
    Exit Function

End Function

Function NFiscalInternaSaidaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalInternaSaida1 As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFiscalInternaSaidaLista_Chama_Tela

    Load X.objNFiscalInternaSaidaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24985
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalInternaSaida1 = Nothing
        
    Else
    
        Set objNFiscalInternaSaida1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFiscalInternaSaidaLista.Trata_Parametros(colSelecao, objNFiscalInternaSaida1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24986
            
    If lErro <> SUCESSO Then Error 24987
        
    X.objNFiscalInternaSaidaLista.Show
            
    NFiscalInternaSaidaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFiscalInternaSaidaLista_Chama_Tela:

    NFiscalInternaSaidaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24985, 24986
            Unload X.objNFiscalInternaSaidaLista
            
        Case 24987
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146897)
        
    End Select
        
    Exit Function

End Function

Function NFInternaSaidaDestinatarioLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objNFiscalInternaSaida As ClassNFiscal
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_NFInternaSaidaDestinatarioLista_Chama_Tela

    Load X.objNFInternaSaidaDestinatarioLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45958
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objNFiscalInternaSaida = Nothing
        
    Else
    
        Set objNFiscalInternaSaida = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objNFInternaSaidaDestinatarioLista.Trata_Parametros(colSelecao, objNFiscalInternaSaida, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45959
            
    If lErro <> SUCESSO Then Error 45960
        
    X.objNFInternaSaidaDestinatarioLista.Show
            
    NFInternaSaidaDestinatarioLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_NFInternaSaidaDestinatarioLista_Chama_Tela:

    NFInternaSaidaDestinatarioLista_Chama_Tela = Err

    Select Case Err
    
        Case 45958, 45959
            Unload X.objNFInternaSaidaDestinatarioLista
            
        Case 45960
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146898)
        
    End Select
        
    Exit Function

End Function

Function OrdemProducaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOrdemProducao1 As ClassOrdemDeProducao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OrdemProducaoLista_Chama_Tela

    Load X.objOrdemProducaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24590
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objOrdemProducao1 = Nothing
        
    Else
    
        Set objOrdemProducao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objOrdemProducaoLista.Trata_Parametros(colSelecao, objOrdemProducao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24591
            
    If lErro <> SUCESSO Then Error 24592
        
    X.objOrdemProducaoLista.Show
            
    OrdemProducaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OrdemProducaoLista_Chama_Tela:

    OrdemProducaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24590, 24591
            Unload X.objOrdemProducaoLista
            
        Case 24592
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146899)
        
    End Select
        
    Exit Function

End Function

Function OrdemProdListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOrdemProducao1 As ClassOrdemDeProducao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OrdemProdListaModal_Chama_Tela

    Load X.objOrdemProdListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29715
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objOrdemProducao1 = Nothing
        
    Else
    
        Set objOrdemProducao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objOrdemProdListaModal.Trata_Parametros(colSelecao, objOrdemProducao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29716
            
    If lErro <> SUCESSO Then Error 29717
        
    X.objOrdemProdListaModal.Show vbModal
            
    OrdemProdListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OrdemProdListaModal_Chama_Tela:

    OrdemProdListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 29715, 29716
            Unload X.objOrdemProdListaModal
            
        Case 29717
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146900)
        
    End Select
        
    Exit Function

End Function

Function OrdProdTodasListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOrdemProducao1 As ClassOrdemDeProducao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OrdProdTodasListaModal_Chama_Tela

    Load X.objOrdProdTodasListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29775
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objOrdemProducao1 = Nothing
        
    Else
    
        Set objOrdemProducao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objOrdProdTodasListaModal.Trata_Parametros(colSelecao, objOrdemProducao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29776
            
    If lErro <> SUCESSO Then Error 29777
        
    X.objOrdProdTodasListaModal.Show vbModal
            
    OrdProdTodasListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OrdProdTodasListaModal_Chama_Tela:

    OrdProdTodasListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 29775, 29776
            Unload X.objOrdProdTodasListaModal
            
        Case 29777
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146901)
        
    End Select
        
    Exit Function

End Function

Function OrdProdBaixadasListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOrdemProducao1 As ClassOrdemDeProducao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OrdProdBaixadasListaModal_Chama_Tela

    Load X.objOrdProdBaixadasListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29778
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objOrdemProducao1 = Nothing
        
    Else
    
        Set objOrdemProducao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objOrdProdBaixadasListaModal.Trata_Parametros(colSelecao, objOrdemProducao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29779
            
    If lErro <> SUCESSO Then Error 29780
        
    X.objOrdProdBaixadasListaModal.Show vbModal
            
    OrdProdBaixadasListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OrdProdBaixadasListaModal_Chama_Tela:

    OrdProdBaixadasListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 29778, 29779
            Unload X.objOrdProdBaixadasListaModal
            
        Case 29780
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146902)
        
    End Select
        
    Exit Function

End Function

Function MovEstoqueTransferenciaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objMovEstoqueTransferencia1 As ClassMovEstoque
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_MovEstoqueTransferenciaLista_Chama_Tela

    Load X.objMovEstoqueTransferenciaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24909
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objMovEstoqueTransferencia1 = Nothing
        
    Else
    
        Set objMovEstoqueTransferencia1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objMovEstoqueTransferenciaLista.Trata_Parametros(colSelecao, objMovEstoqueTransferencia1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24910
            
    If lErro <> SUCESSO Then Error 24911
        
    X.objMovEstoqueTransferenciaLista.Show
            
    MovEstoqueTransferenciaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MovEstoqueTransferenciaLista_Chama_Tela:

    MovEstoqueTransferenciaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24909, 24910
            Unload X.objMovEstoqueTransferenciaLista
            
        Case 24911
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146903)
        
    End Select
        
    Exit Function

End Function

Function MovEstoqueLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objMovEstoque1 As ClassMovEstoque
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_MovEstoqueLista_Chama_Tela

    Load X.objMovEstoqueLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24797
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objMovEstoque1 = Nothing
        
    Else
    
        Set objMovEstoque1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objMovEstoqueLista.Trata_Parametros(colSelecao, objMovEstoque1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24798
            
    If lErro <> SUCESSO Then Error 24799
        
    X.objMovEstoqueLista.Show
            
    MovEstoqueLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MovEstoqueLista_Chama_Tela:

    MovEstoqueLista_Chama_Tela = Err

    Select Case Err
    
        Case 24797, 24798
            Unload X.objMovEstoqueLista
            
        Case 24799
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146904)
        
    End Select
        
    Exit Function

End Function

Function MovEstoqueOPLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objMovEstoque1 As ClassMovEstoque
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_MovEstoqueOPLista_Chama_Tela

    Load X.objMovEstoqueOPLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 54799
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objMovEstoque1 = Nothing
        
    Else
    
        Set objMovEstoque1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objMovEstoqueOPLista.Trata_Parametros(colSelecao, objMovEstoque1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 54800
            
    If lErro <> SUCESSO Then Error 54801
        
    X.objMovEstoqueOPLista.Show
            
    MovEstoqueOPLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MovEstoqueOPLista_Chama_Tela:

    MovEstoqueOPLista_Chama_Tela = Err

    Select Case Err
    
        Case 54799, 54800
            Unload X.objMovEstoqueOPLista
            
        Case 54801
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146905)
        
    End Select
        
    Exit Function

End Function

Function MovEstoqueLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objMovEstoque1 As ClassMovEstoque
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_MovEstoqueLista_Consulta_Chama_Tela

    Load X.objMovEstoqueLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29838
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objMovEstoque1 = Nothing
        
    Else
    
        Set objMovEstoque1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objMovEstoqueLista_Consulta.Trata_Parametros(colSelecao, objMovEstoque1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29839
            
    If lErro <> SUCESSO Then Error 29840
        
    X.objMovEstoqueLista_Consulta.Show
            
    MovEstoqueLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MovEstoqueLista_Consulta_Chama_Tela:

    MovEstoqueLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 29838, 29839
            Unload X.objMovEstoqueLista_Consulta
            
        Case 29840
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146906)
        
    End Select
        
    Exit Function

End Function

Function MovEstoqueInternoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objMovEstoqueInterno1 As ClassMovEstoque
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_MovEstoqueInternoLista_Chama_Tela

    Load X.objMovEstoqueInternoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24836
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objMovEstoqueInterno1 = Nothing
        
    Else
    
        Set objMovEstoqueInterno1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objMovEstoqueInternoLista.Trata_Parametros(colSelecao, objMovEstoqueInterno1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24837
            
    If lErro <> SUCESSO Then Error 24838
        
    X.objMovEstoqueInternoLista.Show
            
    MovEstoqueInternoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MovEstoqueInternoLista_Chama_Tela:

    MovEstoqueInternoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24836, 24837
            Unload X.objMovEstoqueInternoLista
            
        Case 24838
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146907)
        
    End Select
        
    Exit Function

End Function

Function ProdutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoLista_Chama_Tela

    Load X.objProdutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24492
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto1 = Nothing
        
    Else
    
        Set objProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoLista.Trata_Parametros(colSelecao, objProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24493
            
    If lErro <> SUCESSO Then Error 24494
        
    X.objProdutoLista.Show
            
    ProdutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoLista_Chama_Tela:

    ProdutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24492, 24493
            Unload X.objProdutoLista
            
        Case 24494
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146908)
        
    End Select
        
    Exit Function

End Function

Function ProdutoCompraLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoCompraLista_Chama_Tela

    Load X.objProdutoCompraLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29819
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto1 = Nothing
        
    Else
    
        Set objProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoCompraLista.Trata_Parametros(colSelecao, objProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29820
            
    If lErro <> SUCESSO Then Error 29821
        
    X.objProdutoCompraLista.Show
            
    ProdutoCompraLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoCompraLista_Chama_Tela:

    ProdutoCompraLista_Chama_Tela = Err

    Select Case Err
    
        Case 29819, 29820
            Unload X.objProdutoCompraLista
            
        Case 29821
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146909)
        
    End Select
        
    Exit Function

End Function

Function ProdutoProduzivelLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoProduzivelLista_Chama_Tela

    Load X.objProdutoProduzivelLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45905
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto = Nothing
        
    Else
    
        Set objProduto = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoProduzivelLista.Trata_Parametros(colSelecao, objProduto, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45906
            
    If lErro <> SUCESSO Then Error 45907
        
    X.objProdutoProduzivelLista.Show
            
    ProdutoProduzivelLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoProduzivelLista_Chama_Tela:

    ProdutoProduzivelLista_Chama_Tela = Err

    Select Case Err
    
        Case 45905, 45906
            Unload X.objProdutoProduzivelLista
            
        Case 45907
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146910)
        
    End Select
        
    Exit Function

End Function

Function ProdutoProduz_EstoqLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoProduz_EstoqLista_Chama_Tela

    Load X.objProdutoProduz_EstoqLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45945
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto = Nothing
        
    Else
    
        Set objProduto = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoProduz_EstoqLista.Trata_Parametros(colSelecao, objProduto, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45946
            
    If lErro <> SUCESSO Then Error 45947
        
    X.objProdutoProduz_EstoqLista.Show
            
    ProdutoProduz_EstoqLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoProduz_EstoqLista_Chama_Tela:

    ProdutoProduz_EstoqLista_Chama_Tela = Err

    Select Case Err
    
        Case 45945, 45946
            Unload X.objProdutoProduz_EstoqLista
            
        Case 45947
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146911)
        
    End Select
        
    Exit Function

End Function

Function EstoqueInicialLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EstoqueInicialLista_Chama_Tela

    Load X.objEstoqueInicialLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45932
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEstoqueProduto = Nothing
        
    Else
    
        Set objEstoqueProduto = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEstoqueInicialLista.Trata_Parametros(colSelecao, objEstoqueProduto, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45933
            
    If lErro <> SUCESSO Then Error 45934
        
    X.objEstoqueInicialLista.Show
            
    EstoqueInicialLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EstoqueInicialLista_Chama_Tela:

    EstoqueInicialLista_Chama_Tela = Err

    Select Case Err
    
        Case 45932, 45933
            Unload X.objEstoqueInicialLista
            
        Case 45934
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146912)
        
    End Select
        
    Exit Function

End Function

Function ProdutoEstoquePCPLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoEstoquePCPLista_Chama_Tela

    Load X.objProdutoEstoquePCPLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29822
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto1 = Nothing
        
    Else
    
        Set objProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoEstoquePCPLista.Trata_Parametros(colSelecao, objProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29823
            
    If lErro <> SUCESSO Then Error 29824
        
    X.objProdutoEstoquePCPLista.Show
            
    ProdutoEstoquePCPLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoEstoquePCPLista_Chama_Tela:

    ProdutoEstoquePCPLista_Chama_Tela = Err

    Select Case Err
    
        Case 29822, 29823
            Unload X.objProdutoEstoquePCPLista
            
        Case 29824
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146913)
        
    End Select
        
    Exit Function

End Function

Function ProdutoVendaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoVendaLista_Chama_Tela

    Load X.objProdutoVendaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29825
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto1 = Nothing
        
    Else
    
        Set objProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoVendaLista.Trata_Parametros(colSelecao, objProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29826
            
    If lErro <> SUCESSO Then Error 29827
        
    X.objProdutoVendaLista.Show
            
    ProdutoVendaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoVendaLista_Chama_Tela:

    ProdutoVendaLista_Chama_Tela = Err

    Select Case Err
    
        Case 29825, 29826
            Unload X.objProdutoVendaLista
            
        Case 29827
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146914)
        
    End Select
        
    Exit Function

End Function

Function KitLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objKit1 As ClassKit
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_KitLista_Chama_Tela

    Load X.objKitLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24628
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objKit1 = Nothing
        
    Else
    
        Set objKit1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objKitLista.Trata_Parametros(colSelecao, objKit1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24629
            
    If lErro <> SUCESSO Then Error 24630
        
    X.objKitLista.Show
            
    KitLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_KitLista_Chama_Tela:

    KitLista_Chama_Tela = Err

    Select Case Err
    
        Case 24628, 24629
            Unload X.objKitLista
            
        Case 24630
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146915)
        
    End Select
        
    Exit Function

End Function

Function KitVersaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objKit1 As ClassKit
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_KitVersaoLista_Chama_Tela

    Load X.objKitVersaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45566
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objKit1 = Nothing
        
    Else
    
        Set objKit1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objKitVersaoLista.Trata_Parametros(colSelecao, objKit1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45567
            
    If lErro <> SUCESSO Then Error 45568
        
    X.objKitVersaoLista.Show
            
    KitVersaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_KitVersaoLista_Chama_Tela:

    KitVersaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 45566, 45567
            Unload X.objKitVersaoLista
            
        Case 45568
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146916)
        
    End Select
        
    Exit Function

End Function

Function InventarioLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objInventario1 As ClassInventario
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_InventarioLista_Chama_Tela

    Load X.objInventarioLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29859
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objInventario1 = Nothing
        
    Else
    
        Set objInventario1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objInventarioLista.Trata_Parametros(colSelecao, objInventario1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29860
            
    If lErro <> SUCESSO Then Error 29861
        
    X.objInventarioLista.Show
            
    InventarioLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_InventarioLista_Chama_Tela:

    InventarioLista_Chama_Tela = Err

    Select Case Err
    
        Case 29859, 29860
            Unload X.objInventarioLista
            
        Case 29861
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146917)
        
    End Select
        
    Exit Function

End Function

Function InventarioLoteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objInventarioLote1 As ClassInventario
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_InventarioLoteLista_Chama_Tela

    Load X.objInventarioLoteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29862
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objInventarioLote1 = Nothing
        
    Else
    
        Set objInventarioLote1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objInventarioLoteLista.Trata_Parametros(colSelecao, objInventarioLote1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29863
            
    If lErro <> SUCESSO Then Error 29864
        
    X.objInventarioLoteLista.Show
            
    InventarioLoteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_InventarioLoteLista_Chama_Tela:

    InventarioLoteLista_Chama_Tela = Err

    Select Case Err
    
        Case 29862, 29863
            Unload X.objInventarioLoteLista
            
        Case 29864
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146918)
        
    End Select
        
    Exit Function

End Function

Function InventarioLoteLista_Lote_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objInventarioLote As ClassInventario
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_InventarioLoteLista_Lote_Chama_Tela

    Load X.objInventarioLoteLista_Lote
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45892
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objInventarioLote = Nothing
        
    Else
    
        Set objInventarioLote = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objInventarioLoteLista_Lote.Trata_Parametros(colSelecao, objInventarioLote, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45893
            
    If lErro <> SUCESSO Then Error 45894
        
    X.objInventarioLoteLista_Lote.Show
            
    InventarioLoteLista_Lote_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_InventarioLoteLista_Lote_Chama_Tela:

    InventarioLoteLista_Lote_Chama_Tela = Err

    Select Case Err
    
        Case 45892, 45893
            Unload X.objInventarioLoteLista_Lote
            
        Case 45894
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146919)
        
    End Select
        
    Exit Function

End Function

Function InvLotePendenteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objInvLote1 As ClassInvLote
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_InvLotePendenteLista_Chama_Tela

    Load X.objInvLotePendenteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29873
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objInvLote1 = Nothing
        
    Else
    
        Set objInvLote1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objInvLotePendenteLista.Trata_Parametros(colSelecao, objInvLote1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29874
            
    If lErro <> SUCESSO Then Error 29875
        
    X.objInvLotePendenteLista.Show
            
    InvLotePendenteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_InvLotePendenteLista_Chama_Tela:

    InvLotePendenteLista_Chama_Tela = Err

    Select Case Err
    
        Case 29873, 29874
            Unload X.objInvLotePendenteLista
            
        Case 29875
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146920)
        
    End Select
        
    Exit Function

End Function

Function ItensNFiscalSaidaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItensNFiscalSaida1 As ClassItemNF
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ItensNFiscalSaidaLista_Chama_Tela

    Load X.objItensNFiscalSaidaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29633
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objItensNFiscalSaida1 = Nothing
        
    Else
    
        Set objItensNFiscalSaida1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objItensNFiscalSaidaLista.Trata_Parametros(colSelecao, objItensNFiscalSaida1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29634
            
    If lErro <> SUCESSO Then Error 29635
        
    X.objItensNFiscalSaidaLista.Show
            
    ItensNFiscalSaidaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ItensNFiscalSaidaLista_Chama_Tela:

    ItensNFiscalSaidaLista_Chama_Tela = Err

    Select Case Err
    
        Case 29633, 29634
            Unload X.objItensNFiscalSaidaLista
            
        Case 29635
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146921)
        
    End Select
        
    Exit Function

End Function

Function ItensNFiscalEntNaoDevLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItensNFiscalEntNaoDev1 As ClassItemNF
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ItensNFiscalEntNaoDevLista_Chama_Tela

    Load X.objItensNFiscalEntNaoDevLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19509
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objItensNFiscalEntNaoDev1 = Nothing
        
    Else
    
        Set objItensNFiscalEntNaoDev1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objItensNFiscalEntNaoDevLista.Trata_Parametros(colSelecao, objItensNFiscalEntNaoDev1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19510
            
    If lErro <> SUCESSO Then Error 19511
        
    X.objItensNFiscalEntNaoDevLista.Show
            
    ItensNFiscalEntNaoDevLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ItensNFiscalEntNaoDevLista_Chama_Tela:

    ItensNFiscalEntNaoDevLista_Chama_Tela = Err

    Select Case Err
    
        Case 19509, 19510
            Unload X.objItensNFiscalEntNaoDevLista
            
        Case 19511
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146922)
        
    End Select
        
    Exit Function

End Function

Function ProdutoLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProduto1 As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutoLista_Consulta_Chama_Tela

    Load X.objProdutoLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24612
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objProduto1 = Nothing
        
    Else
    
        Set objProduto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutoLista_Consulta.Trata_Parametros(colSelecao, objProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24613
            
    If lErro <> SUCESSO Then Error 24614
        
    X.objProdutoLista_Consulta.Show
            
    ProdutoLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutoLista_Consulta_Chama_Tela:

    ProdutoLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24612, 24613
            Unload X.objProdutoLista_Consulta
            
        Case 24614
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146923)
        
    End Select
        
    Exit Function

End Function

Function AlmoxarifadoLista_Consulta_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAlmoxarifado1 As ClassAlmoxarifado
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AlmoxarifadoLista_Consulta_Chama_Tela

    Load X.objAlmoxarifadoLista_Consulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24738
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objAlmoxarifado1 = Nothing
        
    Else
    
        Set objAlmoxarifado1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objAlmoxarifadoLista_Consulta.Trata_Parametros(colSelecao, objAlmoxarifado1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24739
            
    If lErro <> SUCESSO Then Error 24740
        
    X.objAlmoxarifadoLista_Consulta.Show
            
    AlmoxarifadoLista_Consulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlmoxarifadoLista_Consulta_Chama_Tela:

    AlmoxarifadoLista_Consulta_Chama_Tela = Err

    Select Case Err
    
        Case 24738, 24739
            Unload X.objAlmoxarifadoLista_Consulta
            
        Case 24740
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146924)
        
    End Select
        
    Exit Function

End Function

Function AlmoxarifadoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objAlmoxarifado1 As ClassAlmoxarifado
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_AlmoxarifadoLista_Chama_Tela

    Load X.objAlmoxarifadoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24741
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objAlmoxarifado1 = Nothing
        
    Else
    
        Set objAlmoxarifado1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objAlmoxarifadoLista.Trata_Parametros(colSelecao, objAlmoxarifado1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24742
            
    If lErro <> SUCESSO Then Error 24743
        
    X.objAlmoxarifadoLista.Show
            
    AlmoxarifadoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_AlmoxarifadoLista_Chama_Tela:

    AlmoxarifadoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24741, 24742
            Unload X.objAlmoxarifadoLista
            
        Case 24743
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146925)
        
    End Select
        
    Exit Function

End Function

Function ReservaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objReserva1 As ClassReserva
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ReservaLista_Chama_Tela

    Load X.objReservaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24495
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objReserva1 = Nothing
        
    Else
    
        Set objReserva1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objReservaLista.Trata_Parametros(colSelecao, objReserva1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24496
            
    If lErro <> SUCESSO Then Error 24497
        
    X.objReservaLista.Show
            
    ReservaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ReservaLista_Chama_Tela:

    ReservaLista_Chama_Tela = Err

    Select Case Err
    
        Case 24495, 24496
            Unload X.objReservaLista
            
        Case 24497
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146926)
        
    End Select
        
    Exit Function

End Function

Function EmpenhoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objEmpenho1 As ClassEmpenho
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_EmpenhoLista_Chama_Tela

    Load X.objEmpenhoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29663
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objEmpenho1 = Nothing
        
    Else
    
        Set objEmpenho1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objEmpenhoLista.Trata_Parametros(colSelecao, objEmpenho1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29664
            
    If lErro <> SUCESSO Then Error 29665
        
    X.objEmpenhoLista.Show
            
    EmpenhoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EmpenhoLista_Chama_Tela:

    EmpenhoLista_Chama_Tela = Err

    Select Case Err
    
        Case 29663, 29664
            Unload X.objEmpenhoLista
            
        Case 29665
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146927)
        
    End Select
        
    Exit Function

End Function

Function ItemOrdemProducaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItemOrdemProducao1 As ClassItemOP
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ItemOrdemProducaoLista_Chama_Tela

    Load X.objItemOrdemProducaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29666
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objItemOrdemProducao1 = Nothing
        
    Else
    
        Set objItemOrdemProducao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objItemOrdemProducaoLista.Trata_Parametros(colSelecao, objItemOrdemProducao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29667
            
    If lErro <> SUCESSO Then Error 29668
        
    X.objItemOrdemProducaoLista.Show
            
    ItemOrdemProducaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ItemOrdemProducaoLista_Chama_Tela:

    ItemOrdemProducaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 29666, 29667
            Unload X.objItemOrdemProducaoLista
            
        Case 29668
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146928)
        
    End Select
        
    Exit Function

End Function

Function FaltaEstoque_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim lErro As Long
Dim objItemPedido As ClassItemPedido
Dim dQuantReservaEstoque As Double
Dim dSaldoAlmoxarifadoPadrao As Double
Dim objAlmoxarifadoPadrao As ClassAlmoxarifado
Dim colItemPedido As colItemPedido

On Error GoTo Erro_FaltaEstoque_Chama_Tela

    Load X.objFaltaEstoque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24638
            
    If IsMissing(vP1) Then
    
        Set objItemPedido = Nothing
        
    Else
    
        Set objItemPedido = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set colItemPedido = Nothing
        
    Else
    
        Set colItemPedido = vP2
        
    End If
    
    If vP3 = 0 Then
    
        dQuantReservaEstoque = 0
            
    Else
    
        dQuantReservaEstoque = vP3
        
    End If
    
    If IsMissing(vP4) Then
    
        Set objAlmoxarifadoPadrao = Nothing
            
    Else
    
        Set objAlmoxarifadoPadrao = vP4
        
    End If
    
    If vP5 = 0 Then
    
        dSaldoAlmoxarifadoPadrao = 0
            
    Else
    
        dSaldoAlmoxarifadoPadrao = vP5
        
    End If
    
    lErro = X.objFaltaEstoque.Trata_Parametros(objItemPedido, colItemPedido, dQuantReservaEstoque, objAlmoxarifadoPadrao, dSaldoAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24639
            
    If lErro <> SUCESSO Then Error 24640
        
    X.objFaltaEstoque.Show vbModal
            
    FaltaEstoque_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FaltaEstoque_Chama_Tela:

    FaltaEstoque_Chama_Tela = Err

    Select Case Err
    
        Case 24638, 24639
            Unload X.objFaltaEstoque
            
        Case 24640
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146929)
        
    End Select
        
    Exit Function

End Function

Function RelOpProdutos_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpProdutos_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 24952
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
    
    End If
    
    lErro = X.objRelOpProdutos.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 24953
    
    Load X.objRelOpProdutos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24954
            
    X.objRelOpProdutos.Show
            
    RelOpProdutos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpProdutos_Chama_Tela:

    RelOpProdutos_Chama_Tela = Err

    Select Case Err
    
        Case 24952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 24953
    
        Case 24954
            Unload X.objRelOpProdutos
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146930)
        
    End Select
        
    Exit Function

End Function

Function RelOpSaldoEstoque_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpSaldoEstoque_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 24956
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpSaldoEstoque.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 24957
    
    Load X.objRelOpSaldoEstoque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24958
            
    X.objRelOpSaldoEstoque.Show
            
    RelOpSaldoEstoque_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpSaldoEstoque_Chama_Tela:

    RelOpSaldoEstoque_Chama_Tela = Err

    Select Case Err
    
        Case 24956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 24957
    
        Case 24958
            Unload X.objRelOpSaldoEstoque
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146931)
        
    End Select
        
    Exit Function

End Function

Function RelOpAnaliseMovEst_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpAnaliseMovEst_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 24965
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpAnaliseMovEst.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 24966
    
    Load X.objRelOpAnaliseMovEst
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24975
            
    X.objRelOpAnaliseMovEst.Show
            
    RelOpAnaliseMovEst_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpAnaliseMovEst_Chama_Tela:

    RelOpAnaliseMovEst_Chama_Tela = Err

    Select Case Err
    
        Case 24965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 24966
    
        Case 24975
            Unload X.objRelOpAnaliseMovEst
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146932)
        
    End Select
        
    Exit Function

End Function

Function RelOpBoletimEntrada_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBoletimEntrada_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 24988
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpBoletimEntrada.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 24989
    
    Load X.objRelOpBoletimEntrada
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24990
            
    X.objRelOpBoletimEntrada.Show
            
    RelOpBoletimEntrada_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBoletimEntrada_Chama_Tela:

    RelOpBoletimEntrada_Chama_Tela = Err

    Select Case Err
    
        Case 24988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 24989
    
        Case 24990
            Unload X.objRelOpBoletimEntrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146933)
        
    End Select
        
    Exit Function

End Function

Function RelOpCustoRep_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpCustoRep_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29550
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpCustoRep.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29551
    
    Load X.objRelOpCustoRep
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29552
            
    X.objRelOpCustoRep.Show
            
    RelOpCustoRep_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpCustoRep_Chama_Tela:

    RelOpCustoRep_Chama_Tela = Err

    Select Case Err
    
        Case 29550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29551
    
        Case 29552
            Unload X.objRelOpCustoRep
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146934)
        
    End Select
        
    Exit Function

End Function

Function RelOpEtiqInv_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpEtiqInv_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29553
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpEtiqInv.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29554
    
    Load X.objRelOpEtiqInv
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29555
            
    X.objRelOpEtiqInv.Show
            
    RelOpEtiqInv_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpEtiqInv_Chama_Tela:

    RelOpEtiqInv_Chama_Tela = Err

    Select Case Err
    
        Case 29553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29554
    
        Case 29555
            Unload X.objRelOpEtiqInv
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146935)
        
    End Select
        
    Exit Function

End Function

Function RelOpFaltas_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpFaltas_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29556
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpFaltas.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29557
    
    Load X.objRelOpFaltas
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29558
            
    X.objRelOpFaltas.Show
            
    RelOpFaltas_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpFaltas_Chama_Tela:

    RelOpFaltas_Chama_Tela = Err

    Select Case Err
    
        Case 29556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29557
    
        Case 29558
            Unload X.objRelOpFaltas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146936)
        
    End Select
        
    Exit Function

End Function

Function RelOpKardex_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpKardex_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29571
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpKardex.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29572
    
    Load X.objRelOpKardex
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29573
            
    X.objRelOpKardex.Show
            
    RelOpKardex_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpKardex_Chama_Tela:

    RelOpKardex_Chama_Tela = Err

    Select Case Err
    
        Case 29571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29572
    
        Case 29573
            Unload X.objRelOpKardex
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146937)
        
    End Select
        
    Exit Function

End Function

Function RelOpKardexDia_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpKardexDia_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29574
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpKardexDia.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29575
    
    Load X.objRelOpKardexDia
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29576
            
    X.objRelOpKardexDia.Show
            
    RelOpKardexDia_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpKardexDia_Chama_Tela:

    RelOpKardexDia_Chama_Tela = Err

    Select Case Err
    
        Case 29574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29575
    
        Case 29576
            Unload X.objRelOpKardexDia
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146938)
        
    End Select
        
    Exit Function

End Function

Function RelOpListInv_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpListInv_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29577
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpListInv.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29578
    
    Load X.objRelOpListInv
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29579
            
    X.objRelOpListInv.Show
            
    RelOpListInv_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpListInv_Chama_Tela:

    RelOpListInv_Chama_Tela = Err

    Select Case Err
    
        Case 29577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29578
    
        Case 29579
            Unload X.objRelOpListInv
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146939)
        
    End Select
        
    Exit Function

End Function

Function RelOpMovInt_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpMovInt_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29580
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpMovInt.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29581
    
    Load X.objRelOpMovInt
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29582
            
    X.objRelOpMovInt.Show
            
    RelOpMovInt_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpMovInt_Chama_Tela:

    RelOpMovInt_Chama_Tela = Err

    Select Case Err
    
        Case 29580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29581
    
        Case 29582
            Unload X.objRelOpMovInt
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146940)
        
    End Select
        
    Exit Function

End Function

Function RelOpOrdemProducao_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpOrdemProducao_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29583
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpOrdemProducao.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29584
    
    Load X.objRelOpOrdemProducao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29585
            
    X.objRelOpOrdemProducao.Show
            
    RelOpOrdemProducao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpOrdemProducao_Chama_Tela:

    RelOpOrdemProducao_Chama_Tela = Err

    Select Case Err
    
        Case 29583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29584
    
        Case 29585
            Unload X.objRelOpOrdemProducao
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146941)
        
    End Select
        
    Exit Function

End Function

Function RelOpPickList_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPickList_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29586
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPickList.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29587
    
    Load X.objRelOpPickList
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29588
            
    X.objRelOpPickList.Show
            
    RelOpPickList_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPickList_Chama_Tela:

    RelOpPickList_Chama_Tela = Err

    Select Case Err
    
        Case 29586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29587
    
        Case 29588
            Unload X.objRelOpPickList
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146942)
        
    End Select
        
    Exit Function

End Function

Function RelOpProdVend_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpProdVend_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29602
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpProdVend.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29603
    
    Load X.objRelOpProdVend
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29604
            
    X.objRelOpProdVend.Show
            
    RelOpProdVend_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpProdVend_Chama_Tela:

    RelOpProdVend_Chama_Tela = Err

    Select Case Err
    
        Case 29602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29603
    
        Case 29604
            Unload X.objRelOpProdVend
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146943)
        
    End Select
        
    Exit Function

End Function

Function RelOpReqConsumo_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpReqConsumo_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29605
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpReqConsumo.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29606
    
    Load X.objRelOpReqConsumo
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29607
            
    X.objRelOpReqConsumo.Show
            
    RelOpReqConsumo_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpReqConsumo_Chama_Tela:

    RelOpReqConsumo_Chama_Tela = Err

    Select Case Err
    
        Case 29605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29606
    
        Case 29607
            Unload X.objRelOpReqConsumo
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146944)
        
    End Select
        
    Exit Function

End Function

Function RelOpResESValor_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpResESValor_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29608
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpResESValor.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29609
    
    Load X.objRelOpResESValor
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29610
            
    X.objRelOpResESValor.Show
            
    RelOpResESValor_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpResESValor_Chama_Tela:

    RelOpResESValor_Chama_Tela = Err

    Select Case Err
    
        Case 29608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29609
    
        Case 29610
            Unload X.objRelOpResESValor
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146945)
        
    End Select
        
    Exit Function

End Function

Function RelOpVendasMes_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpVendasMes_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 29611
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpVendasMes.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 29612
    
    Load X.objRelOpVendasMes
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29613
            
    X.objRelOpVendasMes.Show
            
    RelOpVendasMes_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpVendasMes_Chama_Tela:

    RelOpVendasMes_Chama_Tela = Err

    Select Case Err
    
        Case 29611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 29612
    
        Case 29613
            Unload X.objRelOpVendasMes
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146946)
        
    End Select
        
    Exit Function

End Function

Function RelOpMovsEstOP_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpMovsEstOP_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 16939
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpMovsEstOP.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 16940
    
    Load X.objRelOpMovsEstOP
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16941
            
    X.objRelOpMovsEstOP.Show
            
    RelOpMovsEstOP_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpMovsEstOP_Chama_Tela:

    RelOpMovsEstOP_Chama_Tela = Err

    Select Case Err
    
        Case 16939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 16940
    
        Case 16941
            Unload X.objRelOpMovsEstOP
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146947)
        
    End Select
        
    Exit Function

End Function

Function RelOpRealStd_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRealStd_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 16942
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRealStd.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 16943
    
    Load X.objRelOpRealStd
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16944
            
    X.objRelOpRealStd.Show
            
    RelOpRealStd_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRealStd_Chama_Tela:

    RelOpRealStd_Chama_Tela = Err

    Select Case Err
    
        Case 16942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 16943
    
        Case 16944
            Unload X.objRelOpRealStd
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146948)
        
    End Select
        
    Exit Function

End Function

Function RelOpPontoPedido_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPontoPedido_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 16945
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPontoPedido.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 16946
    
    Load X.objRelOpPontoPedido
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 16947
            
    X.objRelOpPontoPedido.Show
            
    RelOpPontoPedido_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPontoPedido_Chama_Tela:

    RelOpPontoPedido_Chama_Tela = Err

    Select Case Err
    
        Case 16945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 16946
    
        Case 16947
            Unload X.objRelOpPontoPedido
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146949)
        
    End Select
        
    Exit Function

End Function

Function RelOpAnaliseEst_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpAnaliseEst_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45145
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpAnaliseEst.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45146
    
    Load X.objRelOpAnaliseEst
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45147
            
    X.objRelOpAnaliseEst.Show
            
    RelOpAnaliseEst_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpAnaliseEst_Chama_Tela:

    RelOpAnaliseEst_Chama_Tela = Err

    Select Case Err
    
        Case 45145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45146
    
        Case 45147
            Unload X.objRelOpAnaliseEst
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146950)
        
    End Select
        
    Exit Function

End Function

Function RelOpKit_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpKit_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45148
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpKit.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45149
    
    Load X.objRelOpKit
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45150
            
    X.objRelOpKit.Show
            
    RelOpKit_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpKit_Chama_Tela:

    RelOpKit_Chama_Tela = Err

    Select Case Err
    
        Case 45148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45149
    
        Case 45150
            Unload X.objRelOpKit
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146951)
        
    End Select
        
    Exit Function

End Function

Function RelOpUsoProd_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpUsoProd_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45151
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpUsoProd.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45152
    
    Load X.objRelOpUsoProd
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45153
            
    X.objRelOpUsoProd.Show
            
    RelOpUsoProd_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpUsoProd_Chama_Tela:

    RelOpUsoProd_Chama_Tela = Err

    Select Case Err
    
        Case 45151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45152
    
        Case 45153
            Unload X.objRelOpUsoProd
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146952)
        
    End Select
        
    Exit Function

End Function

Function RelOpResumoKardex_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpResumoKardex_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45154
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpResumoKardex.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45155
    
    Load X.objRelOpResumoKardex
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45156
            
    X.objRelOpResumoKardex.Show
            
    RelOpResumoKardex_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpResumoKardex_Chama_Tela:

    RelOpResumoKardex_Chama_Tela = Err

    Select Case Err
    
        Case 45154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45155
    
        Case 45156
            Unload X.objRelOpResumoKardex
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146953)
        
    End Select
        
    Exit Function

End Function

Function RelOpClassifABC_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpClassifABC_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45868
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpClassifABC.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45869
    
    Load X.objRelOpClassifABC
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45870
            
    X.objRelOpClassifABC.Show
            
    RelOpClassifABC_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpClassifABC_Chama_Tela:

    RelOpClassifABC_Chama_Tela = Err

    Select Case Err
    
        Case 45868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 45869
    
        Case 45870
            Unload X.objRelOpClassifABC
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146954)
        
    End Select
        
    Exit Function

End Function

Function RelOpRegInvMod7_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRegInvMod7_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 54577
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRegInvMod7.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 54578
    
    Load X.objRelOpRegInvMod7
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 54579
            
    X.objRelOpRegInvMod7.Show
            
    RelOpRegInvMod7_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRegInvMod7_Chama_Tela:

    RelOpRegInvMod7_Chama_Tela = Err

    Select Case Err
    
        Case 54577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 54578
    
        Case 54579
            Unload X.objRelOpRegInvMod7
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146955)
        
    End Select
        
    Exit Function

End Function

Function RelOpDemApInv_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDemApInv_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 54654
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDemApInv.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 54655
    
    Load X.objRelOpDemApInv
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 54656
            
    X.objRelOpDemApInv.Show
            
    RelOpDemApInv_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDemApInv_Chama_Tela:

    RelOpDemApInv_Chama_Tela = Err

    Select Case Err
    
        Case 54654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
        
        Case 54655
    
        Case 54656
            Unload X.objRelOpDemApInv
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146956)
        
    End Select
        
    Exit Function

End Function

Function FechamentoMesEst_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_FechamentoMesEst_Chama_Tela

    Load X.objFechamentoMesEst
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20538
            
    X.objFechamentoMesEst.Show
            
    FechamentoMesEst_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FechamentoMesEst_Chama_Tela:

    FechamentoMesEst_Chama_Tela = Err

    Select Case Err
    
        Case 20538
            Unload X.objFechamentoMesEst
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146957)
        
    End Select
        
    Exit Function

End Function

Function CustoProducao_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CustoProducao_Chama_Tela

    Load X.objCustoProducao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45250
            
    X.objCustoProducao.Show
            
    CustoProducao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CustoProducao_Chama_Tela:

    CustoProducao_Chama_Tela = Err

    Select Case Err
    
        Case 45250
            Unload X.objCustoProducao
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146958)
        
    End Select
        
    Exit Function

End Function

Function CustoMedioProducao_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CustoMedioProducao_Chama_Tela

    Load X.objCustoMedioProducao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45251
            
    X.objCustoMedioProducao.Show
            
    CustoMedioProducao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CustoMedioProducao_Chama_Tela:

    CustoMedioProducao_Chama_Tela = Err

    Select Case Err
    
        Case 45251
            Unload X.objCustoMedioProducao
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146959)
        
    End Select
        
    Exit Function

End Function

Function FechamentoMesEst1_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objEstoqueMes As ClassEstoqueMes

On Error GoTo Erro_FechamentoMesEst1_Chama_Tela

    Load X.objFechamentoMesEst1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20540
            
    Set objEstoqueMes = vP1
        
    lErro = X.objFechamentoMesEst1.Trata_Parametros(objEstoqueMes)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20541
            
    If lErro <> SUCESSO Then Error 20542
        
    X.objFechamentoMesEst1.Show vbModal
            
    FechamentoMesEst1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FechamentoMesEst1_Chama_Tela:

    FechamentoMesEst1_Chama_Tela = Err

    Select Case Err
    
        Case 20540, 20541
            Unload X.objFechamentoMesEst1
            
        Case 20542
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146960)
        
    End Select
        
    Exit Function

End Function

Function PlanoContaESTLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPlanoContaEST As ClassPlanoConta
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PlanoContaESTLista_Chama_Tela

    Load X.objPlanoContaESTLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45739
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPlanoContaEST = Nothing
        
    Else
    
        Set objPlanoContaEST = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objPlanoContaESTLista.Trata_Parametros(colSelecao, objPlanoContaEST, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45740
            
    If lErro <> SUCESSO Then Error 45741
        
    X.objPlanoContaESTLista.Show
            
    PlanoContaESTLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PlanoContaESTLista_Chama_Tela:

    PlanoContaESTLista_Chama_Tela = Err

    Select Case Err
    
        Case 45739, 45740
            Unload X.objPlanoContaESTLista
            
        Case 45741
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146961)
        
    End Select
        
    Exit Function

End Function

Function ItemOrdemProducao_ProdutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ItemOrdemProducao_ProdutoLista_Chama_Tela

    Load X.objItemOrdemProducao_ProdutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45853
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objItemOP = Nothing
        
    Else
    
        Set objItemOP = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objItemOrdemProducao_ProdutoLista.Trata_Parametros(colSelecao, objItemOP, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45854
            
    If lErro <> SUCESSO Then Error 45855
        
    X.objItemOrdemProducao_ProdutoLista.Show
            
    ItemOrdemProducao_ProdutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ItemOrdemProducao_ProdutoLista_Chama_Tela:

    ItemOrdemProducao_ProdutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 45853, 45854
            Unload X.objItemOrdemProducao_ProdutoLista
            
        Case 45855
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146962)
        
    End Select
        
    Exit Function

End Function

Function ItemOrdemProducao_OPLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ItemOrdemProducao_OPLista_Chama_Tela

    Load X.objItemOrdemProducao_OPLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45856
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objItemOP = Nothing
        
    Else
    
        Set objItemOP = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objItemOrdemProducao_OPLista.Trata_Parametros(colSelecao, objItemOP, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45857
            
    If lErro <> SUCESSO Then Error 45858
        
    X.objItemOrdemProducao_OPLista.Show
            
    ItemOrdemProducao_OPLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ItemOrdemProducao_OPLista_Chama_Tela:

    ItemOrdemProducao_OPLista_Chama_Tela = Err

    Select Case Err
    
        Case 45856, 45857
            Unload X.objItemOrdemProducao_OPLista
            
        Case 45858
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146963)
        
    End Select
        
    Exit Function

End Function

Function SegmentosMAT_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objSegmento As ClassSegmento

On Error GoTo Erro_SegmentosMAT_Chama_Tela

    Load X.objSegmentosMAT
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44907
            
    If IsMissing(vP1) Then
    
        Set objSegmento = Nothing
        
    Else
    
        Set objSegmento = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSegmentosMAT.Trata_Parametros(objSegmento)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44908
            
        If lErro <> SUCESSO Then Error 44909
        
    End If
    
    X.objSegmentosMAT.Show
            
    SegmentosMAT_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SegmentosMAT_Chama_Tela:

    SegmentosMAT_Chama_Tela = Err

    Select Case Err
    
        Case 44907
    
        Case 44908, 44909
            Unload X.objSegmentosMAT
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146964)
        
    End Select
        
    Exit Function

End Function



Function ProdutosSubstLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objProdutosSubstLista As ClassProduto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ProdutosSubstLista_Chama_Tela

    Load X.objProdutosSubstLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57290
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
    
    If IsMissing(vP2) Then
    
        Set objProdutosSubstLista = Nothing
        
    Else
        
        Set objProdutosSubstLista = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
        
        Set objEvento = vP3
        
    End If
    
    lErro = X.objProdutosSubstLista.Trata_Parametros(colSelecao, objProdutosSubstLista, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57291
    
    If lErro <> SUCESSO Then Error 57292
    
    X.objProdutosSubstLista.Show
    
    ProdutosSubstLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ProdutosSubstLista_Chama_Tela:

    ProdutosSubstLista_Chama_Tela = Err

    Select Case Err
    
        Case 57290, 57291
            Unload X.objProdutosSubstLista
            
        Case 57292
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146965)
        
    End Select
        
    Exit Function

End Function

Function RecebMaterialFCom_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRecebMaterialFCom As ClassNFiscal

On Error GoTo Erro_RecebMaterialFCom_Chama_Tela

    Load X.objRecebMaterialFCom
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51343
            
    If IsMissing(vP1) Then
    
        Set objRecebMaterialFCom = Nothing
        
    Else
    
        Set objRecebMaterialFCom = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRecebMaterialFCom.Trata_Parametros(objRecebMaterialFCom)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51344
            
        If lErro <> SUCESSO Then Error 51345
        
    End If
    
    X.objRecebMaterialFCom.Show
            
    RecebMaterialFCom_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RecebMaterialFCom_Chama_Tela:

    RecebMaterialFCom_Chama_Tela = Err

    Select Case Err
    
        Case 51343, 51344
            Unload X.objRecebMaterialFCom
            
        Case 51345
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146966)
        
    End Select
        
    Exit Function

End Function

