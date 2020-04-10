VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaTelaCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function Requisitante_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRequisitante As ClassRequisitante

On Error GoTo Erro_Requisitante_Chama_Tela

    Load X.objRequisitante
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51297
            
    If IsMissing(vP1) Then
    
        Set objRequisitante = Nothing
        
    Else
    
        Set objRequisitante = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRequisitante.Trata_Parametros(objRequisitante)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51298
            
        If lErro <> SUCESSO Then Error 51299
        
    End If
    
    X.objRequisitante.Show
            
    Requisitante_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Requisitante_Chama_Tela:

    Requisitante_Chama_Tela = Err

    Select Case Err
    
        Case 51297, 51298
            Unload X.objRequisitante
            
        Case 51299
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146641)
        
    End Select
        
    Exit Function

End Function

Function Comprador_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objComprador As ClassComprador

On Error GoTo Erro_Comprador_Chama_Tela

    Load X.objComprador
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51153
            
    If IsMissing(vP1) Then
    
        Set objComprador = Nothing
        
    Else
    
        Set objComprador = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objComprador.Trata_Parametros(objComprador)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51154
            
        If lErro <> SUCESSO Then Error 51155
        
    End If
    
    X.objComprador.Show
            
    Comprador_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Comprador_Chama_Tela:

    Comprador_Chama_Tela = Err

    Select Case Err
    
        Case 51153, 51154
            Unload X.objComprador
            
        Case 51155
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146642)
        
    End Select
        
    Exit Function

End Function


Function TipoDeBloqueioPC_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoDeBloqueioPC As ClassTipoBloqueioPC

On Error GoTo Erro_TipoDeBloqueioPC_Chama_Tela

    Load X.objTipoDeBloqueioPC
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51156
            
    If IsMissing(vP1) Then
    
        Set objTipoDeBloqueioPC = Nothing
        
    Else
    
        Set objTipoDeBloqueioPC = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoDeBloqueioPC.Trata_Parametros(objTipoDeBloqueioPC)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51157
            
        If lErro <> SUCESSO Then Error 51158
        
    End If
    
    X.objTipoDeBloqueioPC.Show
            
    TipoDeBloqueioPC_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDeBloqueioPC_Chama_Tela:

    TipoDeBloqueioPC_Chama_Tela = Err

    Select Case Err
    
        Case 51156, 51157
            Unload X.objTipoDeBloqueioPC
            
        Case 51158
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146643)
        
    End Select
        
    Exit Function

End Function


Function Alcada_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objAlcada As ClassAlcada

On Error GoTo Erro_Alcada_Chama_Tela

    Load X.objAlcada
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51169
            
    If IsMissing(vP1) Then
    
        Set objAlcada = Nothing
        
    Else
    
        Set objAlcada = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objAlcada.Trata_Parametros(objAlcada)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51170
            
        If lErro <> SUCESSO Then Error 51171
        
    End If
    
    X.objAlcada.Show
            
    Alcada_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Alcada_Chama_Tela:

    Alcada_Chama_Tela = Err

    Select Case Err
    
        Case 51169, 51170
            Unload X.objAlcada
            
        Case 51171
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146644)
        
    End Select
        
    Exit Function

End Function

Function LiberaBloqueioPC_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLiberaBloqueioPC As ClassLiberacaoBloqueiosPC

On Error GoTo Erro_LiberaBloqueioPC_Chama_Tela

    Load X.objLiberaBloqueioPC
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51172
            
    If IsMissing(vP1) Then
    
        Set objLiberaBloqueioPC = Nothing
        
    Else
    
        Set objLiberaBloqueioPC = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLiberaBloqueioPC.Trata_Parametros(objLiberaBloqueioPC)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51173
            
        If lErro <> SUCESSO Then Error 51174
        
    End If
    
    X.objLiberaBloqueioPC.Show
            
    LiberaBloqueioPC_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LiberaBloqueioPC_Chama_Tela:

    LiberaBloqueioPC_Chama_Tela = Err

    Select Case Err
    
        Case 51172, 51173
            Unload X.objLiberaBloqueioPC
            
        Case 51174
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146645)
        
    End Select
        
    Exit Function

End Function

Function ConfiguraCOM_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ConfiguraCOM_Chama_Tela

    Load X.objConfiguraCOM
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51203
            
    X.objConfiguraCOM.Show
            
    ConfiguraCOM_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ConfiguraCOM_Chama_Tela:

    ConfiguraCOM_Chama_Tela = Err

    Select Case Err
    
        Case 51203
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146646)
        
    End Select
        
    Exit Function

End Function



Function Estoque2_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objEstoque As ClassProduto

On Error GoTo Erro_Estoque2_Chama_Tela

    Load X.objEstoque
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51206
            
    If IsMissing(vP1) Then
    
        Set objEstoque = Nothing
        
    Else
    
        Set objEstoque = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objEstoque.Trata_Parametros(objEstoque)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51207
            
        If lErro <> SUCESSO Then Error 51208
        
    End If
    
    X.objEstoque.Show
            
    Estoque2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Estoque2_Chama_Tela:

    Estoque2_Chama_Tela = Err

    Select Case Err
    
        Case 51206, 51207
            Unload X.objEstoque
            
        Case 51208
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146647)
        
    End Select
        
    Exit Function

End Function


Function PedidoCompras_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPedidoCompras As ClassPedidoCompras

On Error GoTo Erro_PedidoCompras_Chama_Tela

    Load X.objPedidoCompras
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51127
            
    If IsMissing(vP1) Then
    
        Set objPedidoCompras = Nothing
        
    Else
    
        Set objPedidoCompras = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPedidoCompras.Trata_Parametros(objPedidoCompras)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51128
            
        If lErro <> SUCESSO Then Error 51129
        
    End If
    
    X.objPedidoCompras.Show
            
    PedidoCompras_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoCompras_Chama_Tela:

    PedidoCompras_Chama_Tela = Err

    Select Case Err
    
        Case 51127, 51128
            Unload X.objPedidoCompras
            
        Case 51129
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146648)
        
    End Select
        
    Exit Function

End Function

Function PedidoCotacao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPedidoCotacao As ClassPedidoCotacao

On Error GoTo Erro_PedidoCotacao_Chama_Tela

    Load X.objPedidoCotacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51297
            
    If IsMissing(vP1) Then
    
        Set objPedidoCotacao = Nothing
        
    Else
    
        Set objPedidoCotacao = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPedidoCotacao.Trata_Parametros(objPedidoCotacao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51298
        If lErro <> SUCESSO Then Error 51299
        
    End If
    
    X.objPedidoCotacao.Show
            
    PedidoCotacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoCotacao_Chama_Tela:

    PedidoCotacao_Chama_Tela = Err

    Select Case Err
    
        Case 51297, 51298
            Unload X.objPedidoCotacao
            
        Case 51299
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146649)
        
    End Select
        
    Exit Function

End Function

Function CondPagto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPedidoCompra As ClassPedidoCompras

On Error GoTo Erro_CondPagto_Chama_Tela

    Load X.objCondPagto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51300
            
    If IsMissing(vP1) Then
    
        Set objPedidoCompra = Nothing
        
    Else
    
        Set objPedidoCompra = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCondPagto.Trata_Parametros(objPedidoCompra)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51301
        If lErro <> SUCESSO Then Error 51302
        
    End If
    
    X.objCondPagto.Show vbModal
            
    CondPagto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CondPagto_Chama_Tela:

    CondPagto_Chama_Tela = Err

    Select Case Err
    
        Case 51300, 51301
            Unload X.objCondPagto
            
        Case 51302
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146650)
        
    End Select
        
    Exit Function

End Function

Function FornFilialProduto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objFornFilialProduto As ClassFornecedorProdutoFF

On Error GoTo Erro_FornFilialProduto_Chama_Tela

    Load X.objFornFilialProduto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51303
            
    If IsMissing(vP1) Then
    
        Set objFornFilialProduto = Nothing
        
    Else
    
        Set objFornFilialProduto = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objFornFilialProduto.Trata_Parametros(objFornFilialProduto)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51304
        If lErro <> SUCESSO Then Error 51305
        
    End If
    
    X.objFornFilialProduto.Show
            
    FornFilialProduto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FornFilialProduto_Chama_Tela:

    FornFilialProduto_Chama_Tela = Err

    Select Case Err
    
        Case 51303, 51304
            Unload X.objFornFilialProduto
            
        Case 51305
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146651)
        
    End Select
        
    Exit Function

End Function
Function ReqComprasEnv_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras

On Error GoTo Erro_ReqComprasEnv_Chama_Tela

    Load X.objReqComprasEnv
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51297
            
    If IsMissing(vP1) Then
    
        Set objReqCompras = Nothing
        
    Else
    
        Set objReqCompras = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objReqComprasEnv.Trata_Parametros(objReqCompras)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51298
        If lErro <> SUCESSO Then Error 51299
        
    End If
    
    X.objReqComprasEnv.Show
            
    ReqComprasEnv_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ReqComprasEnv_Chama_Tela:

    ReqComprasEnv_Chama_Tela = Err

    Select Case Err
    
        Case 51297, 51298
            Unload X.objReqComprasEnv
            
        Case 51299
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146652)
        
    End Select
        
    Exit Function

End Function

Function PedComprasCons_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPedComprasCons As ClassPedidoCompras

On Error GoTo Erro_PedComprasCons_Chama_Tela

    Load X.objPedComprasCons
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51311
            
    If IsMissing(vP1) Then
    
        Set objPedComprasCons = Nothing
        
    Else
    
        Set objPedComprasCons = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPedComprasCons.Trata_Parametros(objPedComprasCons)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51312
            
        If lErro <> SUCESSO Then Error 51313
        
    End If
    
    X.objPedComprasCons.Show
            
    PedComprasCons_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedComprasCons_Chama_Tela:

    PedComprasCons_Chama_Tela = Err

    Select Case Err
    
        Case 51311, 51312
            Unload X.objPedComprasCons
            
        Case 51313
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146653)
        
    End Select
        
    Exit Function

End Function

Function BaixaPedCompras_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BaixaPedCompras_Chama_Tela

    Load X.objBaixaPedCompras
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51310
    
    X.objBaixaPedCompras.Show
            
    BaixaPedCompras_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaPedCompras_Chama_Tela:

    BaixaPedCompras_Chama_Tela = Err

    Select Case Err
    
        Case 51310
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146654)
        
    End Select
        
    Exit Function

End Function

Function RequisicaoModeloLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRequisicaoModelo1 As ClassRequisicaoModelo
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RequisicaoModeloLista_Chama_Tela

    Load X.objRequisicaoModeloLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57175
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objRequisicaoModelo1 = Nothing
        
    Else
        Set objRequisicaoModelo1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objRequisicaoModeloLista.Trata_Parametros(colSelecao, objRequisicaoModelo1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57176
            
    If lErro <> SUCESSO Then Error 57177
        
    X.objRequisicaoModeloLista.Show
            
    RequisicaoModeloLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RequisicaoModeloLista_Chama_Tela:

    RequisicaoModeloLista_Chama_Tela = Err

    Select Case Err
    
        Case 57175, 57176
            Unload X.objRequisicaoModeloLista
            
        Case 57177
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146655)
        
    End Select
        
    Exit Function

End Function

Function RequisitanteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRequisitante1 As ClassRequisitante
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RequisitanteLista_Chama_Tela

    Load X.objRequisitanteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57178
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objRequisitante1 = Nothing
        
    Else
        Set objRequisitante1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objRequisitanteLista.Trata_Parametros(colSelecao, objRequisitante1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57179
            
    If lErro <> SUCESSO Then Error 57180
        
    X.objRequisitanteLista.Show
            
    RequisitanteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RequisitanteLista_Chama_Tela:

    RequisitanteLista_Chama_Tela = Err

    Select Case Err
    
        Case 57178, 57179
            Unload X.objRequisitanteLista
            
        Case 57180
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146656)
        
    End Select
        
    Exit Function

End Function

Function ObservacaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objObservacao1 As ClassObservacao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ObservacaoLista_Chama_Tela

    Load X.objObservacaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57181
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objObservacao1 = Nothing
        
    Else
        Set objObservacao1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objObservacaoLista.Trata_Parametros(colSelecao, objObservacao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57182
            
    If lErro <> SUCESSO Then Error 57183
        
    X.objObservacaoLista.Show
            
    ObservacaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ObservacaoLista_Chama_Tela:

    ObservacaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 57181, 57182
            Unload X.objObservacaoLista
            
        Case 57183
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146657)
        
    End Select
        
    Exit Function

End Function
Function PedComprasNaoEnvLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedComprasNaoEnv1 As ClassPedidoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedComprasNaoEnvLista_Chama_Tela

    Load X.objPedComprasNaoEnvLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57187
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedComprasNaoEnv1 = Nothing
        
    Else
        Set objPedComprasNaoEnv1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedComprasNaoEnvLista.Trata_Parametros(colSelecao, objPedComprasNaoEnv1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57188
            
    If lErro <> SUCESSO Then Error 57189
        
    X.objPedComprasNaoEnvLista.Show
            
    PedComprasNaoEnvLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedComprasNaoEnvLista_Chama_Tela:

    PedComprasNaoEnvLista_Chama_Tela = Err

    Select Case Err
    
        Case 57187, 57188
            Unload X.objPedComprasNaoEnvLista
            
        Case 57189
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146658)
        
    End Select
        
    Exit Function

End Function

Function FiliaisFornProdutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objFiliaisFornProduto1 As ClassFornecedorProdutoFF
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_FiliaisFornProdutoLista_Chama_Tela

    Load X.objFiliaisFornProdutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57184
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objFiliaisFornProduto1 = Nothing
        
    Else
        Set objFiliaisFornProduto1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objFiliaisFornProdutoLista.Trata_Parametros(colSelecao, objFiliaisFornProduto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57185
            
    If lErro <> SUCESSO Then Error 57186
        
    X.objFiliaisFornProdutoLista.Show
            
    FiliaisFornProdutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FiliaisFornProdutoLista_Chama_Tela:

    FiliaisFornProdutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 57184, 57185
            Unload X.objFiliaisFornProdutoLista
            
        Case 57186
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146659)
        
    End Select
        
    Exit Function

End Function


Function PedComprasGeradoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedComprasGerado1 As ClassPedidoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedComprasGeradoLista_Chama_Tela

    Load X.objPedComprasGeradoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57190
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedComprasGerado1 = Nothing
        
    Else
        Set objPedComprasGerado1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedComprasGeradoLista.Trata_Parametros(colSelecao, objPedComprasGerado1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57191
            
    If lErro <> SUCESSO Then Error 57192
        
    X.objPedComprasGeradoLista.Show
            
    PedComprasGeradoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedComprasGeradoLista_Chama_Tela:

    PedComprasGeradoLista_Chama_Tela = Err

    Select Case Err
    
        Case 57190, 57191
            Unload X.objPedComprasGeradoLista
            
        Case 57192
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146660)
        
    End Select
        
    Exit Function

End Function

Function RequisicaoComprasLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRequisicaoCompras1 As ClassRequisicaoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RequisicaoComprasLista_Chama_Tela

    Load X.objRequisicaoComprasLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57193
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objRequisicaoCompras1 = Nothing
        
    Else
        Set objRequisicaoCompras1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objRequisicaoComprasLista.Trata_Parametros(colSelecao, objRequisicaoCompras1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57194
            
    If lErro <> SUCESSO Then Error 57195
        
    X.objRequisicaoComprasLista.Show
            
    RequisicaoComprasLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RequisicaoComprasLista_Chama_Tela:

    RequisicaoComprasLista_Chama_Tela = Err

    Select Case Err
    
        Case 57193, 57194
            Unload X.objRequisicaoComprasLista
            
        Case 57195
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146661)
        
    End Select
        
    Exit Function

End Function
Function ReqComprasEnvLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objReqComprasEnv1 As ClassRequisicaoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ReqComprasEnvLista_Chama_Tela

    Load X.objReqComprasEnvLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57196
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objReqComprasEnv1 = Nothing
        
    Else
        Set objReqComprasEnv1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objReqComprasEnvLista.Trata_Parametros(colSelecao, objReqComprasEnv1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57197
            
    If lErro <> SUCESSO Then Error 57198
        
    X.objReqComprasEnvLista.Show
            
    ReqComprasEnvLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ReqComprasEnvLista_Chama_Tela:

    ReqComprasEnvLista_Chama_Tela = Err

    Select Case Err
    
        Case 57196, 57197
            Unload X.objReqComprasEnvLista
            
        Case 57198
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146662)
        
    End Select
        
    Exit Function

End Function



Function PedComprasEnvLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedComprasEnv1 As ClassPedidoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedComprasEnvLista_Chama_Tela

    Load X.objPedComprasEnvLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57199
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedComprasEnv1 = Nothing
        
    Else
        Set objPedComprasEnv1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedComprasEnvLista.Trata_Parametros(colSelecao, objPedComprasEnv1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57200
    
    If lErro <> SUCESSO Then Error 57201
        
    X.objPedComprasEnvLista.Show
            
    PedComprasEnvLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedComprasEnvLista_Chama_Tela:

    PedComprasEnvLista_Chama_Tela = Err

    Select Case Err
    
        Case 57199, 57200
            Unload X.objPedComprasEnvLista
            
        Case 57201
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146663)
        
    End Select
        
    Exit Function

End Function


Function PedidoCotacaoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoCotacao1 As ClassPedidoCotacao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoCotacaoLista_Chama_Tela

    Load X.objPedidoCotacaoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57202
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedidoCotacao1 = Nothing
        
    Else
        Set objPedidoCotacao1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedidoCotacaoLista.Trata_Parametros(colSelecao, objPedidoCotacao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57203
            
    If lErro <> SUCESSO Then Error 57204
        
    X.objPedidoCotacaoLista.Show
            
    PedidoCotacaoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoCotacaoLista_Chama_Tela:

    PedidoCotacaoLista_Chama_Tela = Err

    Select Case Err
    
        Case 57202, 57203
            Unload X.objPedidoCotacaoLista
            
        Case 57204
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146664)
        
    End Select
        
    Exit Function

End Function

Function PedidoCotacaoAtualizarLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoCotacaoAtualizar1 As ClassPedidoCotacao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoCotacaoAtualizarLista_Chama_Tela

    Load X.objPedidoCotacaoAtualizarLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57205
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedidoCotacaoAtualizar1 = Nothing
        
    Else
        Set objPedidoCotacaoAtualizar1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedidoCotacaoAtualizarLista.Trata_Parametros(colSelecao, objPedidoCotacaoAtualizar1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57206
            
    If lErro <> SUCESSO Then Error 57207
        
    X.objPedidoCotacaoAtualizarLista.Show
            
    PedidoCotacaoAtualizarLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoCotacaoAtualizarLista_Chama_Tela:

    PedidoCotacaoAtualizarLista_Chama_Tela = Err

    Select Case Err
    
        Case 57205, 57206
            Unload X.objPedidoCotacaoAtualizarLista
            
        Case 57207
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146665)
        
    End Select
        
    Exit Function

End Function


Function PedidoCotacaoAtualizadosLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoCotacaoAtualizados1 As ClassPedidoCotacao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoCotacaoAtualizadosLista_Chama_Tela

    Load X.objPedidoCotacaoAtualizadosLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57208
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedidoCotacaoAtualizados1 = Nothing
        
    Else
        Set objPedidoCotacaoAtualizados1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedidoCotacaoAtualizadosLista.Trata_Parametros(colSelecao, objPedidoCotacaoAtualizados1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57209
            
    If lErro <> SUCESSO Then Error 57210
        
    X.objPedidoCotacaoAtualizadosLista.Show
            
    PedidoCotacaoAtualizadosLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoCotacaoAtualizadosLista_Chama_Tela:

    PedidoCotacaoAtualizadosLista_Chama_Tela = Err

    Select Case Err
    
        Case 57208, 57209
            Unload X.objPedidoCotacaoAtualizadosLista
            
        Case 57210
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146666)
        
    End Select
        
    Exit Function

End Function


Function PedidoComprasLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedidoCompras1 As ClassPedidoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedidoComprasLista_Chama_Tela

    Load X.objPedidoComprasLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57211
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedidoCompras1 = Nothing
        
    Else
        Set objPedidoCompras1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedidoComprasLista.Trata_Parametros(colSelecao, objPedidoCompras1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57212
            
    If lErro <> SUCESSO Then Error 57213
        
    X.objPedidoComprasLista.Show
            
    PedidoComprasLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedidoComprasLista_Chama_Tela:

    PedidoComprasLista_Chama_Tela = Err

    Select Case Err
    
        Case 57211, 57212
            Unload X.objPedidoComprasLista
            
        Case 57213
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146667)
        
    End Select
        
    Exit Function

End Function


Function PedComprasItemReqLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPedComprasItemReq1 As ClassPedidoCompras
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PedComprasItemReqLista_Chama_Tela

    Load X.objPedComprasItemReqLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 57214
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPedComprasItemReq1 = Nothing
        
    Else
        Set objPedComprasItemReq1 = vP2
    
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPedComprasItemReqLista.Trata_Parametros(colSelecao, objPedComprasItemReq1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 57215
            
    If lErro <> SUCESSO Then Error 57216
        
    X.objPedComprasItemReqLista.Show
            
    PedComprasItemReqLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedComprasItemReqLista_Chama_Tela:

    PedComprasItemReqLista_Chama_Tela = Err

    Select Case Err
    
        Case 57214, 57215
            Unload X.objPedComprasItemReqLista
            
        Case 57216
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146668)
        
    End Select
        
    Exit Function

End Function


Function ReqCompras_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objReqCompras As ClassRequisicaoCompras

On Error GoTo Erro_ReqCompras_Chama_Tela

    Load X.objReqCompras
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51321
            
    If IsMissing(vP1) Then
    
        Set objReqCompras = Nothing
        
    Else
    
        Set objReqCompras = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objReqCompras.Trata_Parametros(objReqCompras)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51322
        If lErro <> SUCESSO Then Error 51323
        
    End If
    
    X.objReqCompras.Show
            
    ReqCompras_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ReqCompras_Chama_Tela:

    ReqCompras_Chama_Tela = Err

    Select Case Err
    
        Case 51321, 51322
            Unload X.objReqCompras
            
        Case 51323
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146669)
        
    End Select
        
    Exit Function

End Function
Function BaixaReqCompras_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BaixaReqCompras_Chama_Tela

    Load X.objBaixaReqCompras
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51324
    
    X.objBaixaReqCompras.Show
            
    BaixaReqCompras_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BaixaReqCompras_Chama_Tela:

    BaixaReqCompras_Chama_Tela = Err

    Select Case Err
    
        Case 51324
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146670)
        
    End Select
        
    Exit Function

End Function

Function PedComprasGerado_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPedComprasGerado As ClassPedidoCompras

On Error GoTo Erro_PedComprasGerado_Chama_Tela

    Load X.objPedComprasGerado
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51326
            
    If IsMissing(vP1) Then
    
        Set objPedComprasGerado = Nothing
        
    Else
    
        Set objPedComprasGerado = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPedComprasGerado.Trata_Parametros(objPedComprasGerado)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 51327
            
        If lErro <> SUCESSO Then Error 51325
        
    End If
    
    X.objPedComprasGerado.Show
            
    PedComprasGerado_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PedComprasGerado_Chama_Tela:

    PedComprasGerado_Chama_Tela = Err

    Select Case Err
    
        Case 51325
        
        Case 51326, 51327
            Unload X.objPedComprasGerado
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146671)
        
    End Select
        
    Exit Function

End Function

Function GeracaoPedCompraGerCot_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoPedCompraGerCot_Chama_Tela

    Load X.objGeracaoPedCompraGerCot
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51328
    
    X.objGeracaoPedCompraGerCot.Show
            
    GeracaoPedCompraGerCot_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoPedCompraGerCot_Chama_Tela:

    GeracaoPedCompraGerCot_Chama_Tela = Err

    Select Case Err
    
        Case 51328
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146672)
        
    End Select
        
    Exit Function

End Function

Function GeracaoPedCompraReq_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoPedCompraReq_Chama_Tela

    Load X.objGeracaoPedCompraReq
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51329
    
    X.objGeracaoPedCompraReq.Show
            
    GeracaoPedCompraReq_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoPedCompraReq_Chama_Tela:

    GeracaoPedCompraReq_Chama_Tela = Err

    Select Case Err
    
        Case 51329
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146673)
        
    End Select
        
    Exit Function

End Function

Function GeracaoPedCompraConc_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoPedCompraConc_Chama_Tela

    Load X.objGeracaoPedCompraConc
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51330
    
    X.objGeracaoPedCompraConc.Show
            
    GeracaoPedCompraConc_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoPedCompraConc_Chama_Tela:

    GeracaoPedCompraConc_Chama_Tela = Err

    Select Case Err
    
        Case 51330
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146674)
        
    End Select
        
    Exit Function

End Function


Function GeracaoPedCompraAvulsa_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_GeracaoPedCompraAvulsa_Chama_Tela

    Load X.objGeracaoPedCompraAvulsa
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 51331
    
    X.objGeracaoPedCompraAvulsa.Show
            
    GeracaoPedCompraAvulsa_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_GeracaoPedCompraAvulsa_Chama_Tela:

    GeracaoPedCompraAvulsa_Chama_Tela = Err

    Select Case Err
    
        Case 51331
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146675)
        
    End Select
        
    Exit Function

End Function
