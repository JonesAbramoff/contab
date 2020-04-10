VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassChamaTelaTRB"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function TipoDeTributacao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objTipoDeTributacao As ClassTipoDeTributacaoMovto

On Error GoTo Erro_TipoDeTributacao_Chama_Tela

    Load X.objTipoDeTributacao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24599
            
    If IsMissing(vP1) Then
    
        Set objTipoDeTributacao = Nothing
        
    Else
    
        Set objTipoDeTributacao = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objTipoDeTributacao.Trata_Parametros(objTipoDeTributacao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24600
            
        If lErro <> SUCESSO Then Error 24601
        
    End If
    
    X.objTipoDeTributacao.Show
            
    TipoDeTributacao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TipoDeTributacao_Chama_Tela:

    TipoDeTributacao_Chama_Tela = Err

    Select Case Err
    
        Case 24599, 24600
            Unload X.objTipoDeTributacao
            
        Case 24601
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147071)
        
    End Select
        
    Exit Function

End Function

Function PadraoTribSaida_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_PadraoTribSaida_Chama_Tela

    Load X.objPadraoTribSaida
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29629
    
    X.objPadraoTribSaida.Show
            
    PadraoTribSaida_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadraoTribSaida_Chama_Tela:

    PadraoTribSaida_Chama_Tela = Err

    Select Case Err
    
        Case 29629
            Unload X.objPadraoTribSaida
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147072)
        
    End Select
        
    Exit Function

End Function

Function PadraoTribEntrada_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_PadraoTribEntrada_Chama_Tela

    Load X.objPadraoTribEnt
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29719
    
    X.objPadraoTribEnt.Show
            
    PadraoTribEntrada_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadraoTribEntrada_Chama_Tela:

    PadraoTribEntrada_Chama_Tela = Err

    Select Case Err
    
        Case 29719
            Unload X.objPadraoTribEnt
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147073)
        
    End Select
        
    Exit Function

End Function

Function TiposDeTributacaoMovtoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objTiposDeTributacaoMovto1 As ClassTipoDeTributacaoMovto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_TiposDeTributacaoMovtoLista_Chama_Tela

    Load X.objTiposDeTributacaoMovtoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24938
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objTiposDeTributacaoMovto1 = Nothing
        
    Else
    
        Set objTiposDeTributacaoMovto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objTiposDeTributacaoMovtoLista.Trata_Parametros(colSelecao, objTiposDeTributacaoMovto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24939
            
    If lErro <> SUCESSO Then Error 24940
        
    X.objTiposDeTributacaoMovtoLista.Show
            
    TiposDeTributacaoMovtoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TiposDeTributacaoMovtoLista_Chama_Tela:

    TiposDeTributacaoMovtoLista_Chama_Tela = Err

    Select Case Err
    
        Case 24938, 24939
            Unload X.objTiposDeTributacaoMovtoLista
            
        Case 24940
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147074)
        
    End Select
        
    Exit Function

End Function

Function PadroesTribSaidaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPadroesTributacao1 As ClassPadraoTribSaida
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PadroesTribSaidaLista_Chama_Tela

    Load X.objPadroesTribSaidaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29511
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPadroesTributacao1 = Nothing
        
    Else
    
        Set objPadroesTributacao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objPadroesTribSaidaLista.Trata_Parametros(colSelecao, objPadroesTributacao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29512
            
    If lErro <> SUCESSO Then Error 29513
        
    X.objPadroesTribSaidaLista.Show
            
    PadroesTribSaidaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadroesTribSaidaLista_Chama_Tela:

    PadroesTribSaidaLista_Chama_Tela = Err

    Select Case Err
    
        Case 29511, 29512
            Unload X.objPadroesTribSaidaLista
            
        Case 29513
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147075)
        
    End Select
        
    Exit Function

End Function

Function PadroesTribEntLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPadroesTribEnt1 As ClassPadraoTribEnt
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PadroesTribEntLista_Chama_Tela

    Load X.objPadroesTribEntLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 29731
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objPadroesTribEnt1 = Nothing
        
    Else
    
        Set objPadroesTribEnt1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If
    
    lErro = X.objPadroesTribEntLista.Trata_Parametros(colSelecao, objPadroesTribEnt1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 29732
            
    If lErro <> SUCESSO Then Error 29733
        
    X.objPadroesTribEntLista.Show
            
    PadroesTribEntLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadroesTribEntLista_Chama_Tela:

    PadroesTribEntLista_Chama_Tela = Err

    Select Case Err
    
        Case 29731, 29732
            Unload X.objPadroesTribEntLista
            
        Case 29733
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147076)
        
    End Select
        
    Exit Function

End Function

