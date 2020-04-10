VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function Lancamentos_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe

On Error GoTo Erro_Lancamentos_Chama_Tela

    Load X.objLancamentos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9126
            
    If IsMissing(vP1) Then
    
        Set objLanc = Nothing
        
    Else
    
        Set objLanc = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLancamentos.Trata_Parametros(objLanc)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9127
            
        If lErro <> SUCESSO Then Error 9128
        
    End If
    
    X.objLancamentos.Show
            
    Lancamentos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Lancamentos_Chama_Tela:

    Lancamentos_Chama_Tela = Err

    Select Case Err
    
        Case 9126, 9127
            Unload X.objLancamentos
            
        Case 9128
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146566)
        
    End Select
        
    Exit Function

End Function

Function LancamentoConsulta_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLancamentoConsulta As ClassLancamento_Detalhe

On Error GoTo Erro_LancamentoConsulta_Chama_Tela

    Load X.objLancamentoConsulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24644
            
    If IsMissing(vP1) Then
    
        Set objLancamentoConsulta = Nothing
        
    Else
    
        Set objLancamentoConsulta = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLancamentoConsulta.Trata_Parametros(objLancamentoConsulta)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24645
            
        If lErro <> SUCESSO Then Error 24646
        
    End If
    
    X.objLancamentoConsulta.Show
            
    LancamentoConsulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LancamentoConsulta_Chama_Tela:

    LancamentoConsulta_Chama_Tela = Err

    Select Case Err
    
        Case 24644, 24645
            Unload X.objLancamentoConsulta
            
        Case 24646
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146567)
        
    End Select
        
    Exit Function

End Function

Function LoteConsulta_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLoteConsulta As ClassLote

On Error GoTo Erro_LoteConsulta_Chama_Tela

    Load X.objLoteConsulta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24647
            
    If IsMissing(vP1) Then
    
        Set objLoteConsulta = Nothing
        
    Else
    
        Set objLoteConsulta = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLoteConsulta.Trata_Parametros(objLoteConsulta)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24648
            
        If lErro <> SUCESSO Then Error 24649
        
    End If
    
    X.objLoteConsulta.Show
            
    LoteConsulta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteConsulta_Chama_Tela:

    LoteConsulta_Chama_Tela = Err

    Select Case Err
    
        Case 24647, 24648
            Unload X.objLoteConsulta
            
        Case 24649
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146568)
        
    End Select
        
    Exit Function

End Function

Function LoteTela_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objLote1 As ClassLote


On Error GoTo Erro_LoteTela_Chama_Tela

    Load X.objLoteTela
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9131
            
    If IsMissing(vP1) Then
    
        Set objLote1 = Nothing
        
    Else
    
        Set objLote1 = vP1
        
    End If
    
    
    'se estiver com os dois parametros preenchidos
    If Not (IsMissing(vP1)) And Not (IsMissing(vP2)) Then
    
        lErro = X.objLoteTela.Trata_Parametros(objLote1, vP2)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20600
            
        If lErro <> SUCESSO Then Error 20601
        
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    ElseIf Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLoteTela.Trata_Parametros(objLote1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9132
            
        If lErro <> SUCESSO Then Error 9133
        
    End If
    
    X.objLoteTela.Show
            
    LoteTela_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteTela_Chama_Tela:

    LoteTela_Chama_Tela = Err

    Select Case Err
    
        Case 9131, 9132, 20600
            Unload X.objLoteTela
            
        Case 9133, 20601
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146569)
        
    End Select
        
    Exit Function

End Function

Function LotePendenteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLote1 As ClassLote
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LotePendenteLista_Chama_Tela

    Load X.objLotePendenteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9136
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objLote1 = Nothing
        
    Else
    
        Set objLote1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLotePendenteLista.Trata_Parametros(colSelecao, objLote1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9137
            
    If lErro <> SUCESSO Then Error 9138
        
    X.objLotePendenteLista.Show
            
    LotePendenteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LotePendenteLista_Chama_Tela:

    LotePendenteLista_Chama_Tela = Err

    Select Case Err
    
        Case 9136, 9137
            Unload X.objLotePendenteLista
            
        Case 9138
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146570)
        
    End Select
        
    Exit Function

End Function

Function LotePendenteListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLote1 As ClassLote
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LotePendenteListaModal_Chama_Tela

    Load X.objLotePendenteListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45043
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objLote1 = Nothing
        
    Else
    
        Set objLote1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLotePendenteListaModal.Trata_Parametros(colSelecao, objLote1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45044
            
    If lErro <> SUCESSO Then Error 45045
        
    X.objLotePendenteListaModal.Show vbModal
            
    LotePendenteListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LotePendenteListaModal_Chama_Tela:

    LotePendenteListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 45043, 45044
            Unload X.objLotePendenteListaModal
            
        Case 45045
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146571)
        
    End Select
        
    Exit Function

End Function

Function LoteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLote1 As ClassLote
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LoteLista_Chama_Tela

    Load X.objLoteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19859
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
                        
    If IsMissing(vP2) Then
    
        Set objLote1 = Nothing
        
    Else
    
        Set objLote1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLoteLista.Trata_Parametros(colSelecao, objLote1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19860
            
    If lErro <> SUCESSO Then Error 19861
        
    X.objLoteLista.Show
            
    LoteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteLista_Chama_Tela:

    LoteLista_Chama_Tela = Err

    Select Case Err
    
        Case 19859, 19860
            Unload X.objLoteLista
            
        Case 19861
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146572)
        
    End Select
        
    Exit Function

End Function

Function PlanoContaLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objPlanoConta1 As ClassPlanoConta
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_PlanoContaLista_Chama_Tela

    Load X.objPlanoContaLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14149
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objPlanoConta1 = Nothing
        
    Else
    
        Set objPlanoConta1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objPlanoContaLista.Trata_Parametros(colSelecao, objPlanoConta1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14150
            
    If lErro <> SUCESSO Then Error 14151
        
    X.objPlanoContaLista.Show
            
    PlanoContaLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PlanoContaLista_Chama_Tela:

    PlanoContaLista_Chama_Tela = Err

    Select Case Err
    
        Case 14149, 14150
            Unload X.objPlanoContaLista
            
        Case 14151
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146573)
        
    End Select
        
    Exit Function

End Function

Function RateioOnLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRateioOn1 As ClassRateioOn
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RateioOnLista_Chama_Tela

    Load X.objRateioOnLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14161
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objRateioOn1 = Nothing
        
    Else
    
        Set objRateioOn1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objRateioOnLista.Trata_Parametros(colSelecao, objRateioOn1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14162
            
    If lErro <> SUCESSO Then Error 14163
        
    X.objRateioOnLista.Show
            
    RateioOnLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RateioOnLista_Chama_Tela:

    RateioOnLista_Chama_Tela = Err

    Select Case Err
    
        Case 14161, 14162
            Unload X.objRateioOnLista
            
        Case 14163
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146574)
        
    End Select
        
    Exit Function

End Function

Function RateioLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRateio1 As ClassRateioOn
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RateioLista_Chama_Tela

    Load X.objRateioLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24426
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objRateio1 = Nothing
        
    Else
    
        Set objRateio1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objRateioLista.Trata_Parametros(colSelecao, objRateio1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24427
            
    If lErro <> SUCESSO Then Error 24428
        
    X.objRateioLista.Show vbModal
            
    RateioLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RateioLista_Chama_Tela:

    RateioLista_Chama_Tela = Err

    Select Case Err
    
        Case 24426, 24427
            Unload X.objRateioLista
            
        Case 24428
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146575)
        
    End Select
        
    Exit Function

End Function

Function RateioOffLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objRateioOff1 As ClassRateioOff
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_RateioOffLista_Chama_Tela

    Load X.objRateioOffLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14205
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objRateioOff1 = Nothing
        
    Else
    
        Set objRateioOff1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objRateioOffLista.Trata_Parametros(colSelecao, objRateioOff1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14206
            
    If lErro <> SUCESSO Then Error 14207
        
    X.objRateioOffLista.Show
            
    RateioOffLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RateioOffLista_Chama_Tela:

    RateioOffLista_Chama_Tela = Err

    Select Case Err
    
        Case 14205, 14206
            Unload X.objRateioOffLista
            
        Case 14207
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146576)
        
    End Select
        
    Exit Function

End Function

Function HistPadraoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objHistPadrao1 As ClassHistPadrao
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_HistPadraoLista_Chama_Tela

    Load X.objHistPadraoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14208
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objHistPadrao1 = Nothing
        
    Else
    
        Set objHistPadrao1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objHistPadraoLista.Trata_Parametros(colSelecao, objHistPadrao1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14209
            
    If lErro <> SUCESSO Then Error 14210
        
    X.objHistPadraoLista.Show
            
    HistPadraoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_HistPadraoLista_Chama_Tela:

    HistPadraoLista_Chama_Tela = Err

    Select Case Err
    
        Case 14208, 14209
            Unload X.objHistPadraoLista
            
        Case 14210
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146577)
        
    End Select
        
    Exit Function

End Function

Function OrcamentoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objOrcamento1 As ClassOrcamento
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_OrcamentoLista_Chama_Tela

    Load X.objOrcamentoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14175
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objOrcamento1 = Nothing
        
    Else
    
        Set objOrcamento1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objOrcamentoLista.Trata_Parametros(colSelecao, objOrcamento1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14174
            
    If lErro <> SUCESSO Then Error 14173
        
    X.objOrcamentoLista.Show
            
    OrcamentoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OrcamentoLista_Chama_Tela:

    OrcamentoLista_Chama_Tela = Err

    Select Case Err
    
        Case 14175, 14174
            Unload X.objOrcamentoLista
            
        Case 14173
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146578)
        
    End Select
        
    Exit Function

End Function

Function LoteEstorno1_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objLote As ClassLote
Dim objBrowseConfigura As AdmBrowseConfigura
Dim colSelecao As Collection

On Error GoTo Erro_LoteEstorno1_Chama_Tela

    Load X.objLoteEstorno1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45031
            
    If IsMissing(vP1) Then
    
        Set objLote = Nothing
        
    Else
    
        Set objLote = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objBrowseConfigura = Nothing
        
    Else
    
        Set objBrowseConfigura = vP2
        
    End If
    
    lErro = X.objLoteEstorno1.Trata_Parametros(objLote, objBrowseConfigura)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45032
            
    If lErro <> SUCESSO Then Error 45033
        
    X.objLoteEstorno1.Show vbModal
            
    LoteEstorno1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteEstorno1_Chama_Tela:

    LoteEstorno1_Chama_Tela = Err

    Select Case Err
    
        Case 45031, 45032
            Unload X.objLoteEstorno1
            
        Case 45033
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146579)
        
    End Select
        
    Exit Function

End Function

Function LancamentoEstorno1_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objLancamento As ClassLancamento_Cabecalho
Dim objBrowseConfigura As AdmBrowseConfigura
Dim colSelecao As Collection

On Error GoTo Erro_LancamentoEstorno1_Chama_Tela

    Load X.objLancamentoEstorno1
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45051
            
    If IsMissing(vP1) Then
    
        Set objLancamento = Nothing
        
    Else
    
        Set objLancamento = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objBrowseConfigura = Nothing
        
    Else
    
        Set objBrowseConfigura = vP2
        
    End If
    
    lErro = X.objLancamentoEstorno1.Trata_Parametros(objLancamento, objBrowseConfigura)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45052
            
    If lErro <> SUCESSO Then Error 45053
        
    X.objLancamentoEstorno1.Show vbModal
            
    LancamentoEstorno1_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LancamentoEstorno1_Chama_Tela:

    LancamentoEstorno1_Chama_Tela = Err

    Select Case Err
    
        Case 45051, 45052
            Unload X.objLancamentoEstorno1
            
        Case 45053
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146580)
        
    End Select
        
    Exit Function

End Function

Function LanpendenteLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLancamento_Detalhe1 As ClassLancamento_Detalhe
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LanpendenteLista_Chama_Tela

    Load X.objLanpendenteLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14041
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objLancamento_Detalhe1 = Nothing
        
    Else
    
        Set objLancamento_Detalhe1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLanpendenteLista.Trata_Parametros(colSelecao, objLancamento_Detalhe1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14042
            
    If lErro <> SUCESSO Then Error 14043
        
    X.objLanpendenteLista.Show
            
    LanpendenteLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LanpendenteLista_Chama_Tela:

    LanpendenteLista_Chama_Tela = Err

    Select Case Err
    
        Case 14041, 14042
            Unload X.objLanpendenteLista
            
        Case 14043
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146581)
        
    End Select
        
    Exit Function

End Function

Function LanpendenteListaModal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLancamento_Detalhe1 As ClassLancamento_Detalhe
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LanpendenteListaModal_Chama_Tela

    Load X.objLanpendenteListaModal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45069
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objLancamento_Detalhe1 = Nothing
        
    Else
    
        Set objLancamento_Detalhe1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLanpendenteListaModal.Trata_Parametros(colSelecao, objLancamento_Detalhe1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45070
            
    If lErro <> SUCESSO Then Error 45071
        
    X.objLanpendenteListaModal.Show vbModal
            
    LanpendenteListaModal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LanpendenteListaModal_Chama_Tela:

    LanpendenteListaModal_Chama_Tela = Err

    Select Case Err
    
        Case 45069, 45070
            Unload X.objLanpendenteListaModal
            
        Case 45071
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146582)
        
    End Select
        
    Exit Function

End Function

Function LancamentoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLancamento_Detalhe1 As ClassLancamento_Detalhe
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LancamentoLista_Chama_Tela

    Load X.objLancamentoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 19862
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
    If IsMissing(vP2) Then
    
        Set objLancamento_Detalhe1 = Nothing
        
    Else
    
        Set objLancamento_Detalhe1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLancamentoLista.Trata_Parametros(colSelecao, objLancamento_Detalhe1, objEvento)
    
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 19863
            
    If lErro <> SUCESSO Then Error 19864
        
    X.objLancamentoLista.Show
            
    LancamentoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LancamentoLista_Chama_Tela:

    LancamentoLista_Chama_Tela = Err

    Select Case Err
    
        Case 19862, 19863
            Unload X.objLancamentoLista
            
        Case 19864
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146583)
        
    End Select
        
    Exit Function

End Function

Function LanpendenteLista_Lote_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLancamento_Detalhe1 As ClassLancamento_Detalhe
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LanpendenteLista_Lote_Chama_Tela

    Load X.objLanpendenteLista_Lote
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14054
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objLancamento_Detalhe1 = Nothing
        
    Else
    
        Set objLancamento_Detalhe1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objLanpendenteLista_Lote.Trata_Parametros(colSelecao, objLancamento_Detalhe1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14055
            
    If lErro <> SUCESSO Then Error 14056
        
    X.objLanpendenteLista_Lote.Show
            
    LanpendenteLista_Lote_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LanpendenteLista_Lote_Chama_Tela:

    LanpendenteLista_Lote_Chama_Tela = Err

    Select Case Err
    
        Case 14054, 14055
            Unload X.objLanpendenteLista_Lote
            
        Case 14056
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146584)
        
    End Select
        
    Exit Function

End Function

Function LancamentoLista_Lote_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objLancamento_Detalhe1 As ClassLancamento_Detalhe
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_LancamentoLista_Lote_Chama_Tela

    Load X.objLancamentoLista_Lote
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 24663
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objLancamento_Detalhe1 = Nothing
        
    Else
    
        Set objLancamento_Detalhe1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
            
    Else
    
        Set objEvento = vP3
        
    End If

    lErro = X.objLancamentoLista_Lote.Trata_Parametros(colSelecao, objLancamento_Detalhe1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 24664
            
    If lErro <> SUCESSO Then Error 24665
        
    X.objLancamentoLista_Lote.Show
            
    LancamentoLista_Lote_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LancamentoLista_Lote_Chama_Tela:

    LancamentoLista_Lote_Chama_Tela = Err

    Select Case Err
    
        Case 24663, 24664
            Unload X.objLancamentoLista_Lote
            
        Case 24665
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146585)
        
    End Select
        
    Exit Function

End Function

Function DocAutoLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objDocAuto1 As ClassDocAuto
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_DocAutoLista_Chama_Tela

    Load X.objDocAutoLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14065
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objDocAuto1 = Nothing
        
    Else
    
        Set objDocAuto1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objDocAutoLista.Trata_Parametros(colSelecao, objDocAuto1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14066
            
    If lErro <> SUCESSO Then Error 14067
        
    X.objDocAutoLista.Show
            
    DocAutoLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DocAutoLista_Chama_Tela:

    DocAutoLista_Chama_Tela = Err

    Select Case Err
    
        Case 14065, 14066
            Unload X.objDocAutoLista
            
        Case 14067
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146586)
        
    End Select
        
    Exit Function

End Function

Function CclLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objCcl1 As ClassCcl
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_CclLista_Chama_Tela

    Load X.objCclLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14283
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objCcl1 = Nothing
        
    Else
    
        Set objCcl1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objCclLista.Trata_Parametros(colSelecao, objCcl1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14284
            
    If lErro <> SUCESSO Then Error 14285
        
    X.objCclLista.Show
            
    CclLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CclLista_Chama_Tela:

    CclLista_Chama_Tela = Err

    Select Case Err
    
        Case 14283, 14284
            Unload X.objCclLista
            
        Case 14285
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146587)
        
    End Select
        
    Exit Function

End Function

Function ContaCclLista_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant) As Long

Dim lErro As Long
Dim objContaCcl1 As ClassContaCcl
Dim objEvento As AdmEvento
Dim colSelecao As Collection

On Error GoTo Erro_ContaCclLista_Chama_Tela

    Load X.objContaCclLista
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14286
            
    If IsMissing(vP1) Then
    
        Set colSelecao = Nothing
        
    Else
    
        Set colSelecao = vP1
        
    End If
            
            
    If IsMissing(vP2) Then
    
        Set objContaCcl1 = Nothing
        
    Else
    
        Set objContaCcl1 = vP2
        
    End If
    
    If IsMissing(vP3) Then
    
        Set objEvento = Nothing
        
        
    Else
    
        Set objEvento = vP3
        
    End If
    
    
    lErro = X.objContaCclLista.Trata_Parametros(colSelecao, objContaCcl1, objEvento)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 14287
            
    If lErro <> SUCESSO Then Error 14288
        
    X.objContaCclLista.Show
            
    ContaCclLista_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ContaCclLista_Chama_Tela:

    ContaCclLista_Chama_Tela = Err

    Select Case Err
    
        Case 14286, 14287
            Unload X.objContaCclLista
            
        Case 14288
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146588)
        
    End Select
        
    Exit Function

End Function

Function CclTela_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objCcl1 As ClassCcl


On Error GoTo Erro_CclTela_Chama_Tela

    Load X.objCclTela
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9204
            
    If IsMissing(vP1) Then
    
        Set objCcl1 = Nothing
        
    Else
    
        Set objCcl1 = vP1
        
    End If
    
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCclTela.Trata_Parametros(objCcl1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9205
            
        If lErro <> SUCESSO Then Error 9206
        
    End If
        
    X.objCclTela.Show
            
    CclTela_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CclTela_Chama_Tela:

    CclTela_Chama_Tela = Err

    Select Case Err
    
        Case 9204
    
        Case 9205, 9206
            Unload X.objCclTela
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146589)
        
    End Select
        
    Exit Function

End Function

Function ContaCcl_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objContaCcl1 As ClassContaCcl

On Error GoTo Erro_ContaCcl_Chama_Tela

    Load X.objContaCcl
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9208
            
    If IsMissing(vP1) Then
    
        Set objContaCcl1 = Nothing
        
    Else
    
        Set objContaCcl1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objContaCcl.Trata_Parametros(objContaCcl1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9209
            
        If lErro <> SUCESSO Then Error 9210
        
    End If
    
    X.objContaCcl.Show
            
    ContaCcl_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ContaCcl_Chama_Tela:

    ContaCcl_Chama_Tela = Err

    Select Case Err
    
        Case 9208
    
        Case 9209, 9210
            Unload X.objContaCcl
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146590)
        
    End Select
        
    Exit Function

End Function

Function ContaCcl2_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objContaCcl As ClassContaCcl

On Error GoTo Erro_ContaCcl2_Chama_Tela

    Load X.objContaCcl2
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9212
            
    If IsMissing(vP1) Then
    
        Set objContaCcl = Nothing
        
    Else
    
        Set objContaCcl = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objContaCcl2.Trata_Parametros(objContaCcl)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9213
            
        If lErro <> SUCESSO Then Error 9214
        
    End If
    
    X.objContaCcl2.Show
            
    ContaCcl2_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ContaCcl2_Chama_Tela:

    ContaCcl2_Chama_Tela = Err

    Select Case Err
    
        Case 9212
    
        Case 9213, 9214
            Unload X.objContaCcl2
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146591)
        
    End Select
        
    Exit Function

End Function

Function DocAuto_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objDocAuto1 As ClassDocAuto

On Error GoTo Erro_DocAuto_Chama_Tela

    Load X.objDocAuto
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9216
            
    If IsMissing(vP1) Then
    
        Set objDocAuto1 = Nothing
        
    Else
    
        Set objDocAuto1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objDocAuto.Trata_Parametros(objDocAuto1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9217
            
        If lErro <> SUCESSO Then Error 9218
        
    End If
    
    X.objDocAuto.Show
            
    DocAuto_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_DocAuto_Chama_Tela:

    DocAuto_Chama_Tela = Err

    Select Case Err
    
        Case 9216
    
        Case 9217, 9218
            Unload X.objDocAuto
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146592)
        
    End Select
        
    Exit Function

End Function

Function ExercicioTela_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objExercicio As ClassExercicio

On Error GoTo Erro_ExercicioTela_Chama_Tela

    Load X.objExercicioTela
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9220
            
    If IsMissing(vP1) Then
    
        Set objExercicio = Nothing
        
    Else
    
        Set objExercicio = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objExercicioTela.Trata_Parametros(objExercicio)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9221
            
        If lErro <> SUCESSO Then Error 9222
        
    End If
     
    X.objExercicioTela.Show
            
    ExercicioTela_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExercicioTela_Chama_Tela:

    ExercicioTela_Chama_Tela = Err

    Select Case Err
    
        Case 9220
    
        Case 9221, 9222
            Unload X.objExercicioTela
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146593)
        
    End Select
        
    Exit Function

End Function

Function HistoricoPadrao_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_HistoricoPadrao_Chama_Tela

    Load X.objHistoricoPadrao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9224
            
    If IsMissing(vP1) Then
    
        Set objHistPadrao = Nothing
        
    Else
    
        Set objHistPadrao = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objHistoricoPadrao.Trata_Parametros(objHistPadrao)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9225
            
        If lErro <> SUCESSO Then Error 9226
        
    End If
    
    X.objHistoricoPadrao.Show
            
    HistoricoPadrao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_HistoricoPadrao_Chama_Tela:

    HistoricoPadrao_Chama_Tela = Err

    Select Case Err
    
        Case 9224
    
        Case 9225, 9226
            Unload X.objHistoricoPadrao
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146594)
        
    End Select
        
    Exit Function

End Function

Function PlanoConta_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPlanoConta1 As ClassPlanoConta

On Error GoTo Erro_PlanoConta_Chama_Tela

    Load X.objPlanoConta
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9228
            
    If IsMissing(vP1) Then
    
        Set objPlanoConta1 = Nothing
        
    Else
    
        Set objPlanoConta1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objPlanoConta.Trata_Parametros(objPlanoConta1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9229
            
        If lErro <> SUCESSO Then Error 9230
        
    End If
    
    X.objPlanoConta.Show
            
    PlanoConta_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PlanoConta_Chama_Tela:

    PlanoConta_Chama_Tela = Err

    Select Case Err
    
        Case 9228
    
        Case 9229, 9230
            Unload X.objPlanoConta
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146595)
        
    End Select
        
    Exit Function

End Function

Function RateioOff_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRateioOff1 As ClassRateioOff

On Error GoTo Erro_RateioOff_Chama_Tela

    Load X.objRateioOff
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9232
            
    If IsMissing(vP1) Then
    
        Set objRateioOff1 = Nothing
        
    Else
    
        Set objRateioOff1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRateioOff.Trata_Parametros(objRateioOff1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9233
            
        If lErro <> SUCESSO Then Error 9234
        
    End If
    
    X.objRateioOff.Show
            
    RateioOff_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RateioOff_Chama_Tela:

    RateioOff_Chama_Tela = Err

    Select Case Err
    
        Case 9232
    
        Case 9233, 9234
            Unload X.objRateioOff
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146596)
        
    End Select
        
    Exit Function

End Function

Function RateioOn_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objRateioOn1 As ClassRateioOn

On Error GoTo Erro_RateioOn_Chama_Tela

    Load X.objRateioOn
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9235
            
    If IsMissing(vP1) Then
    
        Set objRateioOn1 = Nothing
        
    Else
    
        Set objRateioOn1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRateioOn.Trata_Parametros(objRateioOn1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9236
            
        If lErro <> SUCESSO Then Error 9237
        
    End If
    
    X.objRateioOn.Show
            
    RateioOn_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RateioOn_Chama_Tela:

    RateioOn_Chama_Tela = Err

    Select Case Err
    
        Case 9235
    
        Case 9236, 9237
            Unload X.objRateioOn
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146597)
        
    End Select
        
    Exit Function

End Function

Function SegmentoTela_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objSegmento As ClassSegmento

On Error GoTo Erro_SegmentoTela_Chama_Tela

    Load X.objSegmentoTela
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9240
            
    If IsMissing(vP1) Then
    
        Set objSegmento = Nothing
        
    Else
    
        Set objSegmento = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSegmentoTela.Trata_Parametros(objSegmento)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9241
            
        If lErro <> SUCESSO Then Error 9242
        
    End If
    
    X.objSegmentoTela.Show
            
    SegmentoTela_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SegmentoTela_Chama_Tela:

    SegmentoTela_Chama_Tela = Err

    Select Case Err
    
        Case 9240
    
        Case 9241, 9242
            Unload X.objSegmentoTela
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146598)
        
    End Select
        
    Exit Function

End Function

Function RelOpRazao_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpRazao_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 7057
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpRazao.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 7055
    
    Load X.objRelOpRazao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 7054
            
    X.objRelOpRazao.Show
            
    RelOpRazao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpRazao_Chama_Tela:

    RelOpRazao_Chama_Tela = Err

    Select Case Err
    
        Case 7055
    
        Case 7054
            Unload X.objRelOpRazao
        
        Case 7057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146599)
        
    End Select
        
    Exit Function

End Function

Function RelOpBalPatr_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBalPatr_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14098
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
    
    End If
    
    lErro = X.objRelOpBalPatr.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14099
    
    Load X.objRelOpBalPatr
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14100
            
    X.objRelOpBalPatr.Show
            
    RelOpBalPatr_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBalPatr_Chama_Tela:

    RelOpBalPatr_Chama_Tela = Err

    Select Case Err
    
        Case 14099
    
        Case 14100
            Unload X.objRelOpBalPatr
        
        Case 14098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146600)
        
    End Select
        
    Exit Function

End Function

Function RelOpBalPatrPer_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBalPatrPer_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14101
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpBalPatrPer.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14102
    
    Load X.objRelOpBalPatrPer
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14103
            
    X.objRelOpBalPatrPer.Show
            
    RelOpBalPatrPer_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBalPatrPer_Chama_Tela:

    RelOpBalPatrPer_Chama_Tela = Err

    Select Case Err
    
        Case 14102
    
        Case 14103
            Unload X.objRelOpBalPatrPer
        
        Case 14101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146601)
        
    End Select
        
    Exit Function

End Function

Function RelOpBalVerif_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBalVerif_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14104
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpBalVerif.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14105
    
    Load X.objRelOpBalVerif
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14106
            
    X.objRelOpBalVerif.Show
            
    RelOpBalVerif_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBalVerif_Chama_Tela:

    RelOpBalVerif_Chama_Tela = Err

    Select Case Err
    
        Case 14105
    
        Case 14106
            Unload X.objRelOpBalVerif
        
        Case 14104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146602)
        
    End Select
        
    Exit Function

End Function

Function RelOpDemResEx_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDemResEx_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14107
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDemResEx.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14108
    
    Load X.objRelOpDemResEx
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14109
            
    X.objRelOpDemResEx.Show
            
    RelOpDemResEx_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDemResEx_Chama_Tela:

    RelOpDemResEx_Chama_Tela = Err

    Select Case Err
    
        Case 14108
    
        Case 14109
            Unload X.objRelOpDemResEx
        
        Case 14107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146603)
        
    End Select
        
    Exit Function

End Function

Function RelOpDemResPer_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDemResPer_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14110
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDemResPer.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14111
    
    Load X.objRelOpDemResPer
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14112
            
    X.objRelOpDemResPer.Show
            
    RelOpDemResPer_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDemResPer_Chama_Tela:

    RelOpDemResPer_Chama_Tela = Err

    Select Case Err
    
        Case 14111
    
        Case 14112
            Unload X.objRelOpDemResPer
        
        Case 14110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146604)
        
    End Select
        
    Exit Function

End Function

Function RelOpDespPerCcl_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDespPerCcl_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14113
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDespPerCcl.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14114
    
    Load X.objRelOpDespPerCcl
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14115
            
    X.objRelOpDespPerCcl.Show
            
    RelOpDespPerCcl_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDespPerCcl_Chama_Tela:

    RelOpDespPerCcl_Chama_Tela = Err

    Select Case Err
    
        Case 14114
    
        Case 14115
            Unload X.objRelOpDespPerCcl
        
        Case 14113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146605)
        
    End Select
        
    Exit Function

End Function

Function RelOpDiario_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpDiario_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14116
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpDiario.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14117
    
    Load X.objRelOpDiario
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14118
            
    X.objRelOpDiario.Show
            
    RelOpDiario_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpDiario_Chama_Tela:

    RelOpDiario_Chama_Tela = Err

    Select Case Err
    
        Case 14117
    
        Case 14118
            Unload X.objRelOpDiario
        
        Case 14116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146606)
        
    End Select
        
    Exit Function

End Function

Function RelOpLancCcl_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpLancCcl_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14119
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpLancCcl.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14120
    
    Load X.objRelOpLancCcl
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14121
            
    X.objRelOpLancCcl.Show
            
    RelOpLancCcl_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpLancCcl_Chama_Tela:

    RelOpLancCcl_Chama_Tela = Err

    Select Case Err
    
        Case 14120
    
        Case 14121
            Unload X.objRelOpLancCcl
        
        Case 14119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146607)
        
    End Select
        
    Exit Function

End Function

Function RelOpLancData_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpLancData_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14122
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpLancData.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14123
    
    Load X.objRelOpLancData
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14124
            
    X.objRelOpLancData.Show
            
    RelOpLancData_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpLancData_Chama_Tela:

    RelOpLancData_Chama_Tela = Err

    Select Case Err
    
        Case 14123
    
        Case 14124
            Unload X.objRelOpLancData
        
        Case 14122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146608)
        
    End Select
        
    Exit Function

End Function

Function RelOpLancLote_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpLancLote_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14125
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2

    End If
    
    lErro = X.objRelOpLancLote.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14126
    
    Load X.objRelOpLancLote
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14127
            
    X.objRelOpLancLote.Show
            
    RelOpLancLote_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpLancLote_Chama_Tela:

    RelOpLancLote_Chama_Tela = Err

    Select Case Err
    
        Case 14126
    
        Case 14127
            Unload X.objRelOpLancLote
        
        Case 14125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146609)
        
    End Select
        
    Exit Function

End Function

Function RelOpLote_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpLote_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14128
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpLote.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14129
    
    Load X.objRelOpLote
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14130
            
    X.objRelOpLote.Show
            
    RelOpLote_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpLote_Chama_Tela:

    RelOpLote_Chama_Tela = Err

    Select Case Err
    
        Case 14129
    
        Case 14130
            Unload X.objRelOpLote
        
        Case 14128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146610)
        
    End Select
        
    Exit Function

End Function

Function RelOpLotePend_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpLotePend_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 44105
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpLotePend.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 44106
    
    Load X.objRelOpLotePend
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44107
            
    X.objRelOpLotePend.Show
            
    RelOpLotePend_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpLotePend_Chama_Tela:

    RelOpLotePend_Chama_Tela = Err

    Select Case Err
    
        Case 44105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case 44106
    
        Case 44107
            Unload X.objRelOpLotePend
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146611)
        
    End Select
        
    Exit Function

End Function

Function RelOpLanPend_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpLanPend_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45094
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpLanPend.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45095
    
    Load X.objRelOpLanPend
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45096
            
    X.objRelOpLanPend.Show
            
    RelOpLanPend_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpLanPend_Chama_Tela:

    RelOpLanPend_Chama_Tela = Err

    Select Case Err
    
        Case 45094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case 45095
    
        Case 45096
            Unload X.objRelOpLanPend
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146612)
        
    End Select
        
    Exit Function

End Function

Function RelOpOrcRealCcl_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpOrcRealCcl_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45113
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpOrcRealCcl.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45114
    
    Load X.objRelOpOrcRealCcl
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45115
            
    X.objRelOpOrcRealCcl.Show
            
    RelOpOrcRealCcl_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpOrcRealCcl_Chama_Tela:

    RelOpOrcRealCcl_Chama_Tela = Err

    Select Case Err
    
        Case 45113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case 45114
    
        Case 45115
            Unload X.objRelOpOrcRealCcl
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146613)
        
    End Select
        
    Exit Function

End Function

Function RelOpOrcReal_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpOrcReal_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14131
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
                    
    End If
    
    lErro = X.objRelOpOrcReal.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14132
    
    Load X.objRelOpOrcReal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14133
            
    X.objRelOpOrcReal.Show
            
    RelOpOrcReal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpOrcReal_Chama_Tela:

    RelOpOrcReal_Chama_Tela = Err

    Select Case Err
    
        Case 14132
    
        Case 14133
            Unload X.objRelOpOrcReal
        
        Case 14131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146614)
        
    End Select
        
    Exit Function

End Function

Function RelOpPlanoContas_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPlanoContas_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14134
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPlanoContas.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14135
    
    Load X.objRelOpPlanoContas
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14136
            
    X.objRelOpPlanoContas.Show
            
    RelOpPlanoContas_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPlanoContas_Chama_Tela:

    RelOpPlanoContas_Chama_Tela = Err

    Select Case Err
    
        Case 14135
    
        Case 14136
            Unload X.objRelOpPlanoContas
        
        Case 14134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146615)
        
    End Select
        
    Exit Function

End Function

Function RelOpPlanSaldos_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpPlanSaldos_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 14137
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
        
    End If
    
    lErro = X.objRelOpPlanSaldos.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 14138
    
    Load X.objRelOpPlanSaldos
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 14139
            
    X.objRelOpPlanSaldos.Show
            
    RelOpPlanSaldos_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpPlanSaldos_Chama_Tela:

    RelOpPlanSaldos_Chama_Tela = Err

    Select Case Err
    
        Case 14138
    
        Case 14139
            Unload X.objRelOpPlanSaldos
        
        Case 14137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146616)
        
    End Select
        
    Exit Function

End Function

Function RelOpMutacaoPL_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpMutacaoPL_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45072
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
            
    End If
    
    lErro = X.objRelOpMutacaoPL.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45073
    
    Load X.objRelOpMutacaoPL
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45074
            
    X.objRelOpMutacaoPL.Show
            
    RelOpMutacaoPL_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpMutacaoPL_Chama_Tela:

    RelOpMutacaoPL_Chama_Tela = Err

    Select Case Err
    
        Case 45072
    
        Case 45073
            Unload X.objRelOpMutacaoPL
        
        Case 45074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146617)
        
    End Select
        
    Exit Function

End Function

Function RelOpBalPatrCom_Chama_Tela(X As Object, Optional vP1 As Variant, Optional vP2 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_RelOpBalPatrCom_Chama_Tela

    If IsMissing(vP1) Or IsMissing(vP2) Then
    
        Error 45088
        
    Else
    
        Set objRelatorio = vP1
        Set objRelOpcoes = vP2
    
    End If
    
    lErro = X.objRelOpBalPatrCom.Trata_Parametros(objRelatorio, objRelOpcoes)
    If lErro <> SUCESSO Then Error 45089
    
    Load X.objRelOpBalPatrCom
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45090
            
    X.objRelOpBalPatrCom.Show
            
    RelOpBalPatrCom_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelOpBalPatrCom_Chama_Tela:

    RelOpBalPatrCom_Chama_Tela = Err

    Select Case Err
    
        Case 45088
    
        Case 45089
            Unload X.objRelOpBalPatrCom
        
        Case 45090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146618)
        
    End Select
        
    Exit Function

End Function

Function Orcamento_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objOrcamento1 As ClassOrcamento

On Error GoTo Erro_Orcamento_Chama_Tela

    Load X.objOrcamento
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9278
            
    If IsMissing(vP1) Then
    
        Set objOrcamento1 = Nothing
        
    Else
    
        Set objOrcamento1 = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objOrcamento.Trata_Parametros(objOrcamento1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9279
            
        If lErro <> SUCESSO Then Error 9280
        
    End If
    
    X.objOrcamento.Show
            
    Orcamento_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Orcamento_Chama_Tela:

    Orcamento_Chama_Tela = Err

    Select Case Err
    
        Case 9278, 9279
            Unload X.objOrcamento
            
        Case 9280
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146619)
        
    End Select
        
    Exit Function

End Function

Function Configuracao_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Configuracao_Chama_Tela

    Load X.objConfiguracao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9282
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
            
        lErro = X.objConfiguracao.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9283
            
        If lErro <> SUCESSO Then Error 9284
        
    End If
    
    X.objConfiguracao.Show
            
    Configuracao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Configuracao_Chama_Tela:

    Configuracao_Chama_Tela = Err

    Select Case Err
    
        Case 9282, 9283
            Unload X.objConfiguracao
            
        Case 9284
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146620)
        
    End Select
        
    Exit Function

End Function

Function ConfiguracaoSetup_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ConfiguracaoSetup_Chama_Tela

    Load X.objConfiguracaoSetup
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9296
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objConfiguracaoSetup.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9297
            
        If lErro <> SUCESSO Then Error 9298
        
    End If
    
    X.objConfiguracaoSetup.Show
            
    ConfiguracaoSetup_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ConfiguracaoSetup_Chama_Tela:

    ConfiguracaoSetup_Chama_Tela = Err

    Select Case Err
    
        Case 9296, 9297
            Unload X.objConfiguracaoSetup
            
        Case 9298
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146621)
        
    End Select
        
    Exit Function

End Function

Function LoteAtualiza_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_LoteAtualiza_Chama_Tela

    Load X.objLoteAtualiza
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9361
            
    'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLoteAtualiza.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9362
            
        If lErro <> SUCESSO Then Error 9363
        
    End If
    
    X.objLoteAtualiza.Show
            
    LoteAtualiza_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteAtualiza_Chama_Tela:

    LoteAtualiza_Chama_Tela = Err

    Select Case Err
    
        Case 9361, 9362
            Unload X.objLoteAtualiza
            
        Case 9363
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146622)
        
    End Select
        
    Exit Function

End Function

Function Rateio_Chama_Tela(X As Object, vP1 As Variant, vP2 As Variant) As Long

Dim lErro As Long
Dim objRateioOn As ClassRateioOn
Dim objConfirmaTela As AdmConfirmaTela

On Error GoTo Erro_Rateio_Chama_Tela

    Load X.objRateio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9368
            
    Set objRateioOn = vP1
    
    Set objConfirmaTela = vP2
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objRateio.Trata_Parametros(objRateioOn, objConfirmaTela)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9369
            
        If lErro <> SUCESSO Then Error 9370
        
    End If
    
    X.objRateio.Show 1
            
    Rateio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Rateio_Chama_Tela:

    Rateio_Chama_Tela = Err

    Select Case Err
    
        Case 9368, 9369
            Unload X.objRateio
            
        Case 9370
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146623)
        
    End Select
        
    Exit Function

End Function

Function CriaExercicio_Chama_Tela(X As Object, vP1 As Variant) As Long

Dim lErro As Long
Dim objExercicio As ClassExercicio

On Error GoTo Erro_CriaExercicio_Chama_Tela

    Load X.objCriaExercicio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9416
            
    Set objExercicio = vP1
        
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objCriaExercicio.Trata_Parametros(objExercicio)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9417
            
        If lErro <> SUCESSO Then Error 9418
        
    End If
    
    X.objCriaExercicio.Show 1
            
    CriaExercicio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CriaExercicio_Chama_Tela:

    CriaExercicio_Chama_Tela = Err

    Select Case Err
    
        Case 9416
    
        Case 9417, 9418
            Unload X.objCriaExercicio
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146624)
        
    End Select
        
    Exit Function

End Function

Function ApuraExercicio_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ApuraExercicio_Chama_Tela

    Load X.objApuraExercicio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9466
            
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objApuraExercicio.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9467
            
        If lErro <> SUCESSO Then Error 9468
        
    End If
        
    X.objApuraExercicio.Show
            
    ApuraExercicio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ApuraExercicio_Chama_Tela:

    ApuraExercicio_Chama_Tela = Err

    Select Case Err
    
        Case 9466, 9467
            Unload X.objApuraExercicio
            
        Case 9468
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146625)
        
    End Select
        
    Exit Function

End Function

Function ApuraPeriodo_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ApuraPeriodo_Chama_Tela

    Load X.objApuraPeriodo
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9470
            
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objApuraPeriodo.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9471
            
        If lErro <> SUCESSO Then Error 9472
        
    End If
        
    X.objApuraPeriodo.Show
            
    ApuraPeriodo_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ApuraPeriodo_Chama_Tela:

    ApuraPeriodo_Chama_Tela = Err

    Select Case Err
    
        Case 9470, 9471
            Unload X.objApuraPeriodo
            
        Case 9472
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146626)
        
    End Select
        
    Exit Function

End Function

Function FechamentoExercicio_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_FechamentoExercicio_Chama_Tela

    Load X.objFechamentoExercicio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9502
            
    If lErro_Chama_Tela = SUCESSO Then
            
        lErro = X.objFechamentoExercicio.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9503
            
        If lErro <> SUCESSO Then Error 9504
        
    End If
        
    X.objFechamentoExercicio.Show
            
    FechamentoExercicio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_FechamentoExercicio_Chama_Tela:

    FechamentoExercicio_Chama_Tela = Err

    Select Case Err
    
        Case 9502, 9503
            Unload X.objFechamentoExercicio
            
        Case 9504
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146627)
        
    End Select
        
    Exit Function

End Function

Function ReaberturaExercicio_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ReaberturaExercicio_Chama_Tela

    Load X.objReaberturaExercicio
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9506
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objReaberturaExercicio.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9507
            
        If lErro <> SUCESSO Then Error 9508
        
    End If
    
    X.objReaberturaExercicio.Show
            
    ReaberturaExercicio_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ReaberturaExercicio_Chama_Tela:

    ReaberturaExercicio_Chama_Tela = Err

    Select Case Err
    
        Case 9506, 9507
            Unload X.objReaberturaExercicio
            
        Case 9508
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146628)
        
    End Select
        
    Exit Function

End Function

Function Reprocessamento_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Reprocessamento_Chama_Tela

    Load X.objReprocessamento
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9531
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objReprocessamento.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9532
            
        If lErro <> SUCESSO Then Error 9533
        
    End If
    
    X.objReprocessamento.Show
            
    Reprocessamento_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Reprocessamento_Chama_Tela:

    Reprocessamento_Chama_Tela = Err

    Select Case Err
    
        Case 9531, 9532
            Unload X.objReprocessamento
            
        Case 9533
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146629)
        
    End Select
        
    Exit Function

End Function

Function LancamentosAt_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_LancamentosAt_Chama_Tela

    Load X.objLancamentosAt
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9549
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLancamentosAt.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9550
            
        If lErro <> SUCESSO Then Error 9551
        
    End If
    
    X.objLancamentosAt.Show
            
    LancamentosAt_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LancamentosAt_Chama_Tela:

    LancamentosAt_Chama_Tela = Err

    Select Case Err
    
        Case 9549, 9550
            Unload X.objLancamentosAt
            
        Case 9551
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146630)
        
    End Select
        
    Exit Function

End Function

Function ContaCategoria_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objContaCategoria1 As ClassContaCategoria

On Error GoTo Erro_ContaCategoria_Chama_Tela

    Load X.objContaCategoria
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9684
            
    If IsMissing(vP1) Then
    
        Set objContaCategoria1 = Nothing
        
    Else
    
        Set objContaCategoria1 = vP1
        
    End If

     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objContaCategoria.Trata_Parametros(objContaCategoria1)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9685
            
        If lErro <> SUCESSO Then Error 9686
        
    End If
    
    X.objContaCategoria.Show
            
    ContaCategoria_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ContaCategoria_Chama_Tela:

    ContaCategoria_Chama_Tela = Err

    Select Case Err
    
        Case 9684, 9685
            Unload X.objContaCategoria
            
        Case 9686
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146631)
        
    End Select
        
    Exit Function

End Function

Function ContaConfigura_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_ContaConfigura_Chama_Tela

    Load X.objContaConfigura
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9760
            
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objContaConfigura.Trata_Parametros()
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9761
            
        If lErro <> SUCESSO Then Error 9762
        
    End If
    
    X.objContaConfigura.Show
            
    ContaConfigura_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ContaConfigura_Chama_Tela:

    ContaConfigura_Chama_Tela = Err

    Select Case Err
    
        Case 9760, 9761
            Unload X.objContaConfigura
            
        Case 9762
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146632)
        
    End Select
        
    Exit Function

End Function

Function SaldoInicial_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta

On Error GoTo Erro_SaldoInicial_Chama_Tela

    Load X.objSaldoInicial
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 10036
            
    If IsMissing(vP1) Then
    
        Set objPlanoConta = Nothing
        
    Else
    
        Set objPlanoConta = vP1
        
    End If

     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objSaldoInicial.Trata_Parametros(objPlanoConta)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 10037
            
        If lErro <> SUCESSO Then Error 10038
        
    End If
    
    X.objSaldoInicial.Show
            
    SaldoInicial_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_SaldoInicial_Chama_Tela:

    SaldoInicial_Chama_Tela = Err

    Select Case Err
    
        Case 10036, 10037
            Unload X.objSaldoInicial
            
        Case 10038
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146633)
        
    End Select
        
    Exit Function

End Function

Function LoteEstorno_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLoteEstorno As ClassLote

On Error GoTo Erro_LoteEstorno_Chama_Tela

    Load X.objLoteEstorno
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45028
            
    If IsMissing(vP1) Then
    
        Set objLoteEstorno = Nothing
        
    Else
    
        Set objLoteEstorno = vP1
        
    End If

     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLoteEstorno.Trata_Parametros(objLoteEstorno)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45029
            
        If lErro <> SUCESSO Then Error 45030
        
    End If
    
    X.objLoteEstorno.Show
            
    LoteEstorno_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LoteEstorno_Chama_Tela:

    LoteEstorno_Chama_Tela = Err

    Select Case Err
    
        Case 45028, 45029
            Unload X.objLoteEstorno
            
        Case 45030
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146634)
        
    End Select
        
    Exit Function

End Function

Function LancamentoEstorno_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objLancamentoEstorno As ClassLancamento_Cabecalho

On Error GoTo Erro_LancamentoEstorno_Chama_Tela

    Load X.objLancamentoEstorno
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45048
            
    If IsMissing(vP1) Then
    
        Set objLancamentoEstorno = Nothing
        
    Else
    
        Set objLancamentoEstorno = vP1
        
    End If

     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objLancamentoEstorno.Trata_Parametros(objLancamentoEstorno)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45049
            
        If lErro <> SUCESSO Then Error 45050
        
    End If
    
    X.objLancamentoEstorno.Show
            
    LancamentoEstorno_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LancamentoEstorno_Chama_Tela:

    LancamentoEstorno_Chama_Tela = Err

    Select Case Err
    
        Case 45048, 45049
            Unload X.objLancamentoEstorno
            
        Case 45050
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146635)
        
    End Select
        
    Exit Function

End Function

Function ExercicioFilial_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long
Dim objExercicio As ClassExercicio

On Error GoTo Erro_ExercicioFilial_Chama_Tela

    Load X.objExercicioFilial
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 10265
            
    If IsMissing(vP1) Then
    
        Set objExercicio = Nothing
        
    Else
    
        Set objExercicio = vP1
        
    End If
    
     'se estiver com o parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vP1)) Or lErro_Chama_Tela = SUCESSO Then
    
        lErro = X.objExercicioFilial.Trata_Parametros(objExercicio)
        If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 10266
            
        If lErro <> SUCESSO Then Error 10267
        
    End If
     
    X.objExercicioFilial.Show
            
    ExercicioFilial_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_ExercicioFilial_Chama_Tela:

    ExercicioFilial_Chama_Tela = Err

    Select Case Err
    
        Case 10265
    
        Case 10266, 10267
            Unload X.objExercicioFilial
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146636)
        
    End Select
        
    Exit Function

End Function

Function PadraoContab_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_PadraoContab_Chama_Tela

    Load X.objPadraoContab
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 20469
            
    lErro = X.objPadraoContab.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 20470
            
    If lErro <> SUCESSO Then Error 20471
        
    X.objPadraoContab.Show
            
    PadraoContab_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PadraoContab_Chama_Tela:

    PadraoContab_Chama_Tela = Err

    Select Case Err
    
        Case 20469
        
        Case 20470, 20471
            Unload X.objPadraoContab
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146637)
        
    End Select
        
    Exit Function

End Function

Function MnemonicoGlobal_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_MnemonicoGlobal_Chama_Tela

    Load X.objMnemonicoGlobal
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 36712
            
    lErro = X.objMnemonicoGlobal.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 36713
            
    If lErro <> SUCESSO Then Error 36714
        
    X.objMnemonicoGlobal.Show
            
    MnemonicoGlobal_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_MnemonicoGlobal_Chama_Tela:

    MnemonicoGlobal_Chama_Tela = Err

    Select Case Err
    
        Case 36712
        
        Case 36713, 36714
            Unload X.objMnemonicoGlobal
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146638)
        
    End Select
        
    Exit Function

End Function

Function RateioOffBatch_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_RateioOffBatch_Chama_Tela

    Load X.objRateioOffBatch
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 36800
            
    lErro = X.objRateioOffBatch.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 36801
            
    If lErro <> SUCESSO Then Error 36802
        
    X.objRateioOffBatch.Show
            
    RateioOffBatch_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RateioOffBatch_Chama_Tela:

    RateioOffBatch_Chama_Tela = Err

    Select Case Err
    
        Case 36800
        
        Case 36801, 36802
            Unload X.objRateioOffBatch
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146639)
        
    End Select
        
    Exit Function

End Function

Function RelDREConfig_Chama_Tela(X As Object, Optional vP1 As Variant) As Long

Dim lErro As Long

On Error GoTo Erro_RelDREConfig_Chama_Tela

    Load X.objRelDREConfig
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 41793
            
    lErro = X.objRelDREConfig.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 41794
            
    If lErro <> SUCESSO Then Error 41795
        
    X.objRelDREConfig.Show vbModal
            
    RelDREConfig_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelDREConfig_Chama_Tela:

    RelDREConfig_Chama_Tela = Err

    Select Case Err
    
        Case 41793
        
        Case 41794, 41795
            Unload X.objRelDREConfig
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146640)
        
    End Select
        
    Exit Function

End Function
