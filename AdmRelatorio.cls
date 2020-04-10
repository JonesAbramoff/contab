VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Mantém informações sobre um relatório do sistema"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodRel As String
Private mvarsTelaAuxiliar As String
Private mvarsNomeTsk As String
Private mvarsAutor As String
Private mvarsDescricao As String
Private mvarsUltOpcaoUtilizada As String
Private mvariTipo As Integer
Private mvariOrigem As Integer
Private mvariDispositivoDeSaida As Integer
Private mvarsNomeArqReimp As String
Private mvariFilialEmpresa As Integer
Private mvariLandscape As Integer
Private mvariPrevia As Integer
Private mvariOutOfProcess As Integer
Private mvariCriarArqReimp As Integer

Public objRelOpcoes As New AdmRelOpcoes
Private mvarsOpcao As String
Private mvarbConfiguraImpressora As Boolean

Public Property Let bConfiguraImpressora(ByVal vData As Boolean)
    mvarbConfiguraImpressora = vData
End Property

Public Property Get bConfiguraImpressora() As Boolean
    bConfiguraImpressora = mvarbConfiguraImpressora
End Property

Public Property Let sOpcao(ByVal vData As String)
    mvarsOpcao = vData
End Property

Public Property Get sOpcao() As String
    sOpcao = mvarsOpcao
End Property

Public Function ExecutarDiretoEmail(sCodRelParam As String, sSelecao As String, iPrevia As Integer, Optional sNovoNomeTsk As String, Optional sParam1 As String, Optional sValParam1 As String, Optional sParam2 As String, Optional sValParam2 As String, Optional sParam3 As String, Optional sValParam3 As String, Optional sParam4 As String, Optional sValParam4 As String, Optional sParam5 As String, Optional sValParam5 As String, Optional sParam6 As String, Optional sValParam6 As String, Optional sParam7 As String, Optional sValParam7 As String, Optional sParam8 As String, Optional sValParam8 As String) As Long

Dim lErro As Long, X As New AdmChamaTela
Dim sBuffer As String, objAux As AdmRelatorio
Dim sOpcao As String '??? lixo
Dim sToEmail As String, sSubject As String

On Error GoTo Erro_ExecutarDiretoEmail
    
    sCodRel = sCodRelParam
    Set objAux = Me
    
    lErro = CF("Relatorio_Le", objAux)
    If lErro Then Error 7437
    
    objRelOpcoes.sCodRel = sCodRel
        
    If IsMissing(sNovoNomeTsk) = False And sNovoNomeTsk <> "" Then
        sNomeTsk = sNovoNomeTsk
    End If
    
    If sParam1 <> "" Then
        
        lErro = objRelOpcoes.IncluirParametro(sParam1, sValParam1)
        If lErro <> AD_BOOL_TRUE Then Error 7426

        If sParam2 <> "" Then
            
            lErro = objRelOpcoes.IncluirParametro(sParam2, sValParam2)
            If lErro <> AD_BOOL_TRUE Then Error 7427
            
            If sParam3 <> "" Then
                
                lErro = objRelOpcoes.IncluirParametro(sParam3, sValParam3)
                If lErro <> AD_BOOL_TRUE Then Error 7432
                
                If sParam4 <> "" Then
                    
                    lErro = objRelOpcoes.IncluirParametro(sParam4, sValParam4)
                    If lErro <> AD_BOOL_TRUE Then Error 7433
                    
                    If sParam5 <> "" Then
                        
                        lErro = objRelOpcoes.IncluirParametro(sParam5, sValParam5)
                        If lErro <> AD_BOOL_TRUE Then Error 7434
                        
                        If sParam6 <> "" Then
                            
                            lErro = objRelOpcoes.IncluirParametro(sParam6, sValParam6)
                            If lErro <> AD_BOOL_TRUE Then Error 7435
                            
                            If sParam7 <> "" Then
                                
                                lErro = objRelOpcoes.IncluirParametro(sParam7, sValParam7)
                                If lErro <> AD_BOOL_TRUE Then Error 7435
                                
                                If sParam8 <> "" Then
                                    
                                    lErro = objRelOpcoes.IncluirParametro(sParam8, sValParam8)
                                    If lErro <> AD_BOOL_TRUE Then Error 7435
                                    
                                End If
                            
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End If
    
    End If
    
    If sSelecao <> "" Then

        objRelOpcoes.sSelecao = sSelecao

    End If
    
    lErro = Sist_ImpressoraDlg(0, iLandscape)
    If lErro <> AD_BOOL_TRUE Then Error 59310
    
    lErro = X.Chama_Tela("EnviarEmail", objRelOpcoes, sOpcao)
    If lErro <> SUCESSO Then Error 7428
    If objRelOpcoes.bDesistiu Then Error 7428
    
    lErro = ExecutarAux(iPrevia)
    If lErro Then Error 7428
    
    ExecutarDiretoEmail = SUCESSO

    Exit Function
    
Erro_ExecutarDiretoEmail:
    
    ExecutarDiretoEmail = Err
    
    Select Case Err
        Case 7426 To 7428, 7432 To 7435, 7437, 59310
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142490)
        
    End Select
    
    Exit Function
    
End Function


Private Function ExecutarAux(iPrevia As Integer) As Long

Dim lErro As Long
Dim sBuffer As String, sNomeArqReimp As String
Dim sNomeTskAux As String

On Error GoTo Erro_ExecutarAux
    
    lErro = objRelOpcoes.IncluirParametro("TEMPRESA", CStr(gsNomeEmpresa))
    If lErro <> AD_BOOL_TRUE Then Error 7359

    lErro = CF("Relatorio_Customiza_ExecutarAux", objRelOpcoes, mvarsNomeTsk)
    If lErro <> SUCESSO Then Error 20865

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIAL", CStr(gsNomeFilialEmpresa))
    If lErro <> AD_BOOL_TRUE Then Error 7361

    lErro = objRelOpcoes.IncluirParametro("TUSUARIO", CStr(gsUsuario))
    If lErro <> AD_BOOL_TRUE Then Error 7361

    lErro = objRelOpcoes.ConvParametros
    If lErro <> SUCESSO Then Error 7071
    
    'se o nome do tsk nao contem o path completo
    If InStr(sNomeTsk, "\") = 0 Then
    
        'buscar diretorio configurado
        sBuffer = String(128, 0)
        Call GetPrivateProfileString("Forprint", "DirTsks", "c:\forpw40\", sBuffer, 128, "ADM100.INI")
        
        sBuffer = StringZ(sBuffer)
        If right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
        sNomeArqReimp = sBuffer & sNomeTsk & ".rei"
        sNomeTskAux = sBuffer & sNomeTsk & ".tsk"

    Else
        
        If UCase(right(sNomeTsk, 4)) <> ".TSK" Then
            
            sNomeTskAux = sNomeTsk & ".tsk"
            
        Else
        
            sNomeTskAux = sNomeTsk
            
        End If
        
        sNomeArqReimp = left(sNomeTskAux, Len(sNomeTskAux) - 4) & ".rei"
    
    End If
    
    If (iCriarArqReimp) Then
    
        lErro = CF("Relatorio_AtualizarArqReimp", mvarsCodRel, sNomeArqReimp)
        If lErro <> SUCESSO Then Error 19420
        
    End If
    
    'Disparar a execucao
    lErro = Sistema_ExecutarRel(sNomeTskAux, objRelOpcoes.lRelOpAux, iPrevia, iOutOfProcess)
    If lErro <> AD_BOOL_TRUE Then Error 7439
    
    ExecutarAux = SUCESSO

    Exit Function
    
Erro_ExecutarAux:
    
    ExecutarAux = Err
    
    Select Case Err
        Case 7071, 7359 To 7361, 7439, 19420, 20865
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142491)
        
    End Select
    
    Exit Function
    
End Function

Public Function ExecutarDireto(sCodRelParam As String, sSelecao As String, iPrevia As Integer, Optional sNovoNomeTsk As String, Optional sParam1 As String, Optional sValParam1 As String, Optional sParam2 As String, Optional sValParam2 As String, Optional sParam3 As String, Optional sValParam3 As String, Optional sParam4 As String, Optional sValParam4 As String, Optional sParam5 As String, Optional sValParam5 As String, Optional sParam6 As String, Optional sValParam6 As String, Optional sParam7 As String, Optional sValParam7 As String, Optional sParam8 As String, Optional sValParam8 As String, Optional sParam9 As String, Optional sValParam9 As String, Optional sParam10 As String, Optional sValParam10 As String, Optional sParam11 As String, Optional sValParam11 As String, Optional sParam12 As String, Optional sValParam12 As String, Optional sParam13 As String, Optional sValParam13 As String, Optional sParam14 As String, Optional sValParam14 As String) As Long

Dim lErro As Long, X As New AdmChamaTela
Dim sBuffer As String, objAux As AdmRelatorio
Dim sOpcao As String '??? lixo
Dim sToEmail As String, sSubject As String
Dim iExecutaPrevia As Integer

On Error GoTo Erro_ExecutarDireto
    
    iExecutaPrevia = 0
    sCodRel = sCodRelParam
    Set objAux = Me
    
    lErro = CF("Relatorio_Le", objAux)
    If lErro Then Error 7437
    
    objRelOpcoes.sCodRel = sCodRel
        
    If IsMissing(sNovoNomeTsk) = False And sNovoNomeTsk <> "" Then
        sNomeTsk = sNovoNomeTsk
    End If
    
    If sParam1 <> "" Then
        
        lErro = objRelOpcoes.IncluirParametro(sParam1, sValParam1)
        If lErro <> AD_BOOL_TRUE Then Error 7426

        If sParam2 <> "" Then
            
            lErro = objRelOpcoes.IncluirParametro(sParam2, sValParam2)
            If lErro <> AD_BOOL_TRUE Then Error 7427
            
            If sParam3 <> "" Then
                
                lErro = objRelOpcoes.IncluirParametro(sParam3, sValParam3)
                If lErro <> AD_BOOL_TRUE Then Error 7432
                
                If sParam4 <> "" Then
                    
                    lErro = objRelOpcoes.IncluirParametro(sParam4, sValParam4)
                    If lErro <> AD_BOOL_TRUE Then Error 7433
                    
                    If sParam5 <> "" Then
                        
                        lErro = objRelOpcoes.IncluirParametro(sParam5, sValParam5)
                        If lErro <> AD_BOOL_TRUE Then Error 7434
                        
                        If sParam6 <> "" Then
                            
                            lErro = objRelOpcoes.IncluirParametro(sParam6, sValParam6)
                            If lErro <> AD_BOOL_TRUE Then Error 7435
                            
                            If sParam7 <> "" Then
                                
                                lErro = objRelOpcoes.IncluirParametro(sParam7, sValParam7)
                                If lErro <> AD_BOOL_TRUE Then Error 7435
                                
                                If sParam8 <> "" Then
                                    
                                    lErro = objRelOpcoes.IncluirParametro(sParam8, sValParam8)
                                    If lErro <> AD_BOOL_TRUE Then Error 7435
                                    
                                     If sParam9 <> "" Then
                                        
                                        lErro = objRelOpcoes.IncluirParametro(sParam9, sValParam9)
                                        If lErro <> AD_BOOL_TRUE Then Error 7435
                                        
                                         If sParam10 <> "" Then
                                            
                                            lErro = objRelOpcoes.IncluirParametro(sParam10, sValParam10)
                                            If lErro <> AD_BOOL_TRUE Then Error 7435
                                            
                                             If sParam11 <> "" Then
                                                
                                                lErro = objRelOpcoes.IncluirParametro(sParam11, sValParam11)
                                                If lErro <> AD_BOOL_TRUE Then Error 7435
                                                
                                                If sParam12 <> "" Then
                                                    
                                                    lErro = objRelOpcoes.IncluirParametro(sParam12, sValParam12)
                                                    If lErro <> AD_BOOL_TRUE Then Error 7435
                                                    
                                                    If sParam13 <> "" Then
                                                    
                                                        lErro = objRelOpcoes.IncluirParametro(sParam13, sValParam13)
                                                        If lErro <> AD_BOOL_TRUE Then Error 7435
                                                    
                                                        If sParam14 <> "" Then
                                                        
                                                            lErro = objRelOpcoes.IncluirParametro(sParam14, sValParam14)
                                                            If lErro <> AD_BOOL_TRUE Then Error 7435
                                                        
                                                        End If
                                                        
                                                    End If
                                                   
                                                End If
                                               
                                            End If
                                            
                                        End If
                                        
                                    End If
                                   
                                End If
                                
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End If
    
    End If
    
    If sSelecao <> "" Then

        objRelOpcoes.sSelecao = sSelecao

    End If
    
    If (iDispositivoDeSaida = REL_SAIDA_IMPRESSORA Or iDispositivoDeSaida = REL_SAIDA_EMAIL) Then
    
        If bConfiguraImpressora Then
            lErro = Sist_ImpressoraDlg(0, iLandscape)
            If lErro <> AD_BOOL_TRUE Then Error 59310
        End If
        
        If (iPrevia = 0) Then
        
            lErro = X.Chama_Tela("RelDlgExec", objRelOpcoes)
            If lErro <> SUCESSO Then Error 7428
            If objRelOpcoes.bDesistiu Then Error 7428
            
            Select Case objRelOpcoes.iDispositivoDeSaida
            
                Case REL_SAIDA_EMAIL
                    lErro = objRelOpcoes.ObterParametro("TTO_EMAIL", sToEmail)
                    If lErro <> SUCESSO Then Error 7428
                    
                    lErro = objRelOpcoes.ObterParametro("TSUBJECT", sSubject)
                    If lErro <> SUCESSO Then Error 7428
                    
                    If (Len(Trim(sToEmail)) = 0 Or Len(Trim(sSubject)) = 0) Then
                        
                        lErro = X.Chama_Tela("EnviarEmail", objRelOpcoes, sOpcao)
                        If lErro <> SUCESSO Then Error 7428
                        If objRelOpcoes.bDesistiu Then Error 7428
                
                    End If
                    
                Case REL_SAIDA_PREVIA
                    iExecutaPrevia = 1
                
            End Select
        
        Else
            
            If (iPrevia = 1) Then iExecutaPrevia = 1
            
        End If
           
    End If
    
    lErro = ExecutarAux(iExecutaPrevia)
    If lErro Then Error 7428
    
    ExecutarDireto = SUCESSO

    Exit Function
    
Erro_ExecutarDireto:
    
    ExecutarDireto = Err
    
    Select Case Err
        Case 7426 To 7428, 7432 To 7435, 7437, 59310
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142492)
        
    End Select
    
    Exit Function
    
End Function

Public Function Executar_Prossegue() As Long

Dim lErro As Long, X As New AdmChamaTela
Dim iExecutaPrevia As Integer
Dim sToEmail As String, sSubject As String
Dim sOpcao As String '??? lixo

On Error GoTo Erro_Executar_Prossegue
    
    iExecutaPrevia = 1
    If (iDispositivoDeSaida = REL_SAIDA_IMPRESSORA Or iDispositivoDeSaida = REL_SAIDA_EMAIL) Then
    
        iExecutaPrevia = 0
    
        lErro = Sist_ImpressoraDlg(0, iLandscape)
        If lErro <> AD_BOOL_TRUE Then Error 7072
        
        If (iPrevia = 0) Then
        
'            lErro = MsgBox("Exibe prévia de impressão do relatório ?", vbYesNoCancel, mvarsCodRel)
'            If (lErro = vbYes) Then
'                iExecutaPrevia = 1
'            Else
'                If lErro = vbCancel Then Error 7429
'            End If
            
            lErro = X.Chama_Tela("RelDlgExec", objRelOpcoes)
            If lErro <> SUCESSO Then Error 7429
            If objRelOpcoes.bDesistiu Then Error 7429
            
            Select Case objRelOpcoes.iDispositivoDeSaida
            
                Case REL_SAIDA_EMAIL
                    lErro = objRelOpcoes.ObterParametro("TTO_EMAIL", sToEmail)
                    If lErro <> SUCESSO Then Error 7429
                    
                    lErro = objRelOpcoes.ObterParametro("TSUBJECT", sSubject)
                    If lErro <> SUCESSO Then Error 7429
                    
                    If (Len(Trim(sToEmail)) = 0 Or Len(Trim(sSubject)) = 0) Then
                    
                        lErro = X.Chama_Tela("EnviarEmail", objRelOpcoes, sOpcao)
                        If lErro <> SUCESSO Then Error 7429
                        If objRelOpcoes.bDesistiu Then Error 7429
            
                    End If
                    
                Case REL_SAIDA_PREVIA
                    iExecutaPrevia = 1
                
            End Select
            
        Else
            
            If (iPrevia = 1) Then iExecutaPrevia = 1
            
        End If
        
    End If
        
    lErro = ExecutarAux(iExecutaPrevia)
    If lErro Then Error 7430
        
    Executar_Prossegue = SUCESSO

    Exit Function
    
Erro_Executar_Prossegue:
    
    Executar_Prossegue = Err
    
    Select Case Err
        Case 7072, 7429, 7430
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142493)
        
    End Select
    
    Exit Function
    
End Function

Public Function Executar(Optional vParam As Variant) As Long

Dim lErro As Long
Dim X As New AdmChamaTela

On Error GoTo Erro_Executar
    
    If (sTelaAuxiliar <> "") Then
    
        If Len(Trim(sOpcao)) > 0 Then
            
            If IsMissing(vParam) Then
            
                lErro = X.Chama_Tela_Relatorio(sTelaAuxiliar, Me, objRelOpcoes)
            Else
                lErro = X.Chama_Tela_Relatorio(sTelaAuxiliar, Me, objRelOpcoes, vParam)
            End If
            If lErro <> SUCESSO Then Error 7059
            
        Else
    
            If IsMissing(vParam) Then
                lErro = X.Chama_Tela(sTelaAuxiliar, Me, objRelOpcoes)
            Else
                lErro = X.Chama_Tela(sTelaAuxiliar, Me, objRelOpcoes, vParam)
            End If
            If lErro <> SUCESSO Then Error 7059
        
        End If
        
    Else
    
        lErro = Executar_Prossegue()
        If lErro <> SUCESSO Then Error 7417
    
    End If
    
    Executar = SUCESSO

    Exit Function
    
Erro_Executar:
    
    Select Case Err
    
        Case 7417, 7059
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142494)
        
    End Select
    
    Exit Function
    
End Function

Public Property Let iOutOfProcess(ByVal vData As Integer)
    mvariOutOfProcess = vData
End Property

Public Property Get iOutOfProcess() As Integer
    iOutOfProcess = mvariOutOfProcess
End Property

Public Property Let iCriarArqReimp(ByVal vData As Integer)
    mvariCriarArqReimp = vData
End Property

Public Property Get iCriarArqReimp() As Integer
    iCriarArqReimp = mvariCriarArqReimp
End Property

Public Property Let iPrevia(ByVal vData As Integer)
    mvariPrevia = vData
End Property

Public Property Get iPrevia() As Integer
    iPrevia = mvariPrevia
End Property

Public Property Let iLandscape(ByVal vData As Integer)
    mvariLandscape = vData
End Property

Public Property Get iLandscape() As Integer
    iLandscape = mvariLandscape
End Property

Public Property Let iDispositivoDeSaida(ByVal vData As Integer)
    mvariDispositivoDeSaida = vData
End Property

Public Property Get iDispositivoDeSaida() As Integer
    iDispositivoDeSaida = mvariDispositivoDeSaida
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sUltOpcaoUtilizada(ByVal vData As String)
    mvarsUltOpcaoUtilizada = vData
End Property

Public Property Get sUltOpcaoUtilizada() As String
    sUltOpcaoUtilizada = mvarsUltOpcaoUtilizada
End Property

Public Property Let sNomeTsk(ByVal vData As String)
    mvarsNomeTsk = vData
End Property

Public Property Get sNomeTsk() As String
    sNomeTsk = mvarsNomeTsk
End Property

Public Property Let sNomeArqReimp(ByVal vData As String)
    mvarsNomeArqReimp = vData
End Property

Public Property Get sNomeArqReimp() As String
    sNomeArqReimp = mvarsNomeArqReimp
End Property

Public Property Let sTelaAuxiliar(ByVal vData As String)
    mvarsTelaAuxiliar = vData
End Property

Public Property Get sTelaAuxiliar() As String
    sTelaAuxiliar = mvarsTelaAuxiliar
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sCodRel(ByVal vData As String)
    mvarsCodRel = vData
End Property

Public Property Get sCodRel() As String
    sCodRel = mvarsCodRel
End Property

Public Property Let sAutor(ByVal vData As String)
    mvarsAutor = vData
End Property

Public Property Get sAutor() As String
    sAutor = mvarsAutor
End Property

Public Sub Rel_Menu_Executar(sCodRelParam As String, Optional vParam As Variant)

'Dim objRelatorio As New AdmRelatorio
Dim lErro As Long, objAux As AdmRelatorio

On Error GoTo Erro_Rel_Menu_Executar

    sCodRel = sCodRelParam
    Set objAux = Me
    
    lErro = CF("Relatorio_Le", objAux)
    If lErro Then Error 7438
    
    objRelOpcoes.sCodRel = sCodRelParam
    
    If IsMissing(vParam) Then
        lErro = Executar()
    Else
        lErro = Executar(vParam)
    End If
    
    Exit Sub
    
Erro_Rel_Menu_Executar:

    Select Case Err
        
        Case 7438
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142495)
        
    End Select
        
    Exit Sub

End Sub

Public Function Executar_Prossegue2(objTelaParam As Object) As Long
'fecha a tela de parametros antes de prosseguir (para evitar que a tela de previa do relatorio fique por baixo da tela do corporator)

Dim lErro As Long

On Error GoTo Erro_Executar_Prossegue2
    
    Call objTelaParam.Unload(objTelaParam)
        
    lErro = Executar_Prossegue
    If lErro <> SUCESSO Then Error 59467
    
    Executar_Prossegue2 = SUCESSO

    Exit Function
    
Erro_Executar_Prossegue2:
    
    Executar_Prossegue2 = Err
    
    Select Case Err
    
        Case 59467
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142496)
        
    End Select
    
    Exit Function
    
End Function

Public Function Executar_Prossegue3() As Long

Dim lErro As Long
Dim X As New AdmChamaTela
Dim sOpcao As String

On Error GoTo Erro_Executar_Prossegue3

    Call X.Chama_Tela("EnviarEmail", objRelOpcoes, sOpcao)
    
    If objRelOpcoes.bDesistiu Then gError 97120
    
    lErro = ExecutarAux(0)
    If lErro <> SUCESSO Then gError 97109
    
    Executar_Prossegue3 = SUCESSO
    
    Exit Function
    
Erro_Executar_Prossegue3:
    
    Executar_Prossegue3 = gErr
    
    Select Case gErr
        
        Case 97109
        
        Case 97120
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142497)
    
    End Select
    
    Exit Function
    
End Function

Public Function ExecutarDeNovo(Optional sParam1 As String, Optional sValParam1 As String, Optional sParam2 As String, Optional sValParam2 As String, Optional sParam3 As String, Optional sValParam3 As String, Optional sParam4 As String, Optional sValParam4 As String, Optional sParam5 As String, Optional sValParam5 As String, Optional sParam6 As String, Optional sValParam6 As String) As Long

Dim lErro As Long
Dim sBuffer As String, sNomeArqReimp As String
Dim sNomeTskAux As String

On Error GoTo Erro_ExecutarDeNovo
    
    'se o nome do tsk nao contem o path completo
    If InStr(sNomeTsk, "\") = 0 Then
    
        'buscar diretorio configurado
        sBuffer = String(128, 0)
        Call GetPrivateProfileString("Forprint", "DirTsks", "c:\forpw40\", sBuffer, 128, "ADM100.INI")
        
        sBuffer = StringZ(sBuffer)
        If right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
        sNomeArqReimp = sBuffer & sNomeTsk & ".rei"
        sNomeTskAux = sBuffer & sNomeTsk & ".tsk"

    Else
        
        If UCase(right(sNomeTsk, 4)) <> ".TSK" Then
            
            sNomeTskAux = sNomeTsk & ".tsk"
            
        Else
        
            sNomeTskAux = sNomeTsk
            
        End If
        
        sNomeArqReimp = left(sNomeTskAux, Len(sNomeTskAux) - 4) & ".rei"
    
    End If
    
    
'''''============== alterado por Shirley =============
'''''=
''''    Call objRelOpcoes.Excluir
''''
''''    If sParam1 <> "" Then
''''
''''        lErro = objRelOpcoes.IncluirParametro(sParam1, sValParam1)
''''        If lErro <> AD_BOOL_TRUE Then Error 7426
''''
''''        If sParam2 <> "" Then
''''
''''            lErro = objRelOpcoes.IncluirParametro(sParam2, sValParam2)
''''            If lErro <> AD_BOOL_TRUE Then Error 7427
''''
''''            If sParam3 <> "" Then
''''
''''                lErro = objRelOpcoes.IncluirParametro(sParam3, sValParam3)
''''                If lErro <> AD_BOOL_TRUE Then Error 7432
''''
''''                If sParam4 <> "" Then
''''
''''                    lErro = objRelOpcoes.IncluirParametro(sParam4, sValParam4)
''''                    If lErro <> AD_BOOL_TRUE Then Error 7433
''''
''''                    If sParam5 <> "" Then
''''
''''                        lErro = objRelOpcoes.IncluirParametro(sParam5, sValParam5)
''''                        If lErro <> AD_BOOL_TRUE Then Error 7434
''''
''''                        If sParam6 <> "" Then
''''
''''                            lErro = objRelOpcoes.IncluirParametro(sParam6, sValParam6)
''''                            If lErro <> AD_BOOL_TRUE Then Error 7435
''''
''''                        End If
''''
''''                    End If
''''
''''                End If
''''
''''            End If
''''
''''        End If
''''
''''    End If
''''
''''    lErro = X.Relatorio_AtualizarArqReimp(mvarsCodRel, sNomeArqReimp)
''''    If lErro <> SUCESSO Then Error 19420
''''
'''''=============== fim shirley
''''
''''
    
    
    'Disparar a execucao
    lErro = Sistema_ExecutarRel(sNomeTskAux, objRelOpcoes.lRelOpAux, iPrevia, iOutOfProcess)
    If lErro <> AD_BOOL_TRUE Then Error 7439
    
    ExecutarDeNovo = SUCESSO

    Exit Function
    
Erro_ExecutarDeNovo:
    
    ExecutarDeNovo = Err
    
    Select Case Err
        Case 7071, 7359 To 7361, 7439, 19420
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142498)
        
    End Select
    
    Exit Function
    
End Function


Private Sub Class_Initialize()
    bConfiguraImpressora = True
End Sub
