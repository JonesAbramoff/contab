VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaDan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function Dan_OS_Grava(ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Dan_OS_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184787

    lErro = CF("Dan_OS_Grava_EmTrans", objDan_OS)
    If lErro <> SUCESSO Then gError 184788

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184789

    Dan_OS_Grava = SUCESSO

    Exit Function

Erro_Dan_OS_Grava:

    Dan_OS_Grava = gErr

    Select Case gErr

        Case 184787
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 184788

        Case 184789
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184790)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Dan_OS_Grava_EmTrans(ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_Dan_OS_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184791
    Next

    'Le a tabelaDan_OS
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM Dan_OS WHERE OS= ? ", _
                0, iAux, objDan_OS.sOS)
    If lErro <> AD_SQL_SUCESSO Then gError 184792

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184793

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 184794

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Dan_OS SET Cliente= ? , Modelo= ? , NumSerie= ? ", alComando(0), _
                objDan_OS.lCliente, objDan_OS.sModelo, objDan_OS.sNumSerie)
        If lErro <> AD_SQL_SUCESSO Then gError 184795
        
        lErro = Dan_ItensOS_Exclui(alComando(2), alComando(3), objDan_OS)
        If lErro <> SUCESSO Then gError 198246

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO Dan_OS( OS, Cliente, Modelo, NumSerie) VALUES (?,?,?,?)", _
                objDan_OS.sOS, objDan_OS.lCliente, objDan_OS.sModelo, objDan_OS.sNumSerie)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184796

    End If

    lErro = Dan_ItensOS_Insere(alComando(4), objDan_OS)
    If lErro <> SUCESSO Then gError 198247

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Dan_OS_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Dan_OS_Grava_EmTrans:

    Dan_OS_Grava_EmTrans = gErr

    Select Case gErr

        Case 184791
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184792, 184793
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DAN_OS", gErr)

        Case 184794
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DAN_OS", gErr)

        Case 184795
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_DAN_OS", gErr)

        Case 184796
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DAN_OS", gErr)
            
        Case 198246, 198247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184797)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function Dan_OS_Exclui(ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Dan_OS_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184776

    lErro = CF("Dan_OS_Exclui_EmTrans", objDan_OS)
    If lErro <> SUCESSO Then gError 184777

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184778

    Dan_OS_Exclui = SUCESSO

    Exit Function

Erro_Dan_OS_Exclui:

    Dan_OS_Exclui = gErr

    Select Case gErr

        Case 184776
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 184777

        Case 184778
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184779)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Dan_OS_Exclui_EmTrans(ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_Dan_OS_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184780
    Next

    'Le a tabelaDan_OS
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM Dan_OS WHERE OS= ? ", _
                0, iAux, objDan_OS.sOS)
    If lErro <> AD_SQL_SUCESSO Then gError 184781

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184782

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 184783

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 184784

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Dan_OS", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 184785
    
    lErro = Dan_ItensOS_Exclui(alComando(2), alComando(3), objDan_OS)
    If lErro <> SUCESSO Then gError 198245

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Dan_OS_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Dan_OS_Exclui_EmTrans:

    Dan_OS_Exclui_EmTrans = gErr

    Select Case gErr

        Case 184780
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184781, 184782
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DAN_OS", gErr)

        Case 184783
            Call Rotina_Erro(vbOKOnly, "ERRO_DAN_OS_NAO_CADASTRADO", gErr, objDan_OS.sOS)

        Case 184784
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DAN_OS", gErr)

        Case 184785
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DAN_OS", gErr)
            
        Case 198245

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184786)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Dan_ItensOS_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim iAux As Integer

On Error GoTo Erro_Dan_ItensOS_Exclui

    'Le a tabelaDan_ItensOS
    lErro = Comando_ExecutarPos(lComando1, "SELECT 1 FROM Dan_ItensOS WHERE OS= ? ", _
                0, iAux, objDan_OS.sOS)
    If lErro <> AD_SQL_SUCESSO Then gError 198251

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198252

    'Se não existir => ERRO
    Do While lErro = AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 198254
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Dan_ItensOS", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 198255

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198253

    Loop

    Dan_ItensOS_Exclui = SUCESSO

    Exit Function

Erro_Dan_ItensOS_Exclui:

    Dan_ItensOS_Exclui = gErr

    Select Case gErr

        Case 198251, 198252
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DAN_ITENSOS", gErr)

        Case 198253
            Call Rotina_Erro(vbOKOnly, "ERRO_DAN_ITENSOS_NAO_CADASTRADO", gErr, objDan_OS.sOS)

        Case 198254
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DAN_ITENSOS", gErr)

        Case 198255
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DAN_ITENSOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198256)

    End Select

    Exit Function

End Function

Private Function Dan_ItensOS_Insere(ByVal lComando As Long, ByVal objDan_OS As ClassDan_OS) As Long

Dim lErro As Long
Dim objDan_ItensOS As ClassDan_ItensOS

On Error GoTo Erro_Dan_ItensOS_Insere

    For Each objDan_ItensOS In objDan_OS.colItens

        objDan_ItensOS.sOS = objDan_OS.sOS

        lErro = Comando_Executar(lComando, "INSERT INTO Dan_ItensOS( OS, Item, Produto, Quantidade) VALUES (?,?,?,?)", _
                objDan_ItensOS.sOS, objDan_ItensOS.iItem, objDan_ItensOS.sProduto, objDan_ItensOS.dQuantidade)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198266
        
    Next

    Dan_ItensOS_Insere = SUCESSO

    Exit Function

Erro_Dan_ItensOS_Insere:

    Dan_ItensOS_Insere = gErr

    Select Case gErr
        
        Case 198266
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DAN_ITENSOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198267)

    End Select

    Exit Function

End Function

