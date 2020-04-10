VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function Projetos_Grava(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Projetos_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181309

    lErro = CF("Projetos_Grava_EmTrans", objProjetos)
    If lErro <> SUCESSO Then gError 181310

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181311
    
    Projetos_Grava = SUCESSO

    Exit Function

Erro_Projetos_Grava:

    Projetos_Grava = gErr

    Select Case gErr

        Case 181309
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181310

        Case 181311
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181411)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Projetos_Grava_EmTrans(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim lErro1 As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocEscopo As Long
Dim sCodigo As String
Dim iFilial As Integer
Dim objRegraWFW As ClassRegraWFW
Dim objUsu As ClassUsuarios
Dim sRespAnt As String, sConteudo As String

On Error GoTo Erro_Projetos_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181313
    Next
    
    sCodigo = String(STRING_PRJ_CODIGO, 0)
    
    'Le a tabelaProjetos
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo, FilialEmpresa FROM Projetos WHERE NomeReduzido = ? AND (Codigo<> ? OR (FilialEmpresa <> ? AND FilialEmpresa <> ? AND ? <> ?)) ", _
                0, sCodigo, iFilial, objProjetos.sNomeReduzido, objProjetos.sCodigo, objProjetos.iFilialEmpresa, EMPRESA_TODA, objProjetos.iFilialEmpresa, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 181314

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181315
    
    'Nome reduzido repetido
    If lErro = SUCESSO Then gError 181810
    
    sRespAnt = String(STRING_PRJ_RESPONSAVEL, 0)

    'Le a tabelaProjetos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocEscopo, Responsavel FROM Projetos WHERE Codigo= ? AND (FilialEmpresa = ? OR FilialEmpresa = ? OR ?=?) ", _
                0, lNumIntDoc, lNumIntDocEscopo, sRespAnt, objProjetos.sCodigo, objProjetos.iFilialEmpresa, EMPRESA_TODA, objProjetos.iFilialEmpresa, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 181314

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181315
    
    objProjetos.lNumIntDocEscopo = lNumIntDocEscopo
    objProjetos.objEscopo.lNumIntDoc = lNumIntDocEscopo
    objProjetos.lNumIntDoc = lNumIntDoc
    
    lErro1 = PRJEscopo_Grava_EmTrans(objProjetos.objEscopo)
    If lErro1 <> SUCESSO Then gError 181464
    
    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then
    
        objProjetos.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 181316

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Projetos SET NomeReduzido= ? , Descricao= ? , " & _
                "DataCriacao= ? , Cliente= ? , FilialCliente= ? , Responsavel= ? , Objetivo= ? , " & _
                "Justificativa= ? , Observacao= ? , DataInicio= ? , DataFim= ? , DataInicioReal= ? , " & _
                "DataFimReal= ? , PercentualComplet= ? , NumIntDocEscopo= ?, FilialEmpresa = ?, Segmento = ? ", alComando(0), _
                objProjetos.sNomeReduzido, objProjetos.sDescricao, _
                objProjetos.dtDataCriacao, objProjetos.lCliente, objProjetos.iFilialCliente, objProjetos.sResponsavel, objProjetos.sObjetivo, _
                objProjetos.sJustificativa, objProjetos.sObservacao, objProjetos.dtDataInicio, objProjetos.dtDataFim, objProjetos.dtDataInicioReal, _
                objProjetos.dtDataFimReal, objProjetos.dPercentualComplet, objProjetos.lNumIntDocEscopo, objProjetos.iFilialEmpresa, objProjetos.sSegmento)
        If lErro <> AD_SQL_SUCESSO Then gError 181317

    Else
    
        sRespAnt = ""

        objProjetos.lNumIntDocEscopo = objProjetos.objEscopo.lNumIntDoc

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PROJETOS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 181312

        objProjetos.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO Projetos( NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, " & _
                "Descricao, DataCriacao, Cliente, FilialCliente, Responsavel, " & _
                "Objetivo, Justificativa, Observacao, DataInicio, DataFim, " & _
                "DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo,Segmento) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objProjetos.lNumIntDoc, objProjetos.sCodigo, objProjetos.iFilialEmpresa, objProjetos.sNomeReduzido, _
                objProjetos.sDescricao, objProjetos.dtDataCriacao, objProjetos.lCliente, objProjetos.iFilialCliente, objProjetos.sResponsavel, _
                objProjetos.sObjetivo, objProjetos.sJustificativa, objProjetos.sObservacao, objProjetos.dtDataInicio, objProjetos.dtDataFim, _
                objProjetos.dtDataInicioReal, objProjetos.dtDataFimReal, objProjetos.dPercentualComplet, objProjetos.lNumIntDocEscopo, objProjetos.sSegmento)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181318

    End If
    
    lErro = TiposCamposCust_Grava_EmTrans(objProjetos.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 181569
    
    objProjetos.objCamposCust.lNumIntDocOrigem = objProjetos.lNumIntDoc
    objProjetos.objCamposCust.iTipoNumIntDocOrigem = objProjetos.objTiposCamposCust.iCodigo
    
    lErro = CamposCustomizados_Grava_EmTrans(objProjetos.objCamposCust)
    If lErro <> SUCESSO Then gError 181583
    
    lErro = CF("Projetos_Grava_Customizado", objProjetos)
    If lErro <> SUCESSO Then gError 181583
    
    If Not (objProjetos.objAnexos Is Nothing) Then
        objProjetos.objAnexos.iTipoDoc = ANEXO_TIPO_PRJ
        objProjetos.objAnexos.lNumIntDoc = objProjetos.lNumIntDoc
        lErro = CF("Anexos_Grava", objProjetos.objAnexos)
        If lErro <> SUCESSO Then gError 181583
    End If
    
    lErro = CF("Config_Le", "FATConfig", "PRJ_AVISA_TROCA_RESP", EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 181583
    
    If StrParaInt(sConteudo) = MARCADO Then
    
        Set objRegraWFW = New ClassRegraWFW
        
        If Len(Trim(sRespAnt)) > 0 Then
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sRespAnt
            lErro = CF("Usuarios_Le", objUsu)
            If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
            If lErro = SUCESSO Then objRegraWFW.colUsuarios.Add objUsu.sCodUsuario
        End If
        
        If Len(Trim(objProjetos.sResponsavel)) > 0 Then
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = objProjetos.sResponsavel
            lErro = CF("Usuarios_Le", objUsu)
            If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
            If lErro = SUCESSO Then objRegraWFW.colUsuarios.Add objUsu.sCodUsuario
        End If
        
        If objRegraWFW.colUsuarios.Count <> 0 Then
        
            objRegraWFW.sAvisoMsgRet = "O projeto " & objProjetos.sCodigo & " foi alterado em " & Format(Date, "dd/mm/yyyy") & " às " & Format(Now, "HH:MM:SS") & "hs pelo usuário " & gsUsuario & "."
            If sRespAnt <> objProjetos.sResponsavel Then
            
                If sRespAnt = "" Then
                    objRegraWFW.sAvisoMsgRet = objRegraWFW.sAvisoMsgRet & " Inclusão de responsabilidade para: " & objProjetos.sResponsavel
                ElseIf objProjetos.sResponsavel = "" Then
                    objRegraWFW.sAvisoMsgRet = objRegraWFW.sAvisoMsgRet & " Retirada de responsabilidade de: " & sRespAnt
                Else
                    objRegraWFW.sAvisoMsgRet = objRegraWFW.sAvisoMsgRet & " Transferência de responsabilidade de: " & sRespAnt & " para: " & objProjetos.sResponsavel
                End If
    
            End If
        
                'gravar o aviso
            lErro = CF("AvisoWFW_Grava_EmTrans", objRegraWFW.colUsuarios, objRegraWFW.sAvisoMsgRet, TRANSACAOWFW_PROJETO)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Projetos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Projetos_Grava_EmTrans:

    Projetos_Grava_EmTrans = gErr

    Select Case gErr

        Case 181312, 181464, 181569

        Case 181313
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181314, 181315
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 181316
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PROJETOS", gErr)

        Case 181317
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PROJETOS", gErr)

        Case 181318
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PROJETOS", gErr)
            
        Case 181810
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NOMEREDUZIDO_REPETIDO", gErr, objProjetos.sNomeReduzido, sCodigo, iFilial)

        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181410)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function Projetos_Exclui(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Projetos_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181300

    lErro = CF("Projetos_Exclui_EmTrans", objProjetos)
    If lErro <> SUCESSO Then gError 181301

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181302

    Projetos_Exclui = SUCESSO

    Exit Function

Erro_Projetos_Exclui:

    Projetos_Exclui = gErr

    Select Case gErr

        Case 181300
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181301

        Case 181302
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181409)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Projetos_Exclui_EmTrans(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDocEscopo As Long
Dim lNumIntDoc As Long
Dim objCamposCust As New ClassCamposCustomizados

On Error GoTo Erro_Projetos_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181303
    Next

    'Le a tabelaProjetos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocEscopo FROM Projetos WHERE Codigo= ? AND (FilialEmpresa = ? OR FilialEmpresa = ?) ", _
                0, lNumIntDoc, lNumIntDocEscopo, objProjetos.sCodigo, objProjetos.iFilialEmpresa, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 181304

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181305

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 181306

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 181307

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Projetos", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 181308
    
    objProjetos.objEscopo.lNumIntDoc = lNumIntDocEscopo
    
    lErro = PRJEscopo_Exclui_EmTrans(objProjetos.objEscopo)
    If lErro <> SUCESSO Then gError 181463
    
    objProjetos.lNumIntDoc = lNumIntDoc
    lErro = PRJEtapas_Exclui_Projeto(objProjetos)
    If lErro <> SUCESSO Then gError 182925
    
    objCamposCust.iTipoNumIntDocOrigem = CAMPO_CUSTOMIZADO_TIPO_PROJETO
    objCamposCust.lNumIntDocOrigem = lNumIntDoc
    
    lErro = CamposCustomizados_Exclui_EmTrans(objCamposCust)
    If lErro <> SUCESSO Then gError 181526
    
    '################################################################
    'Inserido por Wagner 26/05/2006
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_PROJETO, lNumIntDoc)
    If lErro <> SUCESSO Then gError 182050
    '################################################################
    
    If Not (objProjetos.objAnexos Is Nothing) Then
        objProjetos.objAnexos.iTipoDoc = ANEXO_TIPO_PRJ
        objProjetos.objAnexos.lNumIntDoc = lNumIntDoc
        lErro = CF("Anexos_Exclui", objProjetos.objAnexos)
        If lErro <> SUCESSO Then gError 182050
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Projetos_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Projetos_Exclui_EmTrans:

    Projetos_Exclui_EmTrans = gErr

    Select Case gErr

        Case 181303
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181304, 181305
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 181306
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjetos.sCodigo, objProjetos.iFilialEmpresa)

        Case 181307
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PROJETOS", gErr)

        Case 181308
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PROJETOS", gErr)
            
        Case 181463, 181526, 182050, 182925

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181408)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJEscopo_Grava(ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJEscopo_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181448

    lErro = CF("PRJEscopo_Grava_EmTrans", objPRJEscopo)
    If lErro <> SUCESSO Then gError 181449

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181450

    PRJEscopo_Grava = SUCESSO

    Exit Function

Erro_PRJEscopo_Grava:

    PRJEscopo_Grava = gErr

    Select Case gErr

        Case 181448
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181449

        Case 181450
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181451)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJEscopo_Grava_EmTrans(ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJEscopo_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181437
    Next

    'Le a tabelaPRJEscopo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM PRJEscopo WHERE NumIntDoc= ? ", _
                0, iAux, objPRJEscopo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181438

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181439

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 181440

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJEscopo SET Descricao= ? , Expectativa= ? , FatoresSucesso= ? , Restricoes= ? , " & _
                "Premissas= ? , ExclusoesEspecificas= ? ", alComando(0), _
                objPRJEscopo.sDescricao, objPRJEscopo.sExpectativa, objPRJEscopo.sFatoresSucesso, objPRJEscopo.sRestricoes, _
                objPRJEscopo.sPremissas, objPRJEscopo.sExclusoesEspecificas)
        If lErro <> AD_SQL_SUCESSO Then gError 181441

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJESCOPO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 181442

        objPRJEscopo.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEscopo( NumIntDoc, Descricao, Expectativa, FatoresSucesso, " & _
                "Restricoes, Premissas, ExclusoesEspecificas) VALUES (?,?,?,?,?,?,?)", _
                objPRJEscopo.lNumIntDoc, objPRJEscopo.sDescricao, objPRJEscopo.sExpectativa, objPRJEscopo.sFatoresSucesso, _
                objPRJEscopo.sRestricoes, objPRJEscopo.sPremissas, objPRJEscopo.sExclusoesEspecificas)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181443

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEscopo_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEscopo_Grava_EmTrans:

    PRJEscopo_Grava_EmTrans = gErr

    Select Case gErr

        Case 181442

        Case 181437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181438, 181439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJESCOPO", gErr)

        Case 181440
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJESCOPO", gErr)

        Case 181441
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJESCOPO", gErr)

        Case 181443
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJESCOPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181444)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJEscopo_Exclui(ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJEscopo_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181444

    lErro = CF("PRJEscopo_Exclui_EmTrans", objPRJEscopo)
    If lErro <> SUCESSO Then gError 181445

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181446

    PRJEscopo_Exclui = SUCESSO

    Exit Function

Erro_PRJEscopo_Exclui:

    PRJEscopo_Exclui = gErr

    Select Case gErr

        Case 181444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181445

        Case 181446
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181447)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJEscopo_Exclui_EmTrans(ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_PRJEscopo_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181452
    Next

    'Le a tabelaPRJEscopo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM PRJEscopo WHERE NumIntDoc= ? ", _
                0, iAux, objPRJEscopo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181453

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181454

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 181455

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 181456

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEscopo", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 181457

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEscopo_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEscopo_Exclui_EmTrans:

    PRJEscopo_Exclui_EmTrans = gErr

    Select Case gErr

        Case 181452
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181453, 181454
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJESCOPO", gErr)

        Case 181455
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJESCOPO_NAO_CADASTRADO", gErr, objPRJEscopo.lNumIntDoc)

        Case 181456
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJESCOPO", gErr)

        Case 181457
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJESCOPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181458)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function CamposCustomizados_Grava(ByVal objCamposCustomizados As ClassCamposCustomizados) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CamposCustomizados_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181477

    lErro = CF("CamposCustomizados_Grava_EmTrans", objCamposCustomizados)
    If lErro <> SUCESSO Then gError 181478

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181479

    CamposCustomizados_Grava = SUCESSO

    Exit Function

Erro_CamposCustomizados_Grava:

    CamposCustomizados_Grava = gErr

    Select Case gErr

        Case 181477
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181478

        Case 181479
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181523)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CamposCustomizados_Grava_EmTrans(ByVal objCamposCustomizados As ClassCamposCustomizados) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CamposCustomizados_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181481
    Next

    'Le a tabelaCamposCustomizados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM CamposCustomizados WHERE NumIntDocOrigem= ?  AND TipoNumIntDocOrigem= ? ", _
                0, iAux, objCamposCustomizados.lNumIntDocOrigem, objCamposCustomizados.iTipoNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 181482

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181483

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 181484

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE CamposCustomizados SET Data1= ? , Data2= ? , Data3= ? , Data4= ? , " & _
                "Data5= ? , Texto1= ? , Texto2= ? , Texto3= ? , Texto4= ? , " & _
                "Texto5= ? , Numero1= ? , Numero2= ? , Numero3= ? , Numero4= ? , " & _
                "Numero5= ? , Valor1= ? , Valor2= ? , Valor3= ? , Valor4= ? , " & _
                "Valor5= ? ", alComando(0), _
                objCamposCustomizados.dtData(1), objCamposCustomizados.dtData(2), objCamposCustomizados.dtData(3), objCamposCustomizados.dtData(4), _
                objCamposCustomizados.dtData(5), objCamposCustomizados.sTexto(1), objCamposCustomizados.sTexto(2), objCamposCustomizados.sTexto(3), objCamposCustomizados.sTexto(4), _
                objCamposCustomizados.sTexto(5), objCamposCustomizados.lNumero(1), objCamposCustomizados.lNumero(2), objCamposCustomizados.lNumero(3), objCamposCustomizados.lNumero(4), _
                objCamposCustomizados.lNumero(5), objCamposCustomizados.dValor(1), objCamposCustomizados.dValor(2), objCamposCustomizados.dValor(3), objCamposCustomizados.dValor(4), _
                objCamposCustomizados.dValor(5))
        If lErro <> AD_SQL_SUCESSO Then gError 181485

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_CAMPOSCUSTOMIZADOS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 181480

        objCamposCustomizados.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO CamposCustomizados( NumIntDoc, NumIntDocOrigem, TipoNumIntDocOrigem, Data1, " & _
                "Data2, Data3, Data4, Data5, Texto1, " & _
                "Texto2, Texto3, Texto4, Texto5, Numero1, " & _
                "Numero2, Numero3, Numero4, Numero5, Valor1, " & _
                "Valor2, Valor3, Valor4, Valor5) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objCamposCustomizados.lNumIntDoc, objCamposCustomizados.lNumIntDocOrigem, objCamposCustomizados.iTipoNumIntDocOrigem, objCamposCustomizados.dtData(1), _
                objCamposCustomizados.dtData(2), objCamposCustomizados.dtData(3), objCamposCustomizados.dtData(4), objCamposCustomizados.dtData(5), objCamposCustomizados.sTexto(1), _
                objCamposCustomizados.sTexto(2), objCamposCustomizados.sTexto(3), objCamposCustomizados.sTexto(4), objCamposCustomizados.sTexto(5), objCamposCustomizados.lNumero(1), _
                objCamposCustomizados.lNumero(2), objCamposCustomizados.lNumero(3), objCamposCustomizados.lNumero(4), objCamposCustomizados.lNumero(5), objCamposCustomizados.dValor(1), _
                objCamposCustomizados.dValor(2), objCamposCustomizados.dValor(3), objCamposCustomizados.dValor(4), objCamposCustomizados.dValor(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181486

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CamposCustomizados_Grava_EmTrans = SUCESSO

    Exit Function

Erro_CamposCustomizados_Grava_EmTrans:

    CamposCustomizados_Grava_EmTrans = gErr

    Select Case gErr

        Case 181480

        Case 181481
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181482, 181483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSCUSTOMIZADOS", gErr)

        Case 181484
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CAMPOSCUSTOMIZADOS", gErr)

        Case 181485
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CAMPOSCUSTOMIZADOS", gErr)

        Case 181486
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CAMPOSCUSTOMIZADOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181522)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function CamposCustomizados_Exclui(ByVal objCamposCustomizados As ClassCamposCustomizados) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CamposCustomizados_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181468

    lErro = CF("CamposCustomizados_Exclui_EmTrans", objCamposCustomizados)
    If lErro <> SUCESSO Then gError 181469

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181470

    CamposCustomizados_Exclui = SUCESSO

    Exit Function

Erro_CamposCustomizados_Exclui:

    CamposCustomizados_Exclui = gErr

    Select Case gErr

        Case 181468
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181469

        Case 181470
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181524)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CamposCustomizados_Exclui_EmTrans(ByVal objCamposCustomizados As ClassCamposCustomizados) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CamposCustomizados_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181471
    Next

    'Le a tabelaCamposCustomizados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM CamposCustomizados WHERE NumIntDocOrigem= ?  AND TipoNumIntDocOrigem= ? ", _
                0, iAux, objCamposCustomizados.lNumIntDocOrigem, objCamposCustomizados.iTipoNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 181472

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181473

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 181474

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 181475

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CamposCustomizados", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 181476

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CamposCustomizados_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_CamposCustomizados_Exclui_EmTrans:

    CamposCustomizados_Exclui_EmTrans = gErr

    Select Case gErr

        Case 181471
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181472, 181473
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSCUSTOMIZADOS", gErr)

        Case 181474
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOSCUSTOMIZADOS_NAO_CADASTRADO", gErr, objCamposCustomizados.lNumIntDocOrigem, objCamposCustomizados.iTipoNumIntDocOrigem)

        Case 181475
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CAMPOSCUSTOMIZADOS", gErr)

        Case 181476
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CAMPOSCUSTOMIZADOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181525)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function TiposCamposCust_Grava(ByVal objTiposCamposCust As ClassTiposCamposCust) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TiposCamposCust_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181550

    lErro = CF("TiposCamposCust_Grava_EmTrans", objTiposCamposCust)
    If lErro <> SUCESSO Then gError 181551

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181552

    TiposCamposCust_Grava = SUCESSO

    Exit Function

Erro_TiposCamposCust_Grava:

    TiposCamposCust_Grava = gErr

    Select Case gErr

        Case 181550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181551

        Case 181552
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181561)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TiposCamposCust_Grava_EmTrans(ByVal objTiposCamposCust As ClassTiposCamposCust) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim tTiposCamposCust As typeTiposCamposCust
Dim lNumIntDoc As Long

On Error GoTo Erro_TiposCamposCust_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181554
    Next

    'Le a tabelaTiposCamposCust
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo, DatasPreenchida, TextosPreenchidos, NumerosPreenchidos, ValoresPreenchidos FROM TiposCamposCust WHERE NomeTela= ? ", _
                0, tTiposCamposCust.iCodigo, tTiposCamposCust.iDatasPreenchida, tTiposCamposCust.iTextosPreenchidos, tTiposCamposCust.iNumerosPreenchidos, tTiposCamposCust.iValoresPreenchidos, objTiposCamposCust.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 181555

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181556

    'Se existir => UPDATE, senão => INSERT
    If lErro <> AD_SQL_SUCESSO Then gError 181568
    
    objTiposCamposCust.iCodigo = tTiposCamposCust.iCodigo
    
'    If objTiposCamposCust.iDatasPreenchida < tTiposCamposCust.iDatasPreenchida Then
'        objTiposCamposCust.iDatasPreenchida = tTiposCamposCust.iDatasPreenchida
'    End If
'    If objTiposCamposCust.iNumerosPreenchidos < tTiposCamposCust.iNumerosPreenchidos Then
'        objTiposCamposCust.iNumerosPreenchidos = tTiposCamposCust.iNumerosPreenchidos
'    End If
'    If objTiposCamposCust.iTextosPreenchidos < tTiposCamposCust.iTextosPreenchidos Then
'        objTiposCamposCust.iTextosPreenchidos = tTiposCamposCust.iTextosPreenchidos
'    End If
'    If objTiposCamposCust.iValoresPreenchidos < tTiposCamposCust.iValoresPreenchidos Then
'        objTiposCamposCust.iValoresPreenchidos = tTiposCamposCust.iValoresPreenchidos
'    End If

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 181557

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposCamposCust SET DatasPreenchida= ? , TextosPreenchidos= ? , " & _
            "NumerosPreenchidos= ? , ValoresPreenchidos= ? ", alComando(0), _
            objTiposCamposCust.iDatasPreenchida, objTiposCamposCust.iTextosPreenchidos, _
            objTiposCamposCust.iNumerosPreenchidos, objTiposCamposCust.iValoresPreenchidos)
    If lErro <> AD_SQL_SUCESSO Then gError 181558

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TiposCamposCust_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TiposCamposCust_Grava_EmTrans:

    TiposCamposCust_Grava_EmTrans = gErr

    Select Case gErr
        Case 181554
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181555, 181556
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSCAMPOSCUST", gErr)

        Case 181557
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TIPOSCAMPOSCUST", gErr)

        Case 181558
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TIPOSCAMPOSCUST", gErr)
            
        Case 181568
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSCAMPOSCUST_NAO_CADASTRADO2", gErr, objTiposCamposCust.sNomeTela)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181560)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function ProjetoInfo_Grava(ByVal objProjetoInfo As ClassProjetoInfo, ByVal iTipoOrigem As Integer, ByVal lNumIntDocOrigem As Long, Optional ByVal sCodOP As String = "", Optional ByVal iFililEmpresa As Integer = 0) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_ProjetoInfo_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181776
    Next
    
    lErro = ProjetoInfo_Exclui(iTipoOrigem, lNumIntDocOrigem, sCodOP, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 181777
    
    If Not (objProjetoInfo Is Nothing) Then

        objProjetoInfo.iTipoOrigem = iTipoOrigem
        objProjetoInfo.lNumIntDocOrigem = lNumIntDocOrigem
        objProjetoInfo.sCodigoOP = sCodOP
        If iFilialEmpresa <> 0 Then
            objProjetoInfo.iFilialEmpresa = iFilialEmpresa
        Else
            objProjetoInfo.iFilialEmpresa = giFilialEmpresa
        End If

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PROJETOINFO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 181778
        
        objProjetoInfo.lNumIntDoc = lNumIntDoc
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO ProjetoInfo( NumIntDoc, NumIntDocPRJ, NumIntDocEtapa, TipoOrigem, NumIntDocOrigem, CodigoOP, FilialEmpresa) VALUES (?,?,?,?,?,?,?)", _
                objProjetoInfo.lNumIntDoc, objProjetoInfo.lNumIntDocPRJ, objProjetoInfo.lNumIntDocEtapa, objProjetoInfo.iTipoOrigem, objProjetoInfo.lNumIntDocOrigem, objProjetoInfo.sCodigoOP, objProjetoInfo.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181779
        
        lErro = PRJCustoReceitaReal_Grava(objProjetoInfo)
        If lErro <> SUCESSO Then gError 181780
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProjetoInfo_Grava = SUCESSO

    Exit Function

Erro_ProjetoInfo_Grava:

    ProjetoInfo_Grava = gErr

    Select Case gErr

        Case 181776
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181777, 181778, 181780

        Case 181779
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PROJETOINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181781)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJCustoReceitaReal_Grava(ByVal objProjetoInfo As ClassProjetoInfo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_PRJCustoReceitaReal_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181658
    Next
    
    For Each objPRJCR In objProjetoInfo.colPRJCustoReceitaReal

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJCUSTORECEITAREAL", lNumIntDoc)
        If lErro <> SUCESSO Then gError 181657

        objPRJCR.lNumIntDoc = lNumIntDoc
        objPRJCR.lNumIntDocPRJInfo = objProjetoInfo.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJCustoReceitaReal( NumIntDoc, NumIntDocPRJ, NumIntDocEtapa, NumIntDocPRJInfo, " & _
                " Valor, Quantidade, Percentual, TipoValor,Observacao, Item, CalcAuto) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJCR.lNumIntDoc, objPRJCR.lNumIntDocPRJ, objPRJCR.lNumIntDocEtapa, objPRJCR.lNumIntDocPRJInfo, _
                objPRJCR.dValor, objPRJCR.dQuantidade, objPRJCR.dPercent, _
                objPRJCR.iTipoValor, objPRJCR.sObservacao, objPRJCR.sItem, objPRJCR.iCalcAuto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181663
        
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJCustoReceitaReal_Grava = SUCESSO

    Exit Function

Erro_PRJCustoReceitaReal_Grava:

    PRJCustoReceitaReal_Grava = gErr

    Select Case gErr

        Case 181657

        Case 181658
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181663
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJCUSTORECEITAREAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181664)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ProjetoInfo_Exclui(ByVal iTipoOrigem As Integer, ByVal lNumIntDocOrigem As Long, Optional ByVal sCodOP As String = "", Optional ByVal iFililEmpresa As Integer = 0) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDocPRJInfo As Long
Dim lNumIntDocPRJCR As Long

On Error GoTo Erro_ProjetoInfo_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181648
    Next

    'Le a tabela ProjetoInfo
    If iTipoOrigem <> PRJ_CR_TIPO_OP Then
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM ProjetoInfo WHERE TipoOrigem= ?  AND NumIntDocOrigem= ? ", _
                    0, lNumIntDocPRJInfo, iTipoOrigem, lNumIntDocOrigem)
    Else
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM ProjetoInfo WHERE TipoOrigem= ?  AND CodigoOP= ? AND FilialEmpresa = ? ", _
                    0, lNumIntDocPRJInfo, iTipoOrigem, sCodOP, giFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 181770

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181771

    'Se  existir
    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> SUCESSO Then gError 181772
    
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ProjetoInfo", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 181773

        'Le a tabela PRJCustoReceitaReal
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NUmIntDoc FROM PRJCustoReceitaReal WHERE NumIntDocPRJInfo= ? ", _
                    0, lNumIntDocPRJCR, lNumIntDocPRJInfo)
        If lErro <> AD_SQL_SUCESSO Then gError 181649
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181650
    
        'Se não existir => ERRO
        Do While lErro <> AD_SQL_SEM_DADOS
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 181652
        
            lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJCustoReceitaReal", alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 181653
        
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181651
        
        Loop
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProjetoInfo_Exclui = SUCESSO

    Exit Function

Erro_ProjetoInfo_Exclui:

    ProjetoInfo_Exclui = gErr

    Select Case gErr

        Case 181648
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181649, 181650, 181651
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCUSTORECEITAREAL", gErr)

        Case 181652
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJCUSTORECEITAREAL", gErr)

        Case 181653
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJCUSTORECEITAREAL", gErr)

        Case 181770, 181771
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOINFO", gErr)

        Case 181772
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PROJETOINFO", gErr)

        Case 181773
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PROJETOINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181667)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJEtapas_Grava(ByVal objPRJEtapas As ClassPRJEtapas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJEtapas_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181721

    lErro = CF("PRJEtapas_Grava_EmTrans", objPRJEtapas)
    If lErro <> SUCESSO Then gError 181722

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181723
    
    PRJEtapas_Grava = SUCESSO

    Exit Function

Erro_PRJEtapas_Grava:

    PRJEtapas_Grava = gErr

    Select Case gErr

        Case 181721
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181722

        Case 181723
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181732)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJEtapas_Grava_EmTrans(ByVal objPRJEtapas As ClassPRJEtapas) As Long

Dim lErro As Long
Dim lErro1 As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocEscopo As Long
Dim objProjeto As ClassProjetos
Dim objRegraWFW As ClassRegraWFW
Dim objUsu As ClassUsuarios
Dim sRespAnt As String, sConteudo As String
Dim lNumIntDocPRJ As Long

On Error GoTo Erro_PRJEtapas_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181725
    Next
    
    sRespAnt = String(STRING_PRJ_RESPONSAVEL, 0)

    'Le a tabelaPRJEtapas
    If objPRJEtapas.lNumIntDoc = 0 Then
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocEscopo, NumIntDocPRJ, Responsavel FROM PRJEtapas WHERE NumIntDocPRJ= ? AND Codigo = ? ", _
                    0, lNumIntDoc, lNumIntDocEscopo, lNumIntDocPRJ, sRespAnt, objPRJEtapas.lNumIntDocPRJ, objPRJEtapas.sCodigo)
    Else
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocEscopo, NumIntDocPRJ, Responsavel FROM PRJEtapas WHERE NumIntDoc= ? ", _
                    0, lNumIntDoc, lNumIntDocEscopo, lNumIntDocPRJ, sRespAnt, objPRJEtapas.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 181726

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181727

    objPRJEtapas.lNumIntDocEscopo = lNumIntDocEscopo
    objPRJEtapas.objEscopo.lNumIntDoc = lNumIntDocEscopo
    
    lErro1 = PRJEscopo_Grava_EmTrans(objPRJEtapas.objEscopo)
    If lErro1 <> SUCESSO Then gError 185519

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then
    
        objPRJEtapas.lNumIntDocPRJ = lNumIntDocPRJ

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 181728

        If objPRJEtapas.lNumIntDoc = 0 Then
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJEtapas SET Referencia= ? , NomeReduzido= ? , Descricao= ? , " & _
                    "Cliente= ? , FilialCliente= ? , Responsavel= ? , Objetivo= ? , Justificativa= ? , " & _
                    "Observacao= ? , DataInicio= ? , DataFim= ? , NumIntDocEtapaPaiOrg= ? , DataInicioReal= ? , " & _
                    "DataFimReal= ? , PercentualComplet= ?, NumIntDocEscopo= ?, Nivel = ?, Seq = ?, Posicao = ? ", alComando(0), _
                    objPRJEtapas.sReferencia, objPRJEtapas.sNomeReduzido, objPRJEtapas.sDescricao, _
                    objPRJEtapas.lCliente, objPRJEtapas.iFilialCliente, objPRJEtapas.sResponsavel, objPRJEtapas.sObjetivo, objPRJEtapas.sJustificativa, _
                    objPRJEtapas.sObservacao, objPRJEtapas.dtDataInicio, objPRJEtapas.dtDataFim, objPRJEtapas.lNumIntDocEtapaPaiOrg, objPRJEtapas.dtDataInicioReal, _
                    objPRJEtapas.dtDataFimReal, objPRJEtapas.dPercentualComplet, objPRJEtapas.lNumIntDocEscopo, objPRJEtapas.iNivel, objPRJEtapas.iSeq, objPRJEtapas.iPosicao)
        Else
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJEtapas SET Referencia= ? , NomeReduzido= ? , Descricao= ? , " & _
                    "Cliente= ? , FilialCliente= ? , Responsavel= ? , Objetivo= ? , Justificativa= ? , " & _
                    "Observacao= ? , DataInicio= ? , DataFim= ? , NumIntDocEtapaPaiOrg= ? , DataInicioReal= ? , " & _
                    "DataFimReal= ? , PercentualComplet= ?, NumIntDocEscopo= ?, Nivel = ?, Seq = ?, Posicao = ?, Codigo = ? ", alComando(0), _
                    objPRJEtapas.sReferencia, objPRJEtapas.sNomeReduzido, objPRJEtapas.sDescricao, _
                    objPRJEtapas.lCliente, objPRJEtapas.iFilialCliente, objPRJEtapas.sResponsavel, objPRJEtapas.sObjetivo, objPRJEtapas.sJustificativa, _
                    objPRJEtapas.sObservacao, objPRJEtapas.dtDataInicio, objPRJEtapas.dtDataFim, objPRJEtapas.lNumIntDocEtapaPaiOrg, objPRJEtapas.dtDataInicioReal, _
                    objPRJEtapas.dtDataFimReal, objPRJEtapas.dPercentualComplet, objPRJEtapas.lNumIntDocEscopo, objPRJEtapas.iNivel, objPRJEtapas.iSeq, objPRJEtapas.iPosicao, objPRJEtapas.sCodigo)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 181729
        
        objPRJEtapas.lNumIntDoc = lNumIntDoc

    Else
    
        sRespAnt = ""

        objPRJEtapas.lNumIntDocEscopo = objPRJEtapas.objEscopo.lNumIntDoc

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 181724

        objPRJEtapas.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapas( NumIntDoc, NumIntDocPRJ, Codigo, Referencia, " & _
                "NomeReduzido, Descricao, Cliente, FilialCliente, Responsavel, " & _
                "Objetivo, Justificativa, Observacao, DataInicio, DataFim, " & _
                "NumIntDocEtapaPaiOrg, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Nivel, Seq, Posicao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJEtapas.lNumIntDoc, objPRJEtapas.lNumIntDocPRJ, objPRJEtapas.sCodigo, objPRJEtapas.sReferencia, _
                objPRJEtapas.sNomeReduzido, objPRJEtapas.sDescricao, objPRJEtapas.lCliente, objPRJEtapas.iFilialCliente, objPRJEtapas.sResponsavel, _
                objPRJEtapas.sObjetivo, objPRJEtapas.sJustificativa, objPRJEtapas.sObservacao, objPRJEtapas.dtDataInicio, objPRJEtapas.dtDataFim, _
                objPRJEtapas.lNumIntDocEtapaPaiOrg, objPRJEtapas.dtDataInicioReal, objPRJEtapas.dtDataFimReal, objPRJEtapas.dPercentualComplet, objPRJEtapas.lNumIntDocEscopo, objPRJEtapas.iNivel, objPRJEtapas.iSeq, objPRJEtapas.iPosicao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181730

    End If
    
    lErro = PRJEtapaItensProd_Grava_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185347
    
    lErro = PRJEtapaMaquinas_Grava_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185348
    
    lErro = PRJEtapaMateriais_Grava_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185349
    
    lErro = PRJEtapaMO_Grava_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185350
    
    lErro = PRJEtapasPredecessoras_Grava_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185351
    
    lErro = TiposCamposCust_Grava_EmTrans(objPRJEtapas.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 185520
    
    objPRJEtapas.objCamposCust.lNumIntDocOrigem = objPRJEtapas.lNumIntDoc
    objPRJEtapas.objCamposCust.iTipoNumIntDocOrigem = objPRJEtapas.objTiposCamposCust.iCodigo
    
    lErro = CamposCustomizados_Grava_EmTrans(objPRJEtapas.objCamposCust)
    If lErro <> SUCESSO Then gError 185521
    
    lErro = CF("Config_Le", "FATConfig", "PRJ_AVISA_TROCA_RESP", EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 185521
    
    If StrParaInt(sConteudo) = MARCADO Then
    
        Set objRegraWFW = New ClassRegraWFW
        Set objProjeto = New ClassProjetos
        
        objProjeto.lNumIntDoc = objPRJEtapas.lNumIntDocPRJ
        
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If Len(Trim(sRespAnt)) > 0 Then
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sRespAnt
            lErro = CF("Usuarios_Le", objUsu)
            If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
            If lErro = SUCESSO Then objRegraWFW.colUsuarios.Add objUsu.sCodUsuario
        End If
        
        If Len(Trim(objPRJEtapas.sResponsavel)) > 0 Then
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = objPRJEtapas.sResponsavel
            lErro = CF("Usuarios_Le", objUsu)
            If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
            If lErro = SUCESSO Then objRegraWFW.colUsuarios.Add objUsu.sCodUsuario
        End If
        
        If objRegraWFW.colUsuarios.Count <> 0 Then
        
            objRegraWFW.sAvisoMsgRet = "A etapa " & objPRJEtapas.sCodigo & " do projeto " & objProjeto.sCodigo & " foi alterada em " & Format(Date, "dd/mm/yyyy") & " às " & Format(Now, "HH:MM:SS") & "hs pelo usuário " & gsUsuario & "."
            If sRespAnt <> objPRJEtapas.sResponsavel Then
            
                If sRespAnt = "" Then
                    objRegraWFW.sAvisoMsgRet = objRegraWFW.sAvisoMsgRet & " Inclusão de responsabilidade para: " & objPRJEtapas.sResponsavel
                ElseIf objPRJEtapas.sResponsavel = "" Then
                    objRegraWFW.sAvisoMsgRet = objRegraWFW.sAvisoMsgRet & " Retirada de responsabilidade de: " & sRespAnt
                Else
                    objRegraWFW.sAvisoMsgRet = objRegraWFW.sAvisoMsgRet & " Transferência de responsabilidade de: " & sRespAnt & " para: " & objPRJEtapas.sResponsavel
                End If
    
            End If
        
                'gravar o aviso
            lErro = CF("AvisoWFW_Grava_EmTrans", objRegraWFW.colUsuarios, objRegraWFW.sAvisoMsgRet, TRANSACAOWFW_PRJETAPAS)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapas_Grava_EmTrans:

    PRJEtapas_Grava_EmTrans = gErr

    Select Case gErr

        Case 181724, 185347 To 185351, 185519, 185520, 185521

        Case 181725
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181726, 181727
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)

        Case 181728
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAS", gErr)

        Case 181729
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJETAPAS", gErr)

        Case 181730
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPAS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181731)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJEtapas_Exclui(ByVal objPRJEtapas As ClassPRJEtapas) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim colEtapasFilhas As New Collection
Dim objPRJEtapasFilhas As ClassPRJEtapas
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_PRJEtapas_Exclui

    'Le as etapas descendentes (filhas, netas, etc ...)
    lErro = CF("PRJEtapas_Le_Descendentes", objPRJEtapas, colEtapasFilhas)
    If lErro <> SUCESSO Then gError 185533
    
    'Se existirem descendentes confirma a exclusão
    If colEtapasFilhas.Count <> 0 Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ETAPA_COM_DESCENDENTES", objPRJEtapas.sCodigo)
        If vbResult = vbNo Then gError 185534
    End If

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181712

    'Exclui a etapa em questão
    lErro = CF("PRJEtapas_Exclui_EmTrans", objPRJEtapas)
    If lErro <> SUCESSO Then gError 181713
    
    'Exclui suas descendentes
    For Each objPRJEtapasFilhas In colEtapasFilhas
        lErro = CF("PRJEtapas_Exclui_EmTrans", objPRJEtapasFilhas)
        If lErro <> SUCESSO Then gError 185535
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 181714

    PRJEtapas_Exclui = SUCESSO

    Exit Function

Erro_PRJEtapas_Exclui:

    PRJEtapas_Exclui = gErr

    Select Case gErr

        Case 181712
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 181713, 185533, 185534, 185535

        Case 181714
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181734)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJEtapas_Exclui_EmTrans(ByVal objPRJEtapas As ClassPRJEtapas, Optional ByVal bNaoValidaPre As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocEscopo As Long
Dim lNumIntDoc As Long
Dim sCodigoEtapa As String
Dim objCamposCust As New ClassCamposCustomizados

On Error GoTo Erro_PRJEtapas_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181715
    Next

    'Le a tabelaPRJEtapas\ Etapas filhas para exclusão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocEscopo FROM PRJEtapas WHERE Codigo = ? AND NumIntDocPRJ = ? ", _
                0, lNumIntDoc, lNumIntDocEscopo, objPRJEtapas.sCodigo, objPRJEtapas.lNumIntDocPRJ)
    If lErro <> AD_SQL_SUCESSO Then gError 181716

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181717

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 181718

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 181719
    
    If Not bNaoValidaPre Then
    
        sCodigoEtapa = String(STRING_ETAPAPRJ_CODIGO, 0)
    
        'Verifica se é predecessora de outra etapa
        lErro = Comando_Executar(alComando(2), "SELECT E.Codigo FROM PRJEtapas AS E, PRJEtapasPredecessoras AS P WHERE E.NumIntDoc = P.NumIntDocEtapa AND P.NumIntDocEtapaPre = ? ", _
                    sCodigoEtapa, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 189023
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189024
    
        'Se não existir => ERRO
        If lErro <> AD_SQL_SEM_DADOS Then gError 189025
        
    End If

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 181720
        
    objPRJEtapas.objEscopo.lNumIntDoc = lNumIntDocEscopo
    
    lErro = PRJEscopo_Exclui_EmTrans(objPRJEtapas.objEscopo)
    If lErro <> SUCESSO Then gError 185525
    
    objCamposCust.iTipoNumIntDocOrigem = CAMPO_CUSTOMIZADO_TIPO_ETAPA
    objCamposCust.lNumIntDocOrigem = lNumIntDoc
    
    lErro = CamposCustomizados_Exclui_EmTrans(objCamposCust)
    If lErro <> SUCESSO Then gError 185526
    
    lErro = PRJEtapaItensProd_Exclui_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185527
    
    lErro = PRJEtapaMaquinas_Exclui_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185528
    
    lErro = PRJEtapaMateriais_Exclui_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185529
    
    lErro = PRJEtapaMO_Exclui_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185530
    
    lErro = PRJEtapasPredecessoras_Exclui_EmTrans(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185531
    
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_ETAPAPRJ, lNumIntDoc)
    If lErro <> SUCESSO Then gError 185532
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapas_Exclui_EmTrans:

    PRJEtapas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 181715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181716, 181717, 189023, 189024
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)
            
        Case 181718
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO2", gErr)
        
        Case 181719
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAS", gErr)

        Case 181720
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPAS", gErr)
            
        Case 185525 To 185532
        
        Case 189025
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ETAPA_PREDECESSORA", gErr, objPRJEtapas.sCodigo, sCodigoEtapa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181733)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapas_Exclui_Projeto(ByVal objProjeto As ClassProjetos) As Long

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer
Dim objEtapa As ClassPRJEtapas
Dim sCodigo As String

On Error GoTo Erro_PRJEtapas_Exclui_Projeto

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185005
    Next
    
    sCodigo = String(STRING_ETAPAPRJ_CODIGO, 0)

    'Le a tabelaPRJEtapas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Codigo FROM PRJEtapas WHERE NumIntDocPRJ = ? ", _
                0, lNumIntDoc, sCodigo, objProjeto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185006

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185007

    'Wnquanto existir
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objEtapa = New ClassPRJEtapas
        
        objEtapa.lNumIntDoc = lNumIntDoc
        objEtapa.sCodigo = sCodigo
        objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc

        'Exclui as etapas
        lErro = CF("PRJEtapas_Exclui_EmTrans", objEtapa, True)
        If lErro <> SUCESSO Then gError 185008
        
        'Busca Primeiro
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185010
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapas_Exclui_Projeto = SUCESSO

    Exit Function

Erro_PRJEtapas_Exclui_Projeto:

    PRJEtapas_Exclui_Projeto = gErr

    Select Case gErr

        Case 185005
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185006, 185007, 185010
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)

        Case 185008

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181733)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function OrganogramaPRJ_Grava(ByVal colEtapas As Collection, ByVal colEtapasExc As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim objEtapa As ClassPRJEtapas
Dim objEtapaAux As ClassPRJEtapas
Dim colEtapasAux As New Collection
Dim objProjeto As New ClassProjetos
Dim sCodigoEtapa As String
Dim bAchou As Boolean

On Error GoTo Erro_OrganogramaPRJ_Grava

    If colEtapas.Count = 0 Then gError 182933

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182927
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185005
    Next
    
    objProjeto.lNumIntDoc = colEtapas.Item(1).lNumIntDocPRJ

    For Each objEtapa In colEtapas
    
        Set objEtapaAux = New ClassPRJEtapas
    
        'Se a etapa já existe
        If Len(Trim(objEtapa.sCodigoAnt)) Then
            
            objEtapaAux.sCodigo = objEtapa.sCodigoAnt
            objEtapaAux.lNumIntDocPRJ = objEtapa.lNumIntDocPRJ
        
            'Lê
            lErro = CF("PRJEtapas_Le", objEtapaAux)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182928
            
            'Recupera as informações que estavam no BD atualizando com o que foi
            'mexido na tela de Organograma
            objEtapaAux.sCodigo = objEtapa.sCodigo
            objEtapaAux.sReferencia = objEtapa.sReferencia
            objEtapaAux.sNomeReduzido = objEtapa.sNomeReduzido
            objEtapaAux.sCodigoPaiOrg = objEtapa.sCodigoPaiOrg
            objEtapaAux.iNivel = objEtapa.iNivel
            objEtapaAux.iSeq = objEtapa.iSeq
            objEtapaAux.iPosicao = objEtapa.iPosicao
            
        Else
        
            Set objEtapaAux = objEtapa
            
        End If
        
        colEtapasAux.Add objEtapaAux

    Next

    'Exclui as etapas que foram removidas da tela
    For Each objEtapa In colEtapasExc
    
        objEtapa.sCodigo = objEtapa.sCodigoAnt
        
        'TEM QUE VERIFICAR SE A ETAPA É PREDECESSORA DE ALGUMA OUTRA QUE
        'NÃO VAI SER EXCLUIDA
        sCodigoEtapa = String(STRING_ETAPAPRJ_CODIGO, 0)
    
        'Verifica se é predecessora de outra etapa
        lErro = Comando_Executar(alComando(0), "SELECT E.Codigo FROM PRJEtapas AS E, PRJEtapasPredecessoras AS P WHERE E.NumIntDoc = P.NumIntDocEtapa AND P.NumIntDocEtapaPre = ? ", _
                    sCodigoEtapa, objEtapa.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 189026
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189027
    
        'Para cada etapa que a que vai ser excluída for predecessora
        Do While lErro <> AD_SQL_SEM_DADOS
        
            'Busca na lista de exclusões
            bAchou = False
            For Each objEtapaAux In colEtapasExc
                If objEtapaAux.sCodigoAnt = sCodigoEtapa Then
                    bAchou = True
                End If
            Next
            
            'Se não encontrar erro (A etapa a ser excluida é predecessora de outra que não está sendo excluida)
            If Not bAchou Then gError 189028
        
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189024
        
        Loop
        
        lErro = CF("PRJEtapas_Exclui_EmTrans", objEtapa, False)
        If lErro <> SUCESSO Then gError 182929
    Next
    
    'Grava as etapas no BD
    For Each objEtapa In colEtapasAux
    
        objEtapa.objTiposCamposCust.iCodigo = CAMPO_CUSTOMIZADO_TIPO_ETAPA
        objEtapa.objTiposCamposCust.sNomeTela = "EtapaPRJ"
        
        If Len(Trim(objEtapa.sDescricao)) = 0 Then objEtapa.sDescricao = objEtapa.sNomeReduzido

        lErro = CF("PRJEtapas_Grava_EmTrans", objEtapa)
        If lErro <> SUCESSO Then gError 182930

        For Each objEtapaAux In colEtapasAux
        
            If objEtapaAux.sCodigoPaiOrg = objEtapa.sCodigo Then
                objEtapaAux.lNumIntDocEtapaPaiOrg = objEtapa.lNumIntDoc
            End If
        
        Next

    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182931
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrganogramaPRJ_Grava = SUCESSO

    Exit Function

Erro_OrganogramaPRJ_Grava:

    OrganogramaPRJ_Grava = gErr

    Select Case gErr

        Case 182927
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 182928 To 182930

        Case 182931
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 182933
            Call Rotina_Erro(vbOKOnly, "ERRO_ORGANOGRAMA_SEM_ETAPAS", gErr)
        
        Case 189026, 189027
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)

        Case 189028
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ETAPA_PREDECESSORA", gErr, objEtapa.sCodigo, sCodigoEtapa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182932)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaItensProd_Grava_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPRJEtapaItensProd As ClassPRJEtapaItensProd

On Error GoTo Erro_PRJEtapaItensProd_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185344
    Next
    
    lErro = PRJEtapaItensProd_Exclui_EmTrans(objEtapa)
    If lErro <> SUCESSO Then gError 185342

    For Each objPRJEtapaItensProd In objEtapa.colItensProduzidos
    
        If objPRJEtapaItensProd.lNumIntDoc = 0 Then
            
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPAITENSPROD", lNumIntDoc)
            If lErro <> SUCESSO Then gError 185343
    
            objPRJEtapaItensProd.lNumIntDoc = lNumIntDoc
        End If
        
        objPRJEtapaItensProd.lNumIntDocEtapaPRJ = objEtapa.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapaItensProd( NumIntDoc, NumIntDocEtapaPRJ, Seq, Produto, " & _
                "Descricao, Versao, UM, Quantidade, Observacao) VALUES (?,?,?,?,?,?,?,?,?)", _
                objPRJEtapaItensProd.lNumIntDoc, objPRJEtapaItensProd.lNumIntDocEtapaPRJ, objPRJEtapaItensProd.iSeq, objPRJEtapaItensProd.sProduto, _
                objPRJEtapaItensProd.sDescricao, objPRJEtapaItensProd.sVersao, objPRJEtapaItensProd.sUM, objPRJEtapaItensProd.dQuantidade, objPRJEtapaItensProd.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185345

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaItensProd_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaItensProd_Grava_EmTrans:

    PRJEtapaItensProd_Grava_EmTrans = gErr

    Select Case gErr

        Case 185342, 185343

        Case 185344
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185345
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPAITENSPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185346)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaItensProd_Exclui_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJEtapaItensProd_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185336
    Next

    'Le a tabelaPRJEtapaItensProd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapaItensProd WHERE NumIntDocEtapaPRJ= ? ", _
                0, lNumIntDoc, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185337

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185338

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185340
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapaItensProd", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 185341
        
        'Busca próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185339
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaItensProd_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaItensProd_Exclui_EmTrans:

    PRJEtapaItensProd_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185336
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185337, 185338, 185339
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAITENSPROD", gErr)

        Case 185340
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAITENSPROD", gErr)

        Case 185341
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPAITENSPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185334)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaMaquinas_Grava_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPRJEtapaMaquinas As ClassPRJEtapaMaquinas

On Error GoTo Erro_PRJEtapaMaquinas_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185365
    Next
    
    lErro = PRJEtapaMaquinas_Exclui_EmTrans(objEtapa)
    If lErro <> SUCESSO Then gError 185366

    For Each objPRJEtapaMaquinas In objEtapa.colMaquinas

        If objPRJEtapaMaquinas.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPAMAQUINAS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 185367
    
            objPRJEtapaMaquinas.lNumIntDoc = lNumIntDoc
            
        End If
        
        objPRJEtapaMaquinas.lNumIntDocEtapa = objEtapa.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapaMaquinas( NumIntDoc, NumIntDocEtapa, Seq, NumIntDocMaq, " & _
                "Descricao, Quantidade, Horas, Custo, Tipo, Data, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJEtapaMaquinas.lNumIntDoc, objPRJEtapaMaquinas.lNumIntDocEtapa, objPRJEtapaMaquinas.iSeq, objPRJEtapaMaquinas.lNumIntDocMaq, _
                objPRJEtapaMaquinas.sDescricao, objPRJEtapaMaquinas.iQuantidade, objPRJEtapaMaquinas.dHoras, objPRJEtapaMaquinas.dCusto, objPRJEtapaMaquinas.iTipo, objPRJEtapaMaquinas.dtData, objPRJEtapaMaquinas.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185368

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaMaquinas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaMaquinas_Grava_EmTrans:

    PRJEtapaMaquinas_Grava_EmTrans = gErr

    Select Case gErr

        Case 185365
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 185366, 185367

        Case 185368
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPAMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185369)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaMaquinas_Exclui_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJEtapaMaquinas_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185359
    Next

    'Le a tabelaPRJEtapaMaquinas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapaMaquinas WHERE NumIntDocEtapa= ? ", _
                0, lNumIntDoc, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185360

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185361

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185363
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapaMaquinas", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 185364
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185362
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaMaquinas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaMaquinas_Exclui_EmTrans:

    PRJEtapaMaquinas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185359
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185360, 185361, 185362
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAMAQUINAS", gErr)

        Case 185363
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAMAQUINAS", gErr)

        Case 185364
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPAMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185333)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaMO_Grava_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPRJEtapaMO As ClassPRJEtapaMO

On Error GoTo Erro_PRJEtapaMO_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185388
    Next
    
    lErro = PRJEtapaMO_Exclui_EmTrans(objEtapa)
    If lErro <> SUCESSO Then gError 185389

    For Each objPRJEtapaMO In objEtapa.colMaoDeObra

        If objPRJEtapaMO.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPAMO", lNumIntDoc)
            If lErro <> SUCESSO Then gError 185390
    
            objPRJEtapaMO.lNumIntDoc = lNumIntDoc
            
        End If
            
        objPRJEtapaMO.lNumIntDocEtapa = objEtapa.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapaMO( NumIntDoc, NumIntDocEtapa, Seq, MaoDeObra, " & _
                "Descricao, Quantidade, Horas, Custo, Tipo, Data, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJEtapaMO.lNumIntDoc, objPRJEtapaMO.lNumIntDocEtapa, objPRJEtapaMO.iSeq, objPRJEtapaMO.iMaoDeObra, _
                objPRJEtapaMO.sDescricao, objPRJEtapaMO.iQuantidade, objPRJEtapaMO.dHoras, objPRJEtapaMO.dCusto, objPRJEtapaMO.iTipo, objPRJEtapaMO.dtData, objPRJEtapaMO.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185391

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaMO_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaMO_Grava_EmTrans:

    PRJEtapaMO_Grava_EmTrans = gErr

    Select Case gErr
        
        Case 185388
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185389, 185390

        Case 185391
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPAMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185392)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaMO_Exclui_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJEtapaMO_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185382
    Next

    'Le a tabelaPRJEtapaMO
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapaMO WHERE NumIntDocEtapa= ? ", _
                0, lNumIntDoc, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185383

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185384

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185386
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapaMO", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 185387
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185385
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaMO_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaMO_Exclui_EmTrans:

    PRJEtapaMO_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185382
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185383, 185384, 185385
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAMO", gErr)

        Case 185386
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAMO", gErr)

        Case 185387
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPAMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185335)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapasPredecessoras_Grava_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPRJEtapasPredecessoras As ClassPRJEtapasPredecessoras

On Error GoTo Erro_PRJEtapasPredecessoras_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185411
    Next
    
    lErro = PRJEtapasPredecessoras_Exclui_EmTrans(objEtapa)
    If lErro <> SUCESSO Then gError 185412

    For Each objPRJEtapasPredecessoras In objEtapa.colPredecessores

        If objPRJEtapasPredecessoras.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPASPREDECESSORAS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 185413
    
            objPRJEtapasPredecessoras.lNumIntDoc = lNumIntDoc
            
        End If
            
        objPRJEtapasPredecessoras.lNumIntDocEtapa = objEtapa.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapasPredecessoras( NumIntDoc, NumIntDocEtapa, NumIntDocEtapaPre) VALUES (?,?,?)", _
                objPRJEtapasPredecessoras.lNumIntDoc, objPRJEtapasPredecessoras.lNumIntDocEtapa, objPRJEtapasPredecessoras.lNumIntDocEtapaPre)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185414

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapasPredecessoras_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapasPredecessoras_Grava_EmTrans:

    PRJEtapasPredecessoras_Grava_EmTrans = gErr

    Select Case gErr
        
        Case 185411, 185412
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 185413

        Case 185414
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPASPREDECESSORAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185415)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Public Function PRJEtapasPredecessoras_Exclui_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJEtapasPredecessoras_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185405
    Next

    'Le a tabelaPRJEtapasPredecessoras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapasPredecessoras WHERE NumIntDocEtapa= ? ", _
                0, lNumIntDoc, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185406

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185407

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185409
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapasPredecessoras", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 185410
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185408
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapasPredecessoras_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapasPredecessoras_Exclui_EmTrans:

    PRJEtapasPredecessoras_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185405
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185406, 185407, 185408
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPASPREDECESSORAS", gErr)

        Case 185409
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPASPREDECESSORAS", gErr)

        Case 185410
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPASPREDECESSORAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185402)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaMateriais_Grava_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPRJEtapaMateriais As ClassPRJEtapaMateriais

On Error GoTo Erro_PRJEtapaMateriais_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185434
    Next
    
    lErro = PRJEtapaMateriais_Exclui_EmTrans(objEtapa)
    If lErro <> SUCESSO Then gError 185435

    For Each objPRJEtapaMateriais In objEtapa.colMateriaPrima

        If objPRJEtapaMateriais.lNumIntDoc = 0 Then
        
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPAMATERIAIS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 185436
    
            objPRJEtapaMateriais.lNumIntDoc = lNumIntDoc
        
        End If
            
        objPRJEtapaMateriais.lNumIntDocEtapa = objEtapa.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapaMateriais( NumIntDoc, NumIntDocEtapa, Seq, Produto, " & _
                "Versao, Descricao, UM, Quantidade, Custo, " & _
                "Tipo, Data, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJEtapaMateriais.lNumIntDoc, objPRJEtapaMateriais.lNumIntDocEtapa, objPRJEtapaMateriais.iSeq, objPRJEtapaMateriais.sProduto, _
                objPRJEtapaMateriais.sVersao, objPRJEtapaMateriais.sDescricao, objPRJEtapaMateriais.sUM, objPRJEtapaMateriais.dQuantidade, objPRJEtapaMateriais.dCusto, _
                objPRJEtapaMateriais.iTipo, objPRJEtapaMateriais.dtData, objPRJEtapaMateriais.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185437

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaMateriais_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaMateriais_Grava_EmTrans:

    PRJEtapaMateriais_Grava_EmTrans = gErr

    Select Case gErr

        Case 185434
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185435, 185436

        Case 185437
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPAMATERIAIS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185438)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJEtapaMateriais_Exclui_EmTrans(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJEtapaMateriais_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185428
    Next

    'Le a tabelaPRJEtapaMateriais
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapaMateriais WHERE NumIntDocEtapa= ? ", _
                0, lNumIntDoc, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185429

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185430

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185432
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapaMateriais", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 185433
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185431
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaMateriais_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaMateriais_Exclui_EmTrans:

    PRJEtapaMateriais_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185428
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185429, 185430, 185431
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAMATERIAIS", gErr)

        Case 185432
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAMATERIAIS", gErr)

        Case 185433
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPAMATERIAIS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185403)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJRecebPagto_Grava(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJRecebPagto_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185567

    lErro = CF("PRJRecebPagto_Grava_EmTrans", objPRJRecebPagto)
    If lErro <> SUCESSO Then gError 185568

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185569

    PRJRecebPagto_Grava = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Grava:

    PRJRecebPagto_Grava = gErr

    Select Case gErr

        Case 185567
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 185568

        Case 185569
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185718)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJRecebPagto_Grava_EmTrans(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumero As Long
Dim iFilialEmpresa As Integer
Dim iIncluiCFF As Integer

On Error GoTo Erro_PRJRecebPagto_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185571
    Next

    'Só pode existir um recebimento por Projeto\Proposta\Contrato
    If objPRJRecebPagto.iTipo = PRJ_TIPO_RECEB Then
        'Le a tabelaPRJRecebPagto
        lErro = Comando_Executar(alComando(2), "SELECT Numero, FilialEmpresa FROM PRJRecebPagto WHERE NumIntDocPRJ= ? AND NumIntDocProposta = ? AND NumIntDocContrato = ? AND Tipo = ? AND (Numero <> ? OR FilialEMpresa <> ?)", _
                    lNumero, iFilialEmpresa, objPRJRecebPagto.lNumIntDocPRJ, objPRJRecebPagto.lNumIntDocProposta, objPRJRecebPagto.lNumIntDocContrato, objPRJRecebPagto.iTipo, objPRJRecebPagto.lNumero, objPRJRecebPagto.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 187486
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187487
        
        If lErro = SUCESSO Then gError 187488
    End If

    'Le a tabelaPRJRecebPagto
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, IncluiCFF FROM PRJRecebPagto WHERE Numero= ? AND FilialEmpresa= ? AND Tipo = ?", _
                0, lNumIntDoc, iIncluiCFF, objPRJRecebPagto.lNumero, objPRJRecebPagto.iFilialEmpresa, objPRJRecebPagto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 185572

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185573

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        objPRJRecebPagto.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185574

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJRecebPagto SET NumIntDocPRJ= ?,  Valor= ? , CliForn= ? , " & _
                "Filial= ? , NumIntDocProposta= ? , IncluiCFF= ? , NumIntDocContrato= ? ", alComando(0), _
                objPRJRecebPagto.lNumIntDocPRJ, objPRJRecebPagto.dValor, objPRJRecebPagto.lCliForn, _
                objPRJRecebPagto.iFilial, objPRJRecebPagto.lNumIntDocProposta, objPRJRecebPagto.iIncluiCFF, objPRJRecebPagto.lNumIntDocContrato)
        If lErro <> AD_SQL_SUCESSO Then gError 185575

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJRECEBPAGTO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 185570

        objPRJRecebPagto.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJRecebPagto( NumIntDoc, NumIntDocPRJ, Tipo, Numero, " & _
                "Valor, CliForn, Filial, NumIntDocProposta, IncluiCFF, " & _
                "FilialEmpresa,NumIntDocContrato) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJRecebPagto.lNumIntDoc, objPRJRecebPagto.lNumIntDocPRJ, objPRJRecebPagto.iTipo, objPRJRecebPagto.lNumero, _
                objPRJRecebPagto.dValor, objPRJRecebPagto.lCliForn, objPRJRecebPagto.iFilial, objPRJRecebPagto.lNumIntDocProposta, objPRJRecebPagto.iIncluiCFF, _
                objPRJRecebPagto.iFilialEmpresa, objPRJRecebPagto.lNumIntDocContrato)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185576

    End If
    
    'Se marcou para inclui no cronograma físico financeiro tem que desmarcar os outros se for um recebimento
    If objPRJRecebPagto.iIncluiCFF = MARCADO And iIncluiCFF = DESMARCADO And objPRJRecebPagto.iTipo = PRJ_TIPO_RECEB Then
    
        'Le a tabelaPRJRecebPagto
        lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM PRJRecebPagto WHERE NumIntDocPRJ = ? AND Tipo = ? AND NumIntDoc <> ?", _
                    0, lNumIntDoc, objPRJRecebPagto.lNumIntDocPRJ, objPRJRecebPagto.iTipo, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 187538
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187539
    
        If lErro = AD_SQL_SUCESSO Then
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(3))
        If lErro <> SUCESSO Then gError 187540

            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJRecebPagto SET IncluiCFF= ?", alComando(3), DESMARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 187541

        End If

    End If
    
    lErro = PRJRecebPagtoRegras_Grava_EmTrans(objPRJRecebPagto)
    If lErro <> SUCESSO Then gError 185730

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJRecebPagto_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Grava_EmTrans:

    PRJRecebPagto_Grava_EmTrans = gErr

    Select Case gErr

        Case 185570, 185730

        Case 185571
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185572, 185573, 187486, 187487, 187538, 187539
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)

        Case 185574, 187540
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJRECEBPAGTO", gErr)

        Case 185575, 187541
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJRECEBPAGTO", gErr)

        Case 185576
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJRECEBPAGTO", gErr)
            
        Case 187488
            If objPRJRecebPagto.iTipo = PRJ_TIPO_RECEB Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PROPOSTA_PROJETO_REPETIDO_RECEB", gErr, lNumero, iFilialEmpresa)
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_PROPOSTA_PROJETO_REPETIDO_PAGTO", gErr, lNumero, iFilialEmpresa)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185717)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJRecebPagtoRegras_Grava_EmTrans(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPRJRecebPagtoRegras As ClassPRJRecebPagtoRegras
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJRecebPagtoRegras_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185719
    Next
    
    lErro = PRJRecebPagtoRegras_Exclui_EmTrans(objPRJRecebPagto)
    If lErro <> SUCESSO Then gError 185720

    For Each objPRJRecebPagtoRegras In objPRJRecebPagto.colRegras

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJRECEBPAGTOREGRAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 185721

        objPRJRecebPagtoRegras.lNumIntDoc = lNumIntDoc
        objPRJRecebPagtoRegras.lNumIntDocRecebPagto = objPRJRecebPagto.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJRecebPagtoRegras( NumIntDoc, NumIntDocRecebPagto, Regra, Percentual, " & _
                "CondPagto, Observacao) VALUES (?,?,?,?,?,?)", _
                objPRJRecebPagtoRegras.lNumIntDoc, objPRJRecebPagtoRegras.lNumIntDocRecebPagto, objPRJRecebPagtoRegras.sRegra, objPRJRecebPagtoRegras.dPercentual, _
                objPRJRecebPagtoRegras.iCondPagto, objPRJRecebPagtoRegras.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185722

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJRecebPagtoRegras_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJRecebPagtoRegras_Grava_EmTrans:

    PRJRecebPagtoRegras_Grava_EmTrans = gErr

    Select Case gErr

        Case 185719
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 185720, 185721

        Case 185722
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJRECEBPAGTOREGRAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185723)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJRecebPagto_Exclui(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJRecebPagto_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185558

    lErro = CF("PRJRecebPagto_Exclui_EmTrans", objPRJRecebPagto)
    If lErro <> SUCESSO Then gError 185559

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185560

    PRJRecebPagto_Exclui = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Exclui:

    PRJRecebPagto_Exclui = gErr

    Select Case gErr

        Case 185558
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 185559

        Case 185560
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185590)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJRecebPagto_Exclui_EmTrans(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_PRJRecebPagto_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185561
    Next

    'Le a tabelaPRJRecebPagto
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJRecebPagto WHERE Numero= ?  AND FilialEmpresa= ? AND Tipo = ? ", _
                0, lNumIntDoc, objPRJRecebPagto.lNumero, objPRJRecebPagto.iFilialEmpresa, objPRJRecebPagto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 185562

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185563

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 185564

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 185565

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJRecebPagto", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 185566
    
    lErro = PRJRecebPagtoRegras_Exclui_EmTrans(objPRJRecebPagto)
    If lErro <> SUCESSO Then gError 185591

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJRecebPagto_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Exclui_EmTrans:

    PRJRecebPagto_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185561
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185562, 185563
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)

        Case 185564
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJRECEBPAGTO_NAO_CADASTRADO", gErr, objPRJRecebPagto.lNumero, objPRJRecebPagto.iFilialEmpresa)

        Case 185565
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJRECEBPAGTO", gErr)

        Case 185566
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJRECEBPAGTO", gErr)
            
        Case 185591

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185592)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJRecebPagtoRegras_Exclui_EmTrans(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_PRJRecebPagtoRegras_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185584
    Next

    'Le a tabelaPRJRecebPagtoRegras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJRecebPagtoRegras WHERE NumIntDocRecebPagto= ? ", _
                0, lNumIntDoc, objPRJRecebPagto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185585

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185586

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185588
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJRecebPagtoRegras", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 185589
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185587
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJRecebPagtoRegras_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJRecebPagtoRegras_Exclui_EmTrans:

    PRJRecebPagtoRegras_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185584
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185585, 185586, 185587
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTOREGRAS", gErr)

        Case 185588
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJRECEBPAGTOREGRAS", gErr)

        Case 185589
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJRECEBPAGTOREGRAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185583)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PRJRecebPagto_Automatico(lCodigo As Long, ByVal iTipo As Integer) As Long
'retorna o número do proximo

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_PRJRecebPagto_Automatico

    If iTipo = PRJ_TIPO_PAGTO Then
        lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_PRJPAGTO", "PagamentoPRJ", "Numero", lNumAuto)
    Else
        lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_PRJRECEB", "RecebimentoPRJ", "Numero", lNumAuto)
    End If
        
    If lErro <> SUCESSO Then gError 185726
    
    lCodigo = lNumAuto
    
    PRJRecebPagto_Automatico = SUCESSO
    
    Exit Function
    
Erro_PRJRecebPagto_Automatico:

    PRJRecebPagto_Automatico = gErr
    
    Select Case gErr

        Case 185726
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185727)

    End Select

    Exit Function
    
End Function

Function RecebPagto_Calcula_Regras(ByVal objTela As Object, ByVal objRecebPagto As ClassPRJRecebPagto) As Long
'Calcula as regras

Dim lErro As Long
Dim iIndice As Integer
Dim objPadraoContabItem As New ClassPadraoContabItem
Dim colPadraoContabItem As New ClassColPadraoContabItem
Dim colMnemonicos As New Collection
Dim colLancamentos As New ClassColLancamentos
Dim objRecebPagtoRegras As ClassPRJRecebPagtoRegras

On Error GoTo Erro_RecebPagto_Calcula_Regras

    For Each objRecebPagtoRegras In objRecebPagto.colRegras
    
        Set objPadraoContabItem = New ClassPadraoContabItem

        objPadraoContabItem.sHistorico = "DATA_TEXTO(" & objRecebPagtoRegras.sRegra & ")"
    
        colPadraoContabItem.Add objPadraoContabItem
        
    Next
    
    'leitura dos mnemonicos no BD
    lErro = CF("MnemonicoPRJ_Le", colMnemonicos)
    If lErro <> SUCESSO Then gError 187445

    lErro = CF("Executa_Regras_CobrEmail", objTela, colPadraoContabItem, colLancamentos, colMnemonicos)
    If lErro <> SUCESSO Then gError 187446
    
    iIndice = 1
    For iIndice = 1 To objRecebPagto.colRegras.Count
    
        Set objRecebPagtoRegras = objRecebPagto.colRegras.Item(iIndice)
    
        objRecebPagtoRegras.dtRegraValor = StrParaDate(colLancamentos.Item(iIndice).sHistorico)
        
    Next
    
    RecebPagto_Calcula_Regras = SUCESSO

    Exit Function

Erro_RecebPagto_Calcula_Regras:

    RecebPagto_Calcula_Regras = gErr

    Select Case gErr
    
        Case 187445, 187446

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187447)

    End Select
   
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJPropostas_Grava(ByVal objPRJPropostas As ClassPRJPropostas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJPropostas_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185870

    lErro = CF("PRJPropostas_Grava_EmTrans", objPRJPropostas)
    If lErro <> SUCESSO Then gError 185871

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185872

    PRJPropostas_Grava = SUCESSO

    Exit Function

Erro_PRJPropostas_Grava:

    PRJPropostas_Grava = gErr

    Select Case gErr

        Case 185870
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 185871

        Case 185872
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187477)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJPropostas_Grava_EmTrans(ByVal objPRJPropostas As ClassPRJPropostas) As Long

Dim lErro As Long
Dim alComando(0 To 9) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocReceb As Long

On Error GoTo Erro_PRJPropostas_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185874
    Next

    'Le a tabelaPRJPropostas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJPropostas WHERE NumIntDocPRJ= ?  AND Codigo= ? ", _
                0, lNumIntDoc, objPRJPropostas.lNumIntDocPRJ, objPRJPropostas.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185875

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185876

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        objPRJPropostas.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185877

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJPropostas SET Data= ? , Observacao= ? , Cliente= ? , FilialCliente= ? , " & _
                "ValorTotal= ? , ValorProdutos= ? , ValorFrete= ? , ValorDesconto= ? , ValorSeguro= ? , " & _
                "ValorOutrasDespesas= ? , CustoInformado= ? , CustoCalculado= ?, ExibirProdutos = ?, ExibirCustoCalc = ?, ExibirCustoInfo = ?, ExibirPreco = ?, NaturezaOp = ? ", alComando(0), _
                objPRJPropostas.dtData, objPRJPropostas.sObservacao, objPRJPropostas.lCliente, objPRJPropostas.iFilialCliente, _
                objPRJPropostas.dValorTotal, objPRJPropostas.dValorProdutos, objPRJPropostas.dValorFrete, objPRJPropostas.dValorDesconto, objPRJPropostas.dValorSeguro, _
                objPRJPropostas.dValorOutrasDespesas, objPRJPropostas.dCustoInformado, objPRJPropostas.dCustoCalculado, objPRJPropostas.iExibirProdutos, objPRJPropostas.iExibirCustoCalc, objPRJPropostas.iExibirCustoInfo, objPRJPropostas.iExibirPreco, objPRJPropostas.sNaturezaOp)
        If lErro <> AD_SQL_SUCESSO Then gError 185878

        lErro = PRJPropostaItens_Exclui(objPRJPropostas)
        If lErro <> SUCESSO Then gError 187508
        
        lErro = PRJPropostaEtapa_Exclui(objPRJPropostas)
        If lErro <> SUCESSO Then gError 187533

        lErro = PRJPropostaTributacao_Exclui(objPRJPropostas, alComando(2), alComando(3), alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 187557
        
        'Le o recebimento
        lErro = Comando_ExecutarPos(alComando(8), "SELECT NumIntDoc FROM PRJRecebPagto WHERE NumIntDocPRJ= ?  AND NumIntDocProposta= ? AND Tipo = ? ", _
                    0, lNumIntDocReceb, objPRJPropostas.lNumIntDocPRJ, lNumIntDoc, PRJ_TIPO_RECEB)
        If lErro <> AD_SQL_SUCESSO Then gError 189012
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189013
    
        'Se existir => UPDATE
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(9), "UPDATE PRJRecebPagto SET Valor = ?, CliForn = ?, Filial = ?", alComando(8), objPRJPropostas.dValorTotal, objPRJPropostas.lCliente, objPRJPropostas.iFilialCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 189014
        
        End If

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJPROPOSTAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 185873

        objPRJPropostas.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJPropostas( NumIntDoc, NumIntDocPRJ, Codigo, Data, " & _
                "Observacao, Cliente, FilialCliente, ValorTotal, ValorProdutos, " & _
                "ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, CustoInformado, " & _
                "CustoCalculado, FilialEmpresa, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco, NaturezaOP) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJPropostas.lNumIntDoc, objPRJPropostas.lNumIntDocPRJ, objPRJPropostas.sCodigo, objPRJPropostas.dtData, _
                objPRJPropostas.sObservacao, objPRJPropostas.lCliente, objPRJPropostas.iFilialCliente, objPRJPropostas.dValorTotal, objPRJPropostas.dValorProdutos, _
                objPRJPropostas.dValorFrete, objPRJPropostas.dValorDesconto, objPRJPropostas.dValorSeguro, objPRJPropostas.dValorOutrasDespesas, objPRJPropostas.dCustoInformado, _
                objPRJPropostas.dCustoCalculado, objPRJPropostas.iFilialEmpresa, objPRJPropostas.iExibirProdutos, objPRJPropostas.iExibirCustoCalc, objPRJPropostas.iExibirCustoInfo, objPRJPropostas.iExibirPreco, objPRJPropostas.sNaturezaOp)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185879

    End If
    
    lErro = PRJPropostaItens_Insere(objPRJPropostas)
    If lErro <> SUCESSO Then gError 187509

    lErro = PRJPropostaEtapa_Insere(objPRJPropostas)
    If lErro <> SUCESSO Then gError 187534
    
    lErro = PRJPropostaTributacao_Grava(objPRJPropostas, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 187557
    
    lErro = TiposCamposCust_Grava_EmTrans(objPRJPropostas.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 189036
    
    objPRJPropostas.objCamposCust.lNumIntDocOrigem = objPRJPropostas.lNumIntDoc
    objPRJPropostas.objCamposCust.iTipoNumIntDocOrigem = objPRJPropostas.objTiposCamposCust.iCodigo
    
    lErro = CamposCustomizados_Grava_EmTrans(objPRJPropostas.objCamposCust)
    If lErro <> SUCESSO Then gError 189037

    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJPropostas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJPropostas_Grava_EmTrans:

    PRJPropostas_Grava_EmTrans = gErr

    Select Case gErr

        Case 185873, 187508, 187509, 187533, 187534, 187557, 189036, 189037

        Case 185874
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185875, 185876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)

        Case 185877
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJPROPOSTAS", gErr)

        Case 185878
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJPROPOSTAS", gErr)

        Case 185879
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJPROPOSTAS", gErr)
            
        Case 189012, 189013
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)
        
        Case 189014
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJRECEBPAGTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187476)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJPropostas_Exclui(ByVal objPRJPropostas As ClassPRJPropostas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJPropostas_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185861

    lErro = CF("PRJPropostas_Exclui_EmTrans", objPRJPropostas)
    If lErro <> SUCESSO Then gError 185862

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185863

    PRJPropostas_Exclui = SUCESSO

    Exit Function

Erro_PRJPropostas_Exclui:

    PRJPropostas_Exclui = gErr

    Select Case gErr

        Case 185861
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 185862

        Case 185863
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187479)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJPropostas_Exclui_EmTrans(ByVal objPRJPropostas As ClassPRJPropostas) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJPropostas_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185864
    Next

    'Le a tabelaPRJPropostas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJPropostas WHERE NumIntDocPRJ= ?  AND Codigo= ? ", _
                0, lNumIntDoc, objPRJPropostas.lNumIntDocPRJ, objPRJPropostas.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185865

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185866

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 185867

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 185868
    
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_PROPOSTA, lNumIntDoc)
    If lErro <> SUCESSO Then gError 189019

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJPropostas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 185869
    
    objPRJPropostas.lNumIntDoc = lNumIntDoc
    
    lErro = PRJPropostaItens_Exclui(objPRJPropostas)
    If lErro <> SUCESSO Then gError 187507

    lErro = PRJPropostaEtapa_Exclui(objPRJPropostas)
    If lErro <> SUCESSO Then gError 187532
    
    lErro = PRJPropostaTributacao_Exclui(objPRJPropostas, alComando(2), alComando(3), alComando(4), alComando(5))
    If lErro <> SUCESSO Then gError 187576

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJPropostas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJPropostas_Exclui_EmTrans:

    PRJPropostas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185864
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185865, 185866
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)

        Case 185867
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJPROPOSTAS_NAO_CADASTRADO", gErr, objPRJPropostas.lNumIntDocPRJ, objPRJPropostas.sCodigo)

        Case 185868
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJPROPOSTAS", gErr)

        Case 185869
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJPROPOSTAS", gErr)
            
        Case 187507, 187532, 187576, 189019

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187478)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJPropostaItens_Insere(ByVal objPRJProposta As ClassPRJPropostas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPRJPropostaItem As ClassPRJPropostaItem
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJPropostaItens_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187495
    Next
    
    For Each objPRJPropostaItem In objPRJProposta.colItens
    
        If objPRJPropostaItem.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJPROPOSTAITENS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 187496
    
            objPRJPropostaItem.lNumIntDoc = lNumIntDoc
            
        End If
        
        objPRJPropostaItem.lNumIntDocProposta = objPRJProposta.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJPropostaItens( NumIntDoc, FilialEmpresa, Item, NumIntDocProposta, " & _
                "Produto, DescProd, NumIntDocEtapa, DescEtapa, Quantidade, " & _
                "PrecoUnitario, PrecoTotal, UnidadeMed, ValorDesconto, DataEntrega, " & _
                "Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJPropostaItem.lNumIntDoc, objPRJPropostaItem.iFilialEmpresa, objPRJPropostaItem.iItem, objPRJPropostaItem.lNumIntDocProposta, _
                objPRJPropostaItem.sProduto, objPRJPropostaItem.sDescProd, objPRJPropostaItem.lNumIntDocEtapa, objPRJPropostaItem.sDescEtapa, objPRJPropostaItem.dQuantidade, _
                objPRJPropostaItem.dPrecoUnitario, objPRJPropostaItem.dPrecoTotal, objPRJPropostaItem.sUM, objPRJPropostaItem.dValorDesconto, objPRJPropostaItem.dtDataEntrega, _
                objPRJPropostaItem.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187497
        
        lErro = PRJPropostaItemTributacao_Grava(objPRJPropostaItem, alComando(0))
        If lErro <> SUCESSO Then gError 187577

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJPropostaItens_Insere = SUCESSO

    Exit Function

Erro_PRJPropostaItens_Insere:

    PRJPropostaItens_Insere = gErr

    Select Case gErr

        Case 187495
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187496, 187577
        
        Case 187497
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJPROPOSTAITENS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187498)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJPropostaItens_Exclui(ByVal objPRJProposta As ClassPRJPropostas) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPropostaItem As ClassPRJPropostaItem

On Error GoTo Erro_PRJPropostaItens_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187500
    Next

    'Le a tabelaPRJPropostaItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJPropostaItens WHERE NumIntDocProposta= ? ", _
                0, lNumIntDoc, objPRJProposta.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187501

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187502

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPropostaItem = New ClassPRJPropostaItem
        
        objPropostaItem.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 187503
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJPropostaItens", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 187504
        
        lErro = PRJPropostaItemTributacao_Exclui(objPropostaItem, alComando(2), alComando(3))
        If lErro <> SUCESSO Then gError 187558
        
        'Busca Próximo
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187505
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJPropostaItens_Exclui = SUCESSO

    Exit Function

Erro_PRJPropostaItens_Exclui:

    PRJPropostaItens_Exclui = gErr

    Select Case gErr

        Case 187500
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187501, 187502, 187505
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAITENS", gErr)
 
        Case 187503
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJPROPOSTAITENS", gErr)

        Case 187504
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJPROPOSTAITENS", gErr)
            
        Case 187558

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187505)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJPropostaEtapa_Insere(ByVal objPRJProposta As ClassPRJPropostas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPRJPropostaEtapa As ClassPRJPropostaEtapa
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJPropostaEtapa_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187521
    Next

    For Each objPRJPropostaEtapa In objPRJProposta.colEtapas

        If objPRJPropostaEtapa.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJPROPOSTAETAPA", lNumIntDoc)
            If lErro <> SUCESSO Then gError 187522
    
            objPRJPropostaEtapa.lNumIntDoc = lNumIntDoc
            
        End If
        
        objPRJPropostaEtapa.lNumIntDocProposta = objPRJProposta.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJPropostaEtapa( NumIntDoc, NumIntDocProposta, NumIntDocEtapa, NumIntDocEtapaItemProd, " & _
                "CustoInformado, Preco, Selecionado,Imprimir,Descricao,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objPRJPropostaEtapa.lNumIntDoc, objPRJPropostaEtapa.lNumIntDocProposta, objPRJPropostaEtapa.lNumIntDocEtapa, objPRJPropostaEtapa.lNumIntDocEtapaItemProd, _
                objPRJPropostaEtapa.dCustoInformado, objPRJPropostaEtapa.dPreco, objPRJPropostaEtapa.iSelecionado, objPRJPropostaEtapa.iImprimir, objPRJPropostaEtapa.sDescricao, objPRJPropostaEtapa.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187523

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJPropostaEtapa_Insere = SUCESSO

    Exit Function

Erro_PRJPropostaEtapa_Insere:

    PRJPropostaEtapa_Insere = gErr

    Select Case gErr

        Case 187521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 187522

        Case 187523
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJPROPOSTAETAPA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187524)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJPropostaEtapa_Exclui(ByVal objPRJProposta As ClassPRJPropostas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJPropostaEtapa_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187526
    Next

    'Le a tabelaPRJPropostaEtapa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJPropostaEtapa WHERE NumIntDocProposta= ? ", _
                0, lNumIntDoc, objPRJProposta.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187527

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187528

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 187529
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJPropostaEtapa", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 187530
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187531
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJPropostaEtapa_Exclui = SUCESSO

    Exit Function

Erro_PRJPropostaEtapa_Exclui:

    PRJPropostaEtapa_Exclui = gErr

    Select Case gErr

        Case 187526
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187527, 187528, 187531
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAETAPA", gErr)

        Case 187529
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJPROPOSTAETAPA", gErr)

        Case 187530
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJPROPOSTAETAPA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187531)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJPropostaTributacao_Exclui(ByVal objProposta As ClassPRJPropostas, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long

Dim lErro As Long
Dim iItem As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PRJPropostaTributacao_Exclui

'    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM TributacaoComplPRJProp WHERE FilialEmpresa = ? AND CodProposta = ? ", 0, iItem, objProposta.iFilialEmpresa, objProposta.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 187559
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187560
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoComplPRJProp", lComando)
'        If lErro <> AD_SQL_SUCESSO Then gError 187561
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187562
'
'    Loop
'
'    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa FROM TributacaoPRJProp WHERE FilialEmpresa = ? AND CodProposta = ? ", 0, iFilialEmpresa, objProposta.iFilialEmpresa, objProposta.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 187563
'
'    lErro = Comando_BuscarPrimeiro(lComando2)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187564
'    If lErro <> AD_SQL_SUCESSO Then gError 187565
'
'    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM TributacaoPRJProp", lComando2)
'    If lErro <> AD_SQL_SUCESSO Then gError 187566

    lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_PRJ_PROP, objProposta.lNumIntDoc)
    If lErro <> SUCESSO Then gError 187566

    PRJPropostaTributacao_Exclui = SUCESSO

    Exit Function

Erro_PRJPropostaTributacao_Exclui:

    PRJPropostaTributacao_Exclui = gErr

    Select Case gErr

'        Case 187559, 187560, 187562
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLPROPOSTA", gErr)
'
'        Case 187561
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLPROPOSTA", gErr)
'
'        Case 187563, 187564
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBPROPOSTA", gErr)
'
'        Case 187565
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBPROPOSTA_NAO_ENCONTRADA", gErr, objProposta.sCodigo)

        Case 187566
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBPROPOSTA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187567)

    End Select

    Exit Function

End Function

Public Function PRJPropostaTributacao_Grava(ByVal objProposta As ClassPRJPropostas, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_PRJPropostaTributacao_Grava
'
'    With objProposta.objTributacaoPRJProp
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoPRJProp (FilialEmpresa, CodProposta, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objProposta.iFilialEmpresa, objProposta.sCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187568
'
'    'inserir os complementos
'    lErro = PRJPropostaTributacao_GravaCompl(objProposta, objProposta.objTributacaoPRJProp.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then gError 187569
'
'    lErro = PRJPropostaTributacao_GravaCompl(objProposta, objProposta.objTributacaoPRJProp.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then gError 187570
'
'    lErro = PRJPropostaTributacao_GravaCompl(objProposta, objProposta.objTributacaoPRJProp.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then gError 187571
'
'    lErro = PRJPropostaTributacao_GravaCompl(objProposta, objProposta.objTributacaoPRJProp.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then gError 187572

    objProposta.objTributacao.lNumIntDoc = objProposta.lNumIntDoc
    objProposta.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_PROP

    lErro = CF("TributacaoDoc_Grava", objProposta.objTributacao)
    If lErro <> SUCESSO Then gError 187572

    PRJPropostaTributacao_Grava = SUCESSO

    Exit Function

Erro_PRJPropostaTributacao_Grava:

    PRJPropostaTributacao_Grava = gErr

    Select Case gErr

        Case 187568
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBNFISCAL", gErr)

        Case 187569 To 187572

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187573)

    End Select

    Exit Function

End Function

'Public Function PRJPropostaTributacao_GravaCompl(ByVal objProposta As ClassPRJPropostas, ByVal objTribComplPRJProp As ClassTributacaoItemPV, ByVal lComando As Long) As Long
'
'Dim lErro As Long, iItem As Integer
'
'On Error GoTo Erro_PRJPropostaTributacao_GravaCompl
'
'    iItem = objTribComplPRJProp.lNumIntDoc
'
'    With objTribComplPRJProp
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplPRJProp (FilialEmpresa, CodProposta, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito  ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objProposta.iFilialEmpresa, objProposta.sCodigo, iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187574
'
'    PRJPropostaTributacao_GravaCompl = SUCESSO
'
'    Exit Function
'
'Erro_PRJPropostaTributacao_GravaCompl:
'
'    PRJPropostaTributacao_GravaCompl = gErr
'
'    Select Case gErr
'
'        Case 187574
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLOV", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187575)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function PRJPropostaItemTributacao_Grava(ByVal objPropostaItem As ClassPRJPropostaItem, ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_PRJPropostaItemTributacao_Grava

'    With objPropostaItem.objTributacaoPRJPropItem
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoPRJPropItem (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito  ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objPropostaItem.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187578

    objPropostaItem.objTributacao.lNumIntDoc = objPropostaItem.lNumIntDocProposta
    objPropostaItem.objTributacao.lNumIntDocItem = objPropostaItem.lNumIntDoc
    objPropostaItem.objTributacao.iItem = objPropostaItem.iItem
    objPropostaItem.objTributacao.dDescontoGrid = objPropostaItem.dValorDesconto
    objPropostaItem.objTributacao.dPrecoUnitario = objPropostaItem.dPrecoUnitario
    objPropostaItem.objTributacao.dPrecoUnitarioMoeda = objPropostaItem.dPrecoUnitario
    objPropostaItem.objTributacao.dQuantidade = objPropostaItem.dQuantidade
    objPropostaItem.objTributacao.sProduto = objPropostaItem.sProduto
    objPropostaItem.objTributacao.sProdutoDescricao = objPropostaItem.sDescProd
    objPropostaItem.objTributacao.sUnidadeMed = objPropostaItem.sUM
    objPropostaItem.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_PROP

    lErro = CF("TributacaoDocItem_Grava", objPropostaItem.objTributacao)
    If lErro <> SUCESSO Then gError 187578

    PRJPropostaItemTributacao_Grava = SUCESSO

    Exit Function

Erro_PRJPropostaItemTributacao_Grava:

    PRJPropostaItemTributacao_Grava = gErr

    Select Case gErr

        Case 187578
            'Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMOV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187579)

    End Select

    Exit Function

End Function

Public Function PRJPropostaItemTributacao_Exclui(ByVal objPropostaItem As ClassPRJPropostaItem, ByVal lComando As Long, ByVal lComando1 As Long) As Long

Dim lErro As Long
Dim iNaturezaOpManual As Integer

On Error GoTo Erro_PRJPropostaItemTributacao_Exclui

'    lErro = Comando_ExecutarPos(lComando, "SELECT NaturezaOpManual FROM TributacaoPRJPropItem WHERE NumIntDoc = ? ", 0, iNaturezaOpManual, objPropostaItem.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 187580
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187581
'    If lErro <> AD_SQL_SUCESSO Then gError 187582
'
'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoPRJPropItem", lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 187583
    
    lErro = CF("TributacaoDocItem_Exclui", TIPODOC_TRIB_PRJ_PROP, objPropostaItem.lNumIntDoc)
    If lErro <> SUCESSO Then gError 187583

    PRJPropostaItemTributacao_Exclui = SUCESSO

    Exit Function

Erro_PRJPropostaItemTributacao_Exclui:

    PRJPropostaItemTributacao_Exclui = gErr

    Select Case gErr
'
'        Case 187580, 187581
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMPROPOSTA", gErr)
'
'        Case 187582
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBITEMPROPOSTA_NAO_CADASTRADO", gErr)

        Case 187583
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMPROPOSTA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187584)

    End Select

    Exit Function

End Function

Function Segmento_Grava_Projeto(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Segmento_Grava_Projeto

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189163

    lErro = Segmento_Grava_Projeto_Trans(colSegmentos)
    If lErro <> SUCESSO Then gError 189164
     
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189165
     
    Segmento_Grava_Projeto = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_Projeto:

    Segmento_Grava_Projeto = gErr

    Select Case gErr

        Case 189163
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 189164

        Case 189165
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189166)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function Segmento_Grava_Projeto_Trans(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento de Produto
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim iPossui As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tSegmento As typeSegmento
Dim sCodigo As String
Dim sCodigo2 As String
Dim objSegmento As New ClassSegmento

On Error GoTo Erro_Segmento_Grava_Projeto_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189167
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 189168
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 189169
  
    sCodigo = String(STRING_SEGMENTO_CODIGO, 0)
    sCodigo2 = String(STRING_SEGMENTO_CODIGO, 0)
    tSegmento.sDelimitador = String(STRING_SEGMENTO_DELIMITADOR, 0)
    
    'pega um objeto da colSegmentos para saber qual o formato( = campo codigo na tabela Segmento)
    For Each objSegmento In colSegmentos
        'guarda o formato em sCodigo para depois fazer o Select na tabela Segmentos
        sCodigo = objSegmento.sCodigo
        Exit For
    Next
       
    If objSegmento.sCodigo <> SEGMENTO_PROJETO Then gError 189170
       
    'verifica existencia de pelo menos um produto cadastrado
    lErro = CF("Projeto_ExisteProjeto", iPossui)
    If lErro <> SUCESSO Then gError 189171

    If iPossui = MARCADO Then
    'so faz modificacao no delimitador e no preenchimento da tabela Segmento
        
         lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel, Delimitador, Preenchimento FROM Segmento WHERE Codigo = ? ORDER BY Nivel", 0, sCodigo2, tSegmento.iNivel, tSegmento.sDelimitador, tSegmento.iPreenchimento, sCodigo)
         If lErro <> AD_SQL_SUCESSO Then gError 189172
     
         'percorre cada registro em colSegmentos para fazer a modificacao nos delimitadores e no preenchimento
         For Each objSegmento In colSegmentos
                     
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO Then gError 189173
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE Segmento SET Codigo =? , Nivel =? , Delimitador =? , Preenchimento =? ", lComando, sCodigo, objSegmento.iNivel, objSegmento.sDelimitador, objSegmento.iPreenchimento)
            If lErro <> AD_SQL_SUCESSO Then gError 189174
            
         Next
    
    Else
    
        'Seleciona registros na tabela de Segmento
        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel FROM Segmento WHERE Codigo = ?", 0, sCodigo2, tSegmento.iNivel, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 189175
        
        'le o primeiro segmento
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189176
        
        'verifica se ainda existe registro a ser deletado
        Do While lErro <> AD_SQL_SEM_DADOS
           
            'exclui o segmento
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Segmento", lComando)
            If lErro <> AD_SQL_SUCESSO Then gError 189177
        
            'le o proximo segmento
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189178
        
        Loop
        
        For Each objSegmento In colSegmentos
           
           lErro = Comando_Executar(lComando2, "INSERT INTO Segmento (Codigo,Nivel,Tipo,Tamanho,Delimitador,Preenchimento) VALUES (?,?,?,?,?,?)", sCodigo, objSegmento.iNivel, objSegmento.iTipo, objSegmento.iTamanho, objSegmento.sDelimitador, objSegmento.iPreenchimento)
           If lErro <> AD_SQL_SUCESSO Then gError 189179
           
        Next
        
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
     
    Segmento_Grava_Projeto_Trans = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_Projeto_Trans:

    Segmento_Grava_Projeto_Trans = gErr

    Select Case gErr

        Case 189167, 189168, 189169
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 189170
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_PROJETO_INVALIDO", gErr, sCodigo)

        Case 189171

        Case 189172, 189173, 189175, 189176, 189178
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SEGMENTO1", gErr, sCodigo)
             
        Case 189174
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SEGMENTO", gErr, sCodigo, tSegmento.iNivel)

        Case 189177
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SEGMENTO", gErr, sCodigo, tSegmento.iNivel)

        Case 189179
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SEGMENTO", gErr, sCodigo, objSegmento.iNivel)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189180)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Segmento_Grava_RefEtapa(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Segmento_Grava_RefEtapa

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189183

    lErro = Segmento_Grava_RefEtapa_Trans(colSegmentos)
    If lErro <> SUCESSO Then gError 189184
     
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189185
     
    Segmento_Grava_RefEtapa = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_RefEtapa:

    Segmento_Grava_RefEtapa = gErr

    Select Case gErr

        Case 189183
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 189184

        Case 189185
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189186)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function Segmento_Grava_RefEtapa_Trans(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento de Produto
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim iPossui As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tSegmento As typeSegmento
Dim sCodigo As String
Dim sCodigo2 As String
Dim objSegmento As New ClassSegmento

On Error GoTo Erro_Segmento_Grava_RefEtapa_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189187
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 189188
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 189189
  
    sCodigo = String(STRING_SEGMENTO_CODIGO, 0)
    sCodigo2 = String(STRING_SEGMENTO_CODIGO, 0)
    tSegmento.sDelimitador = String(STRING_SEGMENTO_DELIMITADOR, 0)
    
    'pega um objeto da colSegmentos para saber qual o formato( = campo codigo na tabela Segmento)
    For Each objSegmento In colSegmentos
        'guarda o formato em sCodigo para depois fazer o Select na tabela Segmentos
        sCodigo = objSegmento.sCodigo
        Exit For
    Next
       
    If objSegmento.sCodigo <> SEGMENTO_REFETAPA Then gError 189190
       
    'verifica existencia de pelo menos um produto cadastrado
    lErro = CF("RefEtapa_ExisteRefEtapa", iPossui)
    If lErro <> SUCESSO Then gError 189191

    If iPossui = MARCADO Then
    'so faz modificacao no delimitador e no preenchimento da tabela Segmento
        
         lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel, Delimitador, Preenchimento FROM Segmento WHERE Codigo = ? ORDER BY Nivel", 0, sCodigo2, tSegmento.iNivel, tSegmento.sDelimitador, tSegmento.iPreenchimento, sCodigo)
         If lErro <> AD_SQL_SUCESSO Then gError 189192
     
         'percorre cada registro em colSegmentos para fazer a modificacao nos delimitadores e no preenchimento
         For Each objSegmento In colSegmentos
                     
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO Then gError 189193
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE Segmento SET Codigo =? , Nivel =? , Delimitador =? , Preenchimento =? ", lComando, sCodigo, objSegmento.iNivel, objSegmento.sDelimitador, objSegmento.iPreenchimento)
            If lErro <> AD_SQL_SUCESSO Then gError 189194
            
         Next
    
    Else
    
        'Seleciona registros na tabela de Segmento
        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel FROM Segmento WHERE Codigo = ?", 0, sCodigo2, tSegmento.iNivel, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 189195
        
        'le o primeiro segmento
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189196
        
        'verifica se ainda existe registro a ser deletado
        Do While lErro <> AD_SQL_SEM_DADOS
           
            'exclui o segmento
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Segmento", lComando)
            If lErro <> AD_SQL_SUCESSO Then gError 189197
        
            'le o proximo segmento
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189198
        
        Loop
        
        For Each objSegmento In colSegmentos
           
           lErro = Comando_Executar(lComando2, "INSERT INTO Segmento (Codigo,Nivel,Tipo,Tamanho,Delimitador,Preenchimento) VALUES (?,?,?,?,?,?)", sCodigo, objSegmento.iNivel, objSegmento.iTipo, objSegmento.iTamanho, objSegmento.sDelimitador, objSegmento.iPreenchimento)
           If lErro <> AD_SQL_SUCESSO Then gError 189199
           
        Next
        
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
     
    Segmento_Grava_RefEtapa_Trans = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_RefEtapa_Trans:

    Segmento_Grava_RefEtapa_Trans = gErr

    Select Case gErr

        Case 189187, 189188, 189189
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 189190
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_REFETAPA_INVALIDO", gErr, sCodigo)

        Case 189191

        Case 189192, 189193, 189195, 189196, 189198
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SEGMENTO1", gErr, sCodigo)
             
        Case 189194
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SEGMENTO", gErr, sCodigo, tSegmento.iNivel)

        Case 189197
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SEGMENTO", gErr, sCodigo, tSegmento.iNivel)

        Case 189199
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SEGMENTO", gErr, sCodigo, objSegmento.iNivel)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189200)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function
































'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJContratos_Grava(ByVal objPRJContratos As ClassPRJContratos) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJContratos_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185870

    lErro = CF("PRJContratos_Grava_EmTrans", objPRJContratos)
    If lErro <> SUCESSO Then gError 185871

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185872

    PRJContratos_Grava = SUCESSO

    Exit Function

Erro_PRJContratos_Grava:

    PRJContratos_Grava = gErr

    Select Case gErr

        Case 185870
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 185871

        Case 185872
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187477)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJContratos_Grava_EmTrans(ByVal objPRJContratos As ClassPRJContratos) As Long

Dim lErro As Long
Dim alComando(0 To 13) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocReceb As Long
Dim lNumIntDocProp As Long
Dim lNumIntDocPropAux As Long

On Error GoTo Erro_PRJContratos_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185874
    Next

    'Le a tabelaPRJContratos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocProposta FROM PRJContratos WHERE NumIntDocPRJ= ?  AND Codigo= ? ", _
                0, lNumIntDoc, lNumIntDocProp, objPRJContratos.lNumIntDocPRJ, objPRJContratos.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185875

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185876

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        objPRJContratos.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 185877

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJContratos SET Data= ? , Observacao= ? , Cliente= ? , FilialCliente= ? , " & _
                "ValorTotal= ? , ValorProdutos= ? , ValorFrete= ? , ValorDesconto= ? , ValorSeguro= ? , " & _
                "ValorOutrasDespesas= ? , CustoInformado= ? , CustoCalculado= ?, ExibirProdutos = ?, ExibirCustoCalc = ?, ExibirCustoInfo = ?, ExibirPreco = ?, NaturezaOp = ? ", alComando(0), _
                objPRJContratos.dtData, objPRJContratos.sObservacao, objPRJContratos.lCliente, objPRJContratos.iFilialCliente, _
                objPRJContratos.dValorTotal, objPRJContratos.dValorProdutos, objPRJContratos.dValorFrete, objPRJContratos.dValorDesconto, objPRJContratos.dValorSeguro, _
                objPRJContratos.dValorOutrasDespesas, objPRJContratos.dCustoInformado, objPRJContratos.dCustoCalculado, objPRJContratos.iExibirProdutos, objPRJContratos.iExibirCustoCalc, objPRJContratos.iExibirCustoInfo, objPRJContratos.iExibirPreco, objPRJContratos.sNaturezaOp)
        If lErro <> AD_SQL_SUCESSO Then gError 185878

        lErro = PRJContratoItens_Exclui(objPRJContratos)
        If lErro <> SUCESSO Then gError 187508
        
        lErro = PRJContratoEtapa_Exclui(objPRJContratos)
        If lErro <> SUCESSO Then gError 187533

        lErro = PRJContratoTributacao_Exclui(objPRJContratos, alComando(2), alComando(3), alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 187557
        
        'Le o recebimento
        lErro = Comando_ExecutarPos(alComando(8), "SELECT NumIntDoc FROM PRJRecebPagto WHERE NumIntDocPRJ= ?  AND NumIntDocContrato= ? AND Tipo = ? ", _
                    0, lNumIntDocReceb, objPRJContratos.lNumIntDocPRJ, lNumIntDoc, PRJ_TIPO_RECEB)
        If lErro <> AD_SQL_SUCESSO Then gError 189012
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189013
    
        'Se existir => UPDATE
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(9), "UPDATE PRJRecebPagto SET Valor = ?, CliForn = ?, Filial = ?, NumIntDocProposta = ? ", alComando(8), objPRJContratos.dValorTotal, objPRJContratos.lCliente, objPRJContratos.iFilialCliente, objPRJContratos.lNumIntDocProposta)
            If lErro <> AD_SQL_SUCESSO Then gError 189014
        
        End If

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJCONTRATOS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 185873

        objPRJContratos.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJContratos( NumIntDoc, NumIntDocPRJ,NumIntDocProposta, Codigo, Data, " & _
                "Observacao, Cliente, FilialCliente, ValorTotal, ValorProdutos, " & _
                "ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, CustoInformado, " & _
                "CustoCalculado, FilialEmpresa, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco, NaturezaOP) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJContratos.lNumIntDoc, objPRJContratos.lNumIntDocPRJ, objPRJContratos.lNumIntDocProposta, objPRJContratos.sCodigo, objPRJContratos.dtData, _
                objPRJContratos.sObservacao, objPRJContratos.lCliente, objPRJContratos.iFilialCliente, objPRJContratos.dValorTotal, objPRJContratos.dValorProdutos, _
                objPRJContratos.dValorFrete, objPRJContratos.dValorDesconto, objPRJContratos.dValorSeguro, objPRJContratos.dValorOutrasDespesas, objPRJContratos.dCustoInformado, _
                objPRJContratos.dCustoCalculado, objPRJContratos.iFilialEmpresa, objPRJContratos.iExibirProdutos, objPRJContratos.iExibirCustoCalc, objPRJContratos.iExibirCustoInfo, objPRJContratos.iExibirPreco, objPRJContratos.sNaturezaOp)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185879

    End If
    
    lErro = PRJContratoItens_Insere(objPRJContratos)
    If lErro <> SUCESSO Then gError 187509

    lErro = PRJContratoEtapa_Insere(objPRJContratos)
    If lErro <> SUCESSO Then gError 187534
    
    lErro = PRJContratoTributacao_Grava(objPRJContratos, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 187557
    
    lErro = TiposCamposCust_Grava_EmTrans(objPRJContratos.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 189036
    
    objPRJContratos.objCamposCust.lNumIntDocOrigem = objPRJContratos.lNumIntDoc
    objPRJContratos.objCamposCust.iTipoNumIntDocOrigem = objPRJContratos.objTiposCamposCust.iCodigo
    
    lErro = CamposCustomizados_Grava_EmTrans(objPRJContratos.objCamposCust)
    If lErro <> SUCESSO Then gError 189037
    
    If lNumIntDocProp <> objPRJContratos.lNumIntDocProposta Then
    
        'Se estava vinculado a uma proposta e já não está mais
        'Acerta a proposta antiga para não mais estar vinculada a esse contrato
        If lNumIntDocProp <> 0 Then
        
            'Le a tabela PRJProposta
            lErro = Comando_ExecutarPos(alComando(10), "SELECT NumIntDoc FROM PRJPropostas WHERE NumIntDoc= ?  ", 0, lNumIntDocPropAux, lNumIntDocProp)
            If lErro <> AD_SQL_SUCESSO Then gError 189384
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189385
            
            If lErro <> AD_SQL_SUCESSO Then gError 189386
            
            lErro = Comando_ExecutarPos(alComando(11), "UPDATE PRJPropostas SET NumIntDocContrato= ? ", alComando(10), 0)
            If lErro <> AD_SQL_SUCESSO Then gError 189387
        
        End If
    
        'Se está vinculando a uma proposta nova
        'Acerta a proposta para apontar para esse contrato
        If objPRJContratos.lNumIntDocProposta <> 0 Then
        
            'Le a tabela PRJProposta
            lErro = Comando_ExecutarPos(alComando(12), "SELECT NumIntDoc FROM PRJPropostas WHERE NumIntDoc= ?  ", 0, lNumIntDocPropAux, objPRJContratos.lNumIntDocProposta)
            If lErro <> AD_SQL_SUCESSO Then gError 189384
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(12))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189385
            
            If lErro <> AD_SQL_SUCESSO Then gError 189386
            
            lErro = Comando_ExecutarPos(alComando(13), "UPDATE PRJPropostas SET NumIntDocContrato= ? ", alComando(12), objPRJContratos.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 189387
        
        End If
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJContratos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJContratos_Grava_EmTrans:

    PRJContratos_Grava_EmTrans = gErr

    Select Case gErr

        Case 185873, 187508, 187509, 187533, 187534, 187557, 189036, 189037

        Case 185874
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185875, 185876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOS", gErr)

        Case 185877
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJCONTRATOS", gErr)

        Case 185878
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJCONTRATOS", gErr)

        Case 185879
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJCONTRATOSs", gErr)
            
        Case 189012, 189013
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)
        
        Case 189014
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJRECEBPAGTO", gErr)
            
        Case 189384, 189385
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)
        
        Case 189386
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJPROPOSTAS_NAO_CADASTRADO2", gErr, objPRJContratos.lNumIntDocProposta)
        
        Case 189387
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJPROPOTAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187476)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJContratos_Exclui(ByVal objPRJContratos As ClassPRJContratos) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJContratos_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185861

    lErro = CF("PRJContratos_Exclui_EmTrans", objPRJContratos)
    If lErro <> SUCESSO Then gError 185862

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185863

    PRJContratos_Exclui = SUCESSO

    Exit Function

Erro_PRJContratos_Exclui:

    PRJContratos_Exclui = gErr

    Select Case gErr

        Case 185861
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 185862

        Case 185863
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187479)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJContratos_Exclui_EmTrans(ByVal objPRJContratos As ClassPRJContratos) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntProposta As Long
Dim lNumIntPropostaAux As Long

On Error GoTo Erro_PRJContratos_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185864
    Next

    'Le a tabelaPRJContratos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocProposta FROM PRJContratos WHERE NumIntDocPRJ= ?  AND Codigo= ? ", _
                0, lNumIntDoc, lNumIntProposta, objPRJContratos.lNumIntDocPRJ, objPRJContratos.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185865

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185866

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 185867

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 185868
    
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_CONTRATOPRJ, lNumIntDoc)
    If lErro <> SUCESSO Then gError 189019

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJContratos", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 185869
    
    objPRJContratos.lNumIntDoc = lNumIntDoc
    
    lErro = PRJContratoItens_Exclui(objPRJContratos)
    If lErro <> SUCESSO Then gError 187507

    lErro = PRJContratoEtapa_Exclui(objPRJContratos)
    If lErro <> SUCESSO Then gError 187532
    
    lErro = PRJContratoTributacao_Exclui(objPRJContratos, alComando(2), alComando(3), alComando(4), alComando(5))
    If lErro <> SUCESSO Then gError 187576
    
    'Desassocia a proposta ao contrato
    If lNumIntProposta <> 0 Then
    
        'Le a tabela PRJProposta
        lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDoc FROM PRJPropostas WHERE NumIntDoc= ?  ", 0, lNumIntPropostaAux, lNumIntProposta)
        If lErro <> AD_SQL_SUCESSO Then gError 189388
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189389
        
        If lErro <> AD_SQL_SUCESSO Then gError 189390
        
        lErro = Comando_ExecutarPos(alComando(7), "UPDATE PRJPropostas SET NumIntDocContrato= ? ", alComando(6), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 189391
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJContratos_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJContratos_Exclui_EmTrans:

    PRJContratos_Exclui_EmTrans = gErr

    Select Case gErr

        Case 185864
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185865, 185866
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOS", gErr)

        Case 185867
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJCONTRATOS_NAO_CADASTRADO", gErr, objPRJContratos.lNumIntDocPRJ, objPRJContratos.sCodigo)

        Case 185868
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJCONTRATOS", gErr)

        Case 185869
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJCONTRATOS", gErr)
            
        Case 187507, 187532, 187576, 189019
        
        Case 189388, 189389
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)
        
        Case 189390
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJPROPOSTAS_NAO_CADASTRADO2", gErr, lNumIntProposta)
        
        Case 189391
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJPROPOTAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187478)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJContratoItens_Insere(ByVal objPRJContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPRJContratoItem As ClassPRJContratoItem
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJContratoItens_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187495
    Next
    
    For Each objPRJContratoItem In objPRJContrato.colItens
    
        If objPRJContratoItem.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJCONTRATOITENS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 187496
    
            objPRJContratoItem.lNumIntDoc = lNumIntDoc
            
        End If
        
        objPRJContratoItem.lNumIntDocContrato = objPRJContrato.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJContratoItens( NumIntDoc, FilialEmpresa, Item, NumIntDocContrato, " & _
                "Produto, DescProd, NumIntDocEtapa, DescEtapa, Quantidade, " & _
                "PrecoUnitario, PrecoTotal, UnidadeMed, ValorDesconto, DataEntrega, " & _
                "Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPRJContratoItem.lNumIntDoc, objPRJContratoItem.iFilialEmpresa, objPRJContratoItem.iItem, objPRJContratoItem.lNumIntDocContrato, _
                objPRJContratoItem.sProduto, objPRJContratoItem.sDescProd, objPRJContratoItem.lNumIntDocEtapa, objPRJContratoItem.sDescEtapa, objPRJContratoItem.dQuantidade, _
                objPRJContratoItem.dPrecoUnitario, objPRJContratoItem.dPrecoTotal, objPRJContratoItem.sUM, objPRJContratoItem.dValorDesconto, objPRJContratoItem.dtDataEntrega, _
                objPRJContratoItem.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187497
        
        lErro = PRJContratoItemTributacao_Grava(objPRJContratoItem, alComando(0))
        If lErro <> SUCESSO Then gError 187577

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJContratoItens_Insere = SUCESSO

    Exit Function

Erro_PRJContratoItens_Insere:

    PRJContratoItens_Insere = gErr

    Select Case gErr

        Case 187495
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187496, 187577
        
        Case 187497
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJCONTRATOITENS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187498)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJContratoItens_Exclui(ByVal objPRJContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objContratoItem As ClassPRJContratoItem

On Error GoTo Erro_PRJContratoItens_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187500
    Next

    'Le a tabelaPRJContratoItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJContratoItens WHERE NumIntDocContrato= ? ", _
                0, lNumIntDoc, objPRJContrato.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187501

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187502

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objContratoItem = New ClassPRJContratoItem
        
        objContratoItem.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 187503
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJContratoItens", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 187504
        
        lErro = PRJContratoItemTributacao_Exclui(objContratoItem, alComando(2), alComando(3))
        If lErro <> SUCESSO Then gError 187558
        
        'Busca Próximo
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187505
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJContratoItens_Exclui = SUCESSO

    Exit Function

Erro_PRJContratoItens_Exclui:

    PRJContratoItens_Exclui = gErr

    Select Case gErr

        Case 187500
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187501, 187502, 187505
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOITENS", gErr)
 
        Case 187503
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJCONTRATOITENS", gErr)

        Case 187504
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJCONTRATOITENS", gErr)
            
        Case 187558

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187505)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJContratoEtapa_Insere(ByVal objPRJContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPRJContratoEtapa As ClassPRJContratoEtapa
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJContratoEtapa_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187521
    Next

    For Each objPRJContratoEtapa In objPRJContrato.colEtapas

        If objPRJContratoEtapa.lNumIntDoc = 0 Then

            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJContratoETAPA", lNumIntDoc)
            If lErro <> SUCESSO Then gError 187522
    
            objPRJContratoEtapa.lNumIntDoc = lNumIntDoc
            
        End If
        
        objPRJContratoEtapa.lNumIntDocContrato = objPRJContrato.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJContratoEtapa( NumIntDoc, NumIntDocContrato, NumIntDocEtapa, NumIntDocEtapaItemProd, " & _
                "CustoInformado, Preco, Selecionado,Imprimir,Descricao,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objPRJContratoEtapa.lNumIntDoc, objPRJContratoEtapa.lNumIntDocContrato, objPRJContratoEtapa.lNumIntDocEtapa, objPRJContratoEtapa.lNumIntDocEtapaItemProd, _
                objPRJContratoEtapa.dCustoInformado, objPRJContratoEtapa.dPreco, objPRJContratoEtapa.iSelecionado, objPRJContratoEtapa.iImprimir, objPRJContratoEtapa.sDescricao, objPRJContratoEtapa.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187523

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJContratoEtapa_Insere = SUCESSO

    Exit Function

Erro_PRJContratoEtapa_Insere:

    PRJContratoEtapa_Insere = gErr

    Select Case gErr

        Case 187521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 187522

        Case 187523
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJCONTRATOETAPA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187524)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJContratoEtapa_Exclui(ByVal objPRJContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PRJContratoEtapa_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 187526
    Next

    'Le a tabelaPRJContratoEtapa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJContratoEtapa WHERE NumIntDocContrato= ? ", _
                0, lNumIntDoc, objPRJContrato.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187527

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187528

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 187529
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJContratoEtapa", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 187530
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187531
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJContratoEtapa_Exclui = SUCESSO

    Exit Function

Erro_PRJContratoEtapa_Exclui:

    PRJContratoEtapa_Exclui = gErr

    Select Case gErr

        Case 187526
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187527, 187528, 187531
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOETAPA", gErr)

        Case 187529
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJCONTRATOETAPA", gErr)

        Case 187530
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJCONTRATOETAPA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187531)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PRJContratoTributacao_Exclui(ByVal objContrato As ClassPRJContratos, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long

Dim lErro As Long
Dim iItem As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PRJContratoTributacao_Exclui

'    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM TributacaoComplPRJCTR WHERE FilialEmpresa = ? AND CodContrato = ? ", 0, iItem, objContrato.iFilialEmpresa, objContrato.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 187559
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187560
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoComplPRJCTR", lComando)
'        If lErro <> AD_SQL_SUCESSO Then gError 187561
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187562
'
'    Loop
'
'    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa FROM TributacaoPRJCTR WHERE FilialEmpresa = ? AND CodContrato = ? ", 0, iFilialEmpresa, objContrato.iFilialEmpresa, objContrato.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 187563
'
'    lErro = Comando_BuscarPrimeiro(lComando2)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187564
'    If lErro <> AD_SQL_SUCESSO Then gError 187565
'
'    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM TributacaoPRJCTR", lComando2)
'    If lErro <> AD_SQL_SUCESSO Then gError 187566

    lErro = CF("TributacaoDoc_Exclui", TIPODOC_TRIB_PRJ_CTR, objContrato.lNumIntDoc)
    If lErro <> SUCESSO Then gError 187566

    PRJContratoTributacao_Exclui = SUCESSO

    Exit Function

Erro_PRJContratoTributacao_Exclui:

    PRJContratoTributacao_Exclui = gErr

    Select Case gErr
'
'        Case 187559, 187560, 187562
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCOMPLCONTRATO", gErr)
'
'        Case 187561
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCOMPLCONTRATO", gErr)
'
'        Case 187563, 187564
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCONTRATO", gErr)
'
'        Case 187565
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBCONTRATO_NAO_ENCONTRADA", gErr, objContrato.sCodigo)

        Case 187566
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBCONTRATO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187567)

    End Select

    Exit Function

End Function

Public Function PRJContratoTributacao_Grava(ByVal objContrato As ClassPRJContratos, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long

On Error GoTo Erro_PRJContratoTributacao_Grava

'    With objContrato.objTributacaoPRJCTR
'        lErro = Comando_Executar(lComando1, "INSERT INTO TributacaoPRJCTR (FilialEmpresa, CodContrato, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objContrato.iFilialEmpresa, objContrato.sCodigo, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187568
'
'    'inserir os complementos
'    lErro = PRJContratoTributacao_GravaCompl(objContrato, objContrato.objTributacaoPRJCTR.objTributacaoDesconto, lComando2)
'    If lErro <> SUCESSO Then gError 187569
'
'    lErro = PRJContratoTributacao_GravaCompl(objContrato, objContrato.objTributacaoPRJCTR.objTributacaoFrete, lComando2)
'    If lErro <> SUCESSO Then gError 187570
'
'    lErro = PRJContratoTributacao_GravaCompl(objContrato, objContrato.objTributacaoPRJCTR.objTributacaoOutras, lComando2)
'    If lErro <> SUCESSO Then gError 187571
'
'    lErro = PRJContratoTributacao_GravaCompl(objContrato, objContrato.objTributacaoPRJCTR.objTributacaoSeguro, lComando2)
'    If lErro <> SUCESSO Then gError 187572

    objContrato.objTributacao.lNumIntDoc = objContrato.lNumIntDoc
    objContrato.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_CTR

    lErro = CF("TributacaoDoc_Grava", objContrato.objTributacao)
    If lErro <> SUCESSO Then gError 187572

    PRJContratoTributacao_Grava = SUCESSO

    Exit Function

Erro_PRJContratoTributacao_Grava:

    PRJContratoTributacao_Grava = gErr

    Select Case gErr

'        Case 187568
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBNFISCAL", gErr)

        Case 187569 To 187572

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187573)

    End Select

    Exit Function

End Function

'Public Function PRJContratoTributacao_GravaCompl(ByVal objContrato As ClassPRJContratos, ByVal objTribComplPRJCTR As ClassTributacaoItemPV, ByVal lComando As Long) As Long
'
'Dim lErro As Long, iItem As Integer
'
'On Error GoTo Erro_PRJContratoTributacao_GravaCompl
'
'    iItem = objTribComplPRJCTR.lNumIntDoc
'
'    With objTribComplPRJCTR
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoComplPRJCTR (FilialEmpresa, CodContrato, Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPICodProduto, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objContrato.iFilialEmpresa, objContrato.sCodigo, iItem, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, "", .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187574
'
'    PRJContratoTributacao_GravaCompl = SUCESSO
'
'    Exit Function
'
'Erro_PRJContratoTributacao_GravaCompl:
'
'    PRJContratoTributacao_GravaCompl = gErr
'
'    Select Case gErr
'
'        Case 187574
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBCOMPLOV", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187575)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function PRJContratoItemTributacao_Grava(ByVal objContratoItem As ClassPRJContratoItem, ByVal lComando As Long) As Long

Dim lErro As Long

On Error GoTo Erro_PRJContratoItemTributacao_Grava

'    With objContratoItem.objTributacaoPRJCTRItem
'
'        lErro = Comando_Executar(lComando, "INSERT INTO TributacaoPRJCTRItem (NumIntDoc, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            objContratoItem.lNumIntDoc, .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187578
    
    objContratoItem.objTributacao.lNumIntDoc = objContratoItem.lNumIntDocContrato
    objContratoItem.objTributacao.lNumIntDocItem = objContratoItem.lNumIntDoc
    objContratoItem.objTributacao.iItem = objContratoItem.iItem
    objContratoItem.objTributacao.dDescontoGrid = objContratoItem.dValorDesconto
    objContratoItem.objTributacao.dPrecoUnitario = objContratoItem.dPrecoUnitario
    objContratoItem.objTributacao.dPrecoUnitarioMoeda = objContratoItem.dPrecoUnitario
    objContratoItem.objTributacao.dQuantidade = objContratoItem.dQuantidade
    objContratoItem.objTributacao.sProduto = objContratoItem.sProduto
    objContratoItem.objTributacao.sProdutoDescricao = objContratoItem.sDescProd
    objContratoItem.objTributacao.sUnidadeMed = objContratoItem.sUM
    objContratoItem.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_CTR

    lErro = CF("TributacaoDocItem_Grava", objContratoItem.objTributacao)
    If lErro <> SUCESSO Then gError 187578

    PRJContratoItemTributacao_Grava = SUCESSO

    Exit Function

Erro_PRJContratoItemTributacao_Grava:

    PRJContratoItemTributacao_Grava = gErr

    Select Case gErr

        Case 187578
            'Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRIBITEMOV", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187579)

    End Select

    Exit Function

End Function

Public Function PRJContratoItemTributacao_Exclui(ByVal objContratoItem As ClassPRJContratoItem, ByVal lComando As Long, ByVal lComando1 As Long) As Long

Dim lErro As Long
Dim iNaturezaOpManual As Integer

On Error GoTo Erro_PRJContratoItemTributacao_Exclui

'    lErro = Comando_ExecutarPos(lComando, "SELECT NaturezaOpManual FROM TributacaoPRJCTRItem WHERE NumIntDoc = ? ", 0, iNaturezaOpManual, objContratoItem.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 187580
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187581
'    If lErro <> AD_SQL_SUCESSO Then gError 187582
'
'    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TributacaoPRJCTRItem", lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 187583

    lErro = CF("TributacaoDocItem_Exclui", TIPODOC_TRIB_PRJ_CTR, objContratoItem.lNumIntDoc)
    If lErro <> SUCESSO Then gError 187583

    PRJContratoItemTributacao_Exclui = SUCESSO

    Exit Function

Erro_PRJContratoItemTributacao_Exclui:

    PRJContratoItemTributacao_Exclui = gErr

    Select Case gErr

'        Case 187580, 187581
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBITEMContrato", gErr)
'
'        Case 187582
'            Call Rotina_Erro(vbOKOnly, "ERRO_TRIBITEMCONTRATO_NAO_CADASTRADO", gErr)

        Case 187583
            'Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRIBITEMCONTRATO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187584)

    End Select

    Exit Function

End Function

Public Function RelPRJCustoReceita_Prepara(lNumIntRel As Long, ByVal lNumIntDocProjeto As Long, ByVal lNumIntDocEtapa As Long, ByVal dtDataDe As Date, ByVal dtDataAte As Date, Optional ByVal sNatDe As String = "", Optional ByVal sNatAte As String = "") As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objFlx As New ClassFluxoPRJ
Dim iIndice As Integer
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iFilialEmpresa As Integer
Dim objFluxoItem As ClassPRJFluxoFinanc
Dim iSeq As Integer
Dim lCliente As Long
Dim sDescricao As String
Dim sProjetoMask As String

On Error GoTo Erro_RelPRJCustoReceita_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194257

    'Abertura de Comando
    For iIndice = 0 To 25
        objFlx.alComando(iIndice) = Comando_Abrir()
        If objFlx.alComando(iIndice) = 0 Then gError 194258
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_FLUXO_PRJ", lNumIntRel)
    If lErro <> SUCESSO Then gError 194259
    
    objFlx.lNumIntRel = lNumIntRel
    objFlx.dtDataDe = dtDataDe
    objFlx.dtDataAte = dtDataAte
    
    sProjeto = String(STRING_PRJ_CODIGO, 0)
    sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    
    If lNumIntDocProjeto = 0 Then
        lErro = Comando_Executar(objFlx.alComando(1), "SELECT NumIntDoc, Codigo, FilialEmpresa, Descricao, Cliente FROM Projetos ", lNumIntDocPRJ, sProjeto, iFilialEmpresa, sDescricao, lCliente)
    Else
        lErro = Comando_Executar(objFlx.alComando(1), "SELECT NumIntDoc, Codigo, FilialEmpresa, Descricao, Cliente FROM Projetos WHERE NumIntDoc = ?", lNumIntDocPRJ, sProjeto, iFilialEmpresa, sDescricao, lCliente, lNumIntDocProjeto)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 194260

    lErro = Comando_BuscarPrimeiro(objFlx.alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194261
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objFlx.objProjeto = New ClassProjetos
        Set objFlx.colFluxo = New Collection
        
        objFlx.objProjeto.sCodigo = sProjeto
        objFlx.objProjeto.lNumIntDoc = lNumIntDocPRJ
        objFlx.objProjeto.iFilialEmpresa = iFilialEmpresa
        objFlx.objProjeto.lCliente = lCliente
        objFlx.objProjeto.sDescricao = sDescricao
        
        objFlx.lNumIntPRJ = lNumIntDocPRJ
        objFlx.lNumIntEtapa = lNumIntDocEtapa
        objFlx.dtDataDe = dtDataDe
        objFlx.dtDataAte = dtDataAte
        
        lErro = CF("Projetos_Le_NumIntDoc", objFlx.objProjeto, objFlx.alComando(0))
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194262

        lErro = CF("RelPRJCustoReceita_Le", objFlx)
        If lErro <> SUCESSO Then gError 194262
        
        Call Retorno_Projeto_Tela2(objFlx.objProjeto.sCodigo, sProjetoMask)
        
        For Each objFluxoItem In objFlx.colFluxo
        
            If (lNumIntDocEtapa = objFluxoItem.objEtapa.lNumIntDoc Or lNumIntDocEtapa = 0) And (dtDataDe = DATA_NULA Or dtDataDe <= objFluxoItem.dtData) And (dtDataAte = DATA_NULA Or dtDataAte >= objFluxoItem.dtData) And ((sNatDe = "" Or sNatDe <= objFluxoItem.sNatureza) And (sNatAte = "" Or sNatAte >= objFluxoItem.sNatureza)) Then
        
                iSeq = iSeq + 1
            
                lErro = Comando_Executar(objFlx.alComando(2), "INSERT INTO RelFluxoPRJ(NumIntRel,Seq,CustoProdPrevCalc,CustoProdPrevInfo,CustoProdRealCalc,CustoProdRealInfo,DespesaPlan,DespesaPrev,DespesaReal,ReceitaPlan,ReceitaPrev,ReceitaReal,Data,Descricao,DocOrigem,NumIntDocPRJ,Projeto,DescricaoPRJ,Cliente,NumIntDocEtapa,Etapa,DescricaoEtapa, Obs, Natureza, NaturezaDesc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objFlx.lNumIntRel, iSeq, objFluxoItem.dCustoProdPrevCalc, objFluxoItem.dCustoProdPrevInfo, objFluxoItem.dCustoProdRealCalc, objFluxoItem.dCustoProdRealInfo, objFluxoItem.dDespesaPlan, objFluxoItem.dDespesaPrev, objFluxoItem.dDespesaReal, objFluxoItem.dReceitaPlan, objFluxoItem.dReceitaPrev, objFluxoItem.dReceitaReal, StrParaDate(objFluxoItem.dtData), objFluxoItem.sDescricao, objFluxoItem.sDocOrigem, objFluxoItem.objProjeto.lNumIntDoc, sProjetoMask, objFluxoItem.objProjeto.sDescricao, objFluxoItem.objProjeto.lCliente, objFluxoItem.objEtapa.lNumIntDoc, objFluxoItem.objEtapa.sCodigo, objFluxoItem.objEtapa.sDescricao, objFluxoItem.sObs, objFluxoItem.sNatureza, objFluxoItem.sDescNatureza)
                If lErro <> AD_SQL_SUCESSO Then
                    gError 194263
                End If
                
            End If
        
        Next
    
        lErro = Comando_BuscarProximo(objFlx.alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194264
    
    Loop
    
    'Fecha Comando
    For iIndice = 0 To 25
        Call Comando_Fechar(objFlx.alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194265

    RelPRJCustoReceita_Prepara = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Prepara:

    RelPRJCustoReceita_Prepara = gErr

    Select Case gErr
    
        Case 194257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 194258
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 194259, 194262
        
        Case 194260, 194261, 194264
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 194263
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELFLUXOPRJ", gErr)

        Case 194265
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194266)

    End Select
    
    Call Transacao_Rollback

    'Fecha Comando
    For iIndice = 0 To 25
        Call Comando_Fechar(objFlx.alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function RelMatUtiPerPRJ_Prepara(lNumIntRel As Long, ByVal lNumIntDocProjeto As Long, ByVal lNumIntDocEtapa As Long, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal sProdutoDe As String, ByVal sProdutoAte As String, ByVal iPorData As Integer) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iFilialEmpresa As Integer
Dim objProjeto As ClassProjetos
Dim iSeq As Integer
Dim lCliente As Long
Dim sDescricao As String
Dim objEtapa As ClassPRJEtapas
Dim objEtapaMat As ClassPRJEtapaMateriais
Dim colMat As Collection
Dim objRelMatUtiPer As ClassRelMatUtiPerPRJ
Dim bAchou As Boolean
Dim iPrevUtiCalc As Integer
Dim iRealUtiCalc As Integer
Dim bRegistroValido1 As Boolean
Dim bRegistroValido2 As Boolean
Dim dPerc As Double
Dim sProjetoMask As String

On Error GoTo Erro_RelMatUtiPerPRJ_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194400

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194401
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_MATUTIPER_PRJ", lNumIntRel)
    If lErro <> SUCESSO Then gError 194402
    
    sProjeto = String(STRING_PRJ_CODIGO, 0)
    sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    
    If lNumIntDocProjeto = 0 Then
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Codigo, FilialEmpresa, Descricao, Cliente FROM Projetos ", lNumIntDocPRJ, sProjeto, iFilialEmpresa, sDescricao, lCliente)
    Else
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Codigo, FilialEmpresa, Descricao, Cliente FROM Projetos WHERE NumIntDoc = ?", lNumIntDocPRJ, sProjeto, iFilialEmpresa, sDescricao, lCliente, lNumIntDocProjeto)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 194403

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194404
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProjeto = New ClassProjetos
        
        objProjeto.sCodigo = sProjeto
        objProjeto.lNumIntDoc = lNumIntDocPRJ
        objProjeto.iFilialEmpresa = iFilialEmpresa
        objProjeto.lCliente = lCliente
        objProjeto.sDescricao = sDescricao

        lErro = CF("PRJEtapas_Le_Projeto", objProjeto, True)
        If lErro <> SUCESSO Then gError 194405
        
        Call objProjeto.Calcula_Dados_Calculados
        
        Call Retorno_Projeto_Tela2(objProjeto.sCodigo, sProjetoMask)
        
        For Each objEtapa In objProjeto.colEtapas
        
            If objEtapa.dPercentualComplet <> 0 Then
                dPerc = objEtapa.dPercentualComplet
            Else
                dPerc = objEtapa.objDadosCalc.dPercentualComplet
            End If
        
            iPrevUtiCalc = MARCADO
            iRealUtiCalc = MARCADO
                
            Set colMat = New Collection
            
            'Filtra pela etapa
            If objEtapa.lNumIntDoc = lNumIntDocEtapa Or lNumIntDocEtapa = 0 Then
            
                For Each objEtapaMat In objEtapa.colMateriaPrima
                
                    'faz o filtro por produto se existir
                    bRegistroValido1 = (objEtapaMat.sProduto >= sProdutoDe Or sProdutoDe = "") And (objEtapaMat.sProduto <= sProdutoAte Or sProdutoAte = "")
                    
                    If iPorData = MARCADO Then
                
                        'A utilização do material tem que estar dentro do período
                        bRegistroValido2 = ((dtDataDe <= objEtapaMat.dtData Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapaMat.dtData Or dtDataAte = DATA_NULA))
                             
                    Else
                        'Alguma parte da etapa (Real\Previsto) tem que estar dentro do período
                        bRegistroValido2 = ((dtDataDe <= objEtapa.dtDataInicio Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataInicio Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe <= objEtapa.dtDataFim Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataFim Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe >= objEtapa.dtDataInicio Or dtDataDe = DATA_NULA) And (dtDataAte <= objEtapa.dtDataFim Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe <= objEtapa.dtDataInicioReal Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataInicioReal Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe <= objEtapa.dtDataFimReal Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataFimReal Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe >= objEtapa.dtDataInicioReal Or dtDataDe = DATA_NULA) And (dtDataAte <= objEtapa.dtDataFimReal Or dtDataAte = DATA_NULA))
                        
                    End If
                    
                    'Filtra pelo produto
                    If bRegistroValido1 And bRegistroValido2 Then
                    
                        bAchou = False
'                        For Each objRelMatUtiPer In colMat
'                            If objEtapaMat.sProduto = objRelMatUtiPer.sProduto And objEtapaMat.dtData = objRelMatUtiPer.dtData Then
'                                bAchou = True
'                                Exit For
'                            End If
'                        Next
                        
                        If Not bAchou Then
                        
                            Set objRelMatUtiPer = New ClassRelMatUtiPerPRJ
                            
                            iSeq = iSeq + 1
                            
                            objRelMatUtiPer.iSeq = iSeq
                            objRelMatUtiPer.lCliente = objProjeto.lCliente
                            objRelMatUtiPer.lNumIntDocEtapa = objEtapa.lNumIntDoc
                            'objRelMatUtiPer.lNumIntDocEtapaPRJMat
                            objRelMatUtiPer.lNumIntDocPRJ = objProjeto.lNumIntDoc
                            objRelMatUtiPer.lNumIntRel = lNumIntRel
                            objRelMatUtiPer.sDescricaoEtapa = objEtapa.sDescricao
                            objRelMatUtiPer.sDescricaoPRJ = objProjeto.sDescricao
                            objRelMatUtiPer.sDescricaoProd = objEtapaMat.sDescricao
                            objRelMatUtiPer.sEtapa = objEtapa.sCodigo
                            objRelMatUtiPer.sProduto = objEtapaMat.sProduto
                            objRelMatUtiPer.sProjeto = objProjeto.sCodigo
                            objRelMatUtiPer.sUM = objEtapaMat.sUM
                            objRelMatUtiPer.dtData = objEtapaMat.dtData
                            objRelMatUtiPer.sObs = objEtapaMat.sObservacao
                        
                            colMat.Add objRelMatUtiPer
                            
                        End If
                        
                        Select Case objEtapaMat.iTipo
                        
                            Case INDICE_INF_PREV
                                iPrevUtiCalc = DESMARCADO
                                objRelMatUtiPer.dCustoPrevInf = objRelMatUtiPer.dCustoPrevInf + objEtapaMat.dCusto '(objEtapaMat.dQuantidade * objEtapaMat.dCusto)
                                objRelMatUtiPer.dQuantidadePrevInf = objRelMatUtiPer.dQuantidadePrevInf + objEtapaMat.dQuantidade
                            
                            Case INDICE_CALC_PREV
                                objRelMatUtiPer.dCustoPrevCalc = objRelMatUtiPer.dCustoPrevCalc + objEtapaMat.dCusto '(objEtapaMat.dQuantidade * objEtapaMat.dCusto)
                                objRelMatUtiPer.dQuantidadePrevCalc = objRelMatUtiPer.dQuantidadePrevCalc + objEtapaMat.dQuantidade
                            
                            Case INDICE_INF_REAL
                                iRealUtiCalc = DESMARCADO
                                objRelMatUtiPer.dCustoRealInf = objRelMatUtiPer.dCustoRealInf + objEtapaMat.dCusto '(objEtapaMat.dQuantidade * objEtapaMat.dCusto)
                                objRelMatUtiPer.dQuantidadeRealInf = objRelMatUtiPer.dQuantidadeRealInf + objEtapaMat.dQuantidade
                            
                            Case INDICE_CALC_REAL
                                objRelMatUtiPer.dCustoRealCalc = objRelMatUtiPer.dCustoRealCalc + objEtapaMat.dCusto '(objEtapaMat.dQuantidade * objEtapaMat.dCusto)
                                objRelMatUtiPer.dQuantidadeRealCalc = objRelMatUtiPer.dQuantidadeRealCalc + objEtapaMat.dQuantidade

                        End Select
                    
                    End If
                
                Next
            
            End If
            
            For Each objRelMatUtiPer In colMat
    
                lErro = Comando_Executar(alComando(2), "INSERT INTO RelMatUtiPerPRJ( NumIntRel, Seq, Produto, DescricaoProd, " & _
                        "CustoPrevInf, QuantidadePrevInf, CustoPrevCalc, QuantidadePrevCalc, CustoRealInf, " & _
                        "QuantidadeRealInf, CustoRealCalc, QuantidadeRealCalc, UM, NumIntDocEtapaPRJMat, " & _
                        "NumIntDocPRJ, Projeto, DescricaoPRJ, Cliente, NumIntDocEtapa, " & _
                        "Etapa, DescricaoEtapa,PrevUtiCalc,RealUtiCalc, Data, Perc, Obs) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelMatUtiPer.lNumIntRel, objRelMatUtiPer.iSeq, objRelMatUtiPer.sProduto, objRelMatUtiPer.sDescricaoProd, _
                        objRelMatUtiPer.dCustoPrevInf, objRelMatUtiPer.dQuantidadePrevInf, objRelMatUtiPer.dCustoPrevCalc, objRelMatUtiPer.dQuantidadePrevCalc, objRelMatUtiPer.dCustoRealInf, _
                        objRelMatUtiPer.dQuantidadeRealInf, objRelMatUtiPer.dCustoRealCalc, objRelMatUtiPer.dQuantidadeRealCalc, objRelMatUtiPer.sUM, objRelMatUtiPer.lNumIntDocEtapaPRJMat, _
                        objRelMatUtiPer.lNumIntDocPRJ, sProjetoMask, objRelMatUtiPer.sDescricaoPRJ, objRelMatUtiPer.lCliente, objRelMatUtiPer.lNumIntDocEtapa, _
                        objRelMatUtiPer.sEtapa, objRelMatUtiPer.sDescricaoEtapa, iPrevUtiCalc, iRealUtiCalc, objRelMatUtiPer.dtData, dPerc, objRelMatUtiPer.sObs)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194406

            Next

        Next
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194407
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194408

    RelMatUtiPerPRJ_Prepara = SUCESSO

    Exit Function

Erro_RelMatUtiPerPRJ_Prepara:

    RelMatUtiPerPRJ_Prepara = gErr

    Select Case gErr
    
        Case 194400
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 194401
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 194402, 194405
        
        Case 194403, 194404, 194407
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 194407
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMATUTIPERPRJ", gErr)

        Case 194408
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194409)

    End Select
    
    Call Transacao_Rollback

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function ApontPRJ_Grava(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ApontPRJ_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194487

    lErro = CF("ApontPRJ_Grava_EmTrans", objApontPRJ)
    If lErro <> SUCESSO Then gError 194488

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194489

    ApontPRJ_Grava = SUCESSO

    Exit Function

Erro_ApontPRJ_Grava:

    ApontPRJ_Grava = gErr

    Select Case gErr

        Case 194487
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 194488

        Case 194489
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194490)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function ApontPRJ_Grava_EmTrans(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim alComando(0 To 10) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_ApontPRJ_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194487
    Next

    'Le a tabelaApontPRJ
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ApontPRJ WHERE Codigo= ? ", _
                0, lNumIntDoc, objApontPRJ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 194488

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194489

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then
    
        objApontPRJ.lNumIntDoc = lNumIntDoc
    
        lErro = ApontMaqPRJ_Exclui(alComando(2), alComando(3), objApontPRJ)
        If lErro <> SUCESSO Then gError 194490

        lErro = ApontMOPRJ_Exclui(alComando(4), alComando(5), objApontPRJ)
        If lErro <> SUCESSO Then gError 194491

        lErro = ApontProdPRJ_Exclui(alComando(6), alComando(7), objApontPRJ)
        If lErro <> SUCESSO Then gError 194492

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 194493

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ApontPRJ SET NumIntDocPRJ= ? , NumIntDocEtapa= ? , Data= ? , Descricao= ? , " & _
                "Observacao= ? , Usuario= ? , DataRegistro= ? , HoraRegistro= ? ", alComando(0), _
                objApontPRJ.lNumIntDocPRJ, objApontPRJ.lNumIntDocEtapa, objApontPRJ.dtData, objApontPRJ.sDescricao, _
                objApontPRJ.sObservacao, gsUsuario, Date, CDbl(Time))
        If lErro <> AD_SQL_SUCESSO Then gError 194494

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_APONTPRJ", lNumIntDoc)
        If lErro <> SUCESSO Then gError 194495

        objApontPRJ.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO ApontPRJ( NumIntDoc, Codigo, NumIntDocPRJ, NumIntDocEtapa, " & _
                "Data, Descricao, Observacao, UsuarioAlt, DataUltAlt, " & _
                "HoraUltAlt) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objApontPRJ.lNumIntDoc, objApontPRJ.lCodigo, objApontPRJ.lNumIntDocPRJ, objApontPRJ.lNumIntDocEtapa, _
                objApontPRJ.dtData, objApontPRJ.sDescricao, objApontPRJ.sObservacao, gsUsuario, Date, _
                CDbl(Time))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194496

    End If
    
    lErro = ApontMaqPRJ_Insere(alComando(8), objApontPRJ)
    If lErro <> SUCESSO Then gError 194497

    lErro = ApontMOPRJ_Insere(alComando(9), objApontPRJ)
    If lErro <> SUCESSO Then gError 194498

    lErro = ApontProdPRJ_Insere(alComando(10), objApontPRJ)
    If lErro <> SUCESSO Then gError 194499

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ApontPRJ_Grava_EmTrans = SUCESSO

    Exit Function

Erro_ApontPRJ_Grava_EmTrans:

    ApontPRJ_Grava_EmTrans = gErr

    Select Case gErr

        Case 194487
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194488, 194489
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTPRJ", gErr)
            
        Case 194490 To 194492, 194495, 194496 To 194499

        Case 194493
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTPRJ", gErr)

        Case 194494
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_APONTPRJ", gErr)

        Case 194496
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_APONTPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194500)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function ApontPRJ_Exclui(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ApontPRJ_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194501

    lErro = CF("ApontPRJ_Exclui_EmTrans", objApontPRJ)
    If lErro <> SUCESSO Then gError 194502

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194503

    ApontPRJ_Exclui = SUCESSO

    Exit Function

Erro_ApontPRJ_Exclui:

    ApontPRJ_Exclui = gErr

    Select Case gErr

        Case 194501
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 194502

        Case 194503
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194504)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function ApontPRJ_Exclui_EmTrans(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_ApontPRJ_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194505
    Next

    'Le a tabelaApontPRJ
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ApontPRJ WHERE Codigo= ? ", _
                0, lNumIntDoc, objApontPRJ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 194506

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194507

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 194508
    
    objApontPRJ.lNumIntDoc = lNumIntDoc

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 194509

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ApontPRJ", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 194510

    lErro = ApontMaqPRJ_Exclui(alComando(2), alComando(3), objApontPRJ)
    If lErro <> SUCESSO Then gError 194511

    lErro = ApontMOPRJ_Exclui(alComando(4), alComando(5), objApontPRJ)
    If lErro <> SUCESSO Then gError 194512

    lErro = ApontProdPRJ_Exclui(alComando(6), alComando(7), objApontPRJ)
    If lErro <> SUCESSO Then gError 194513
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ApontPRJ_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_ApontPRJ_Exclui_EmTrans:

    ApontPRJ_Exclui_EmTrans = gErr

    Select Case gErr

        Case 194505
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194506, 194507
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTPRJ", gErr)

        Case 194508
            Call Rotina_Erro(vbOKOnly, "ERRO_APONTPRJ_NAO_CADASTRADO", gErr, objApontPRJ.lCodigo)

        Case 194509
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTPRJ", gErr)

        Case 194510
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APONTPRJ", gErr)
            
        Case 194511 To 194513

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194514)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ApontProdPRJ_Insere(ByVal lComando As Long, ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objApontProdPRJ As ClassApontProdPRJ

On Error GoTo Erro_ApontProdPRJ_Insere

    For Each objApontProdPRJ In objApontPRJ.colMateriaPrima

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_APONTPRODPRJ", lNumIntDoc)
        If lErro <> SUCESSO Then gError 194466

        objApontProdPRJ.lNumIntDoc = lNumIntDoc
        objApontProdPRJ.lNumIntDocApont = objApontPRJ.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO ApontProdPRJ( NumIntDoc, NumIntDocApont, Seq, Produto, " & _
                "UM, Qtd, Custo, Obs) VALUES (?,?,?,?,?,?,?,?)", _
                objApontProdPRJ.lNumIntDoc, objApontProdPRJ.lNumIntDocApont, objApontProdPRJ.iSeq, objApontProdPRJ.sProduto, _
                objApontProdPRJ.sUM, objApontProdPRJ.dQtd, objApontProdPRJ.dCusto, objApontProdPRJ.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194467

    Next

    ApontProdPRJ_Insere = SUCESSO

    Exit Function

Erro_ApontProdPRJ_Insere:

    ApontProdPRJ_Insere = gErr

    Select Case gErr

        Case 194466

        Case 194467
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_APONTPRODPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194468)

    End Select

    Exit Function

End Function

Public Function ApontProdPRJ_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ApontProdPRJ_Exclui

    'Le a tabelaApontProdPRJ
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM ApontProdPRJ WHERE NumIntDocApont= ? ", _
                0, lNumIntDoc, objApontPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194460

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194461

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 194462
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ApontProdPRJ", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 194463

        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194464

    Loop

    ApontProdPRJ_Exclui = SUCESSO

    Exit Function

Erro_ApontProdPRJ_Exclui:

    ApontProdPRJ_Exclui = gErr

    Select Case gErr

        Case 194460, 194461, 194464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTPRODPRJ", gErr)

        Case 194462
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTPRODPRJ", gErr)

        Case 194463
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APONTPRODPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194465)

    End Select

    Exit Function

End Function

Public Function ApontMaqPRJ_Insere(ByVal lComando As Long, ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objApontMaqPRJ As ClassApontMaqPRJ

On Error GoTo Erro_ApontMaqPRJ_Insere

    For Each objApontMaqPRJ In objApontPRJ.colMaquinas

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_APONTMAQPRJ", lNumIntDoc)
        If lErro <> SUCESSO Then gError 194469

        objApontMaqPRJ.lNumIntDoc = lNumIntDoc
        objApontMaqPRJ.lNumIntDocApont = objApontPRJ.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO ApontMaqPRJ( NumIntDoc, NumIntDocApont, Seq, CodMaq, " & _
                "Horas, Qtd, Custo, Obs) VALUES (?,?,?,?,?,?,?,?)", _
                objApontMaqPRJ.lNumIntDoc, objApontMaqPRJ.lNumIntDocApont, objApontMaqPRJ.iSeq, objApontMaqPRJ.iCodMaq, _
                objApontMaqPRJ.dHoras, objApontMaqPRJ.iQtd, objApontMaqPRJ.dCusto, objApontMaqPRJ.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194470

    Next

    ApontMaqPRJ_Insere = SUCESSO

    Exit Function

Erro_ApontMaqPRJ_Insere:

    ApontMaqPRJ_Insere = gErr

    Select Case gErr

        Case 194469

        Case 194470
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_APONTMAQPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194471)

    End Select

    Exit Function

End Function

Public Function ApontMaqPRJ_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ApontMaqPRJ_Exclui

    'Le a tabelaApontMaqPRJ
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM ApontMaqPRJ WHERE NumIntDocApont= ? ", _
                0, lNumIntDoc, objApontPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194472

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194473

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 194474
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ApontMaqPRJ", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 194475

        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194476

    Loop

    ApontMaqPRJ_Exclui = SUCESSO

    Exit Function

Erro_ApontMaqPRJ_Exclui:

    ApontMaqPRJ_Exclui = gErr

    Select Case gErr

        Case 194472, 194473, 194476
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTMAQPRJ", gErr)

        Case 194474
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTMAQPRJ", gErr)

        Case 194475
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APONTMAQPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194477)

    End Select

    Exit Function

End Function


Public Function ApontMOPRJ_Insere(ByVal lComando As Long, ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objApontMOPRJ As ClassApontMOPRJ

On Error GoTo Erro_ApontMOPRJ_Insere

    For Each objApontMOPRJ In objApontPRJ.colMaoDeObra

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_APONTMOPRJ", lNumIntDoc)
        If lErro <> SUCESSO Then gError 194478

        objApontMOPRJ.lNumIntDoc = lNumIntDoc
        objApontMOPRJ.lNumIntDocApont = objApontPRJ.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO ApontMOPRJ( NumIntDoc, NumIntDocApont, Seq, CodMO, " & _
                "Horas, Qtd, Custo, Obs) VALUES (?,?,?,?,?,?,?,?)", _
                objApontMOPRJ.lNumIntDoc, objApontMOPRJ.lNumIntDocApont, objApontMOPRJ.iSeq, objApontMOPRJ.iCodMO, _
                objApontMOPRJ.dHoras, objApontMOPRJ.iQtd, objApontMOPRJ.dCusto, objApontMOPRJ.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194479

    Next

    ApontMOPRJ_Insere = SUCESSO

    Exit Function

Erro_ApontMOPRJ_Insere:

    ApontMOPRJ_Insere = gErr

    Select Case gErr

        Case 194478

        Case 194479
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_APONTMOPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194480)

    End Select

    Exit Function

End Function

Public Function ApontMOPRJ_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ApontMOPRJ_Exclui

    'Le a tabelaApontMOPRJ
    lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc FROM ApontMOPRJ WHERE NumIntDocApont= ? ", _
                0, lNumIntDoc, objApontPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194481

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194482

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 194483
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ApontMOPRJ", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 194484

        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194485

    Loop

    ApontMOPRJ_Exclui = SUCESSO

    Exit Function

Erro_ApontMOPRJ_Exclui:

    ApontMOPRJ_Exclui = gErr

    Select Case gErr

        Case 194481, 194482, 194485
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTMOPRJ", gErr)

        Case 194483
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTMOPRJ", gErr)

        Case 194484
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APONTMOPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194486)

    End Select

    Exit Function

End Function

Function ApontPRJ_Automatico(lCodigo As Long) As Long
'retorna o número do proximo

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_ApontPRJ_Automatico

    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_APONTPRJ", "ApontPRJ", "Codigo", lNumAuto)
    If lErro <> SUCESSO Then gError 194530
    
    lCodigo = lNumAuto
    
    ApontPRJ_Automatico = SUCESSO
    
    Exit Function
    
Erro_ApontPRJ_Automatico:

    ApontPRJ_Automatico = gErr
    
    Select Case gErr

        Case 194530
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194531)

    End Select

    Exit Function
    
End Function

Public Function RelMOUtiPerPRJ_Prepara(lNumIntRel As Long, ByVal lNumIntDocProjeto As Long, ByVal lNumIntDocEtapa As Long, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iMODe As Integer, ByVal iMOAte As Integer, ByVal iPorData As Integer) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iFilialEmpresa As Integer
Dim objProjeto As ClassProjetos
Dim iSeq As Integer
Dim lCliente As Long
Dim sDescricao As String
Dim objEtapa As ClassPRJEtapas
Dim objEtapaMO As ClassPRJEtapaMO
Dim colMO As Collection
Dim objRelMOUtiPer As ClassRelMOUtiPerPRJ
Dim bAchou As Boolean
Dim iPrevUtiCalc As Integer
Dim iRealUtiCalc As Integer
Dim bRegistroValido1 As Boolean
Dim bRegistroValido2 As Boolean
Dim dPerc As Double
Dim sProjetoMask As String

On Error GoTo Erro_RelMOUtiPerPRJ_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194690

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194691
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_MOUTIPER_PRJ", lNumIntRel)
    If lErro <> SUCESSO Then gError 194692
    
    sProjeto = String(STRING_PRJ_CODIGO, 0)
    sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    
    If lNumIntDocProjeto = 0 Then
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Codigo, FilialEmpresa, Descricao, Cliente FROM Projetos ", lNumIntDocPRJ, sProjeto, iFilialEmpresa, sDescricao, lCliente)
    Else
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Codigo, FilialEmpresa, Descricao, Cliente FROM Projetos WHERE NumIntDoc = ?", lNumIntDocPRJ, sProjeto, iFilialEmpresa, sDescricao, lCliente, lNumIntDocProjeto)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 194693

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194694
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProjeto = New ClassProjetos
        
        objProjeto.sCodigo = sProjeto
        objProjeto.lNumIntDoc = lNumIntDocPRJ
        objProjeto.iFilialEmpresa = iFilialEmpresa
        objProjeto.lCliente = lCliente
        objProjeto.sDescricao = sDescricao

        lErro = CF("PRJEtapas_Le_Projeto", objProjeto, True)
        If lErro <> SUCESSO Then gError 194695
        
        Call objProjeto.Calcula_Dados_Calculados
        
        Call Retorno_Projeto_Tela2(objProjeto.sCodigo, sProjetoMask)
        
        For Each objEtapa In objProjeto.colEtapas
        
            If objEtapa.dPercentualComplet <> 0 Then
                dPerc = objEtapa.dPercentualComplet
            Else
                dPerc = objEtapa.objDadosCalc.dPercentualComplet
            End If
        
            iPrevUtiCalc = MARCADO
            iRealUtiCalc = MARCADO
                
            Set colMO = New Collection
            
            'Filtra pela etapa
            If objEtapa.lNumIntDoc = lNumIntDocEtapa Or lNumIntDocEtapa = 0 Then
            
                For Each objEtapaMO In objEtapa.colMaoDeObra
                
                    'faz o filtro por produto se existir
                    bRegistroValido1 = (objEtapaMO.iMaoDeObra >= iMODe Or iMODe = 0) And (objEtapaMO.iMaoDeObra <= iMOAte Or iMOAte = 0)
                    
                    If iPorData = MARCADO Then
                
                        'A utilização do material tem que estar dentro do período
                        bRegistroValido2 = ((dtDataDe <= objEtapaMO.dtData Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapaMO.dtData Or dtDataAte = DATA_NULA))
                             
                    Else
                        'Alguma parte da etapa (Real\Previsto) tem que estar dentro do período
                        bRegistroValido2 = ((dtDataDe <= objEtapa.dtDataInicio Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataInicio Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe <= objEtapa.dtDataFim Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataFim Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe >= objEtapa.dtDataInicio Or dtDataDe = DATA_NULA) And (dtDataAte <= objEtapa.dtDataFim Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe <= objEtapa.dtDataInicioReal Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataInicioReal Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe <= objEtapa.dtDataFimReal Or dtDataDe = DATA_NULA) And (dtDataAte >= objEtapa.dtDataFimReal Or dtDataAte = DATA_NULA)) Or _
                                        ((dtDataDe >= objEtapa.dtDataInicioReal Or dtDataDe = DATA_NULA) And (dtDataAte <= objEtapa.dtDataFimReal Or dtDataAte = DATA_NULA))
                        
                    End If
                    
                    'Filtra pelo produto
                    If bRegistroValido1 And bRegistroValido2 Then
                    
                        bAchou = False
'                        For Each objRelMOUtiPer In colMO
'                            If objEtapaMO.sProduto = objRelMOUtiPer.sProduto And objEtapaMO.dtData = objRelMOUtiPer.dtData Then
'                                bAchou = True
'                                Exit For
'                            End If
'                        Next
                        
                        If Not bAchou Then
                        
                            Set objRelMOUtiPer = New ClassRelMOUtiPerPRJ
                            
                            iSeq = iSeq + 1
                            
                            objRelMOUtiPer.iSeq = iSeq
                            objRelMOUtiPer.lCliente = objProjeto.lCliente
                            objRelMOUtiPer.lNumIntDocEtapa = objEtapa.lNumIntDoc
                            'objRelMOUtiPer.lNumIntDocEtapaPRJMat
                            objRelMOUtiPer.lNumIntDocPRJ = objProjeto.lNumIntDoc
                            objRelMOUtiPer.lNumIntRel = lNumIntRel
                            objRelMOUtiPer.sDescricaoEtapa = objEtapa.sDescricao
                            objRelMOUtiPer.sDescricaoPRJ = objProjeto.sDescricao
                            objRelMOUtiPer.sDescricaoMO = objEtapaMO.sDescricao
                            objRelMOUtiPer.sEtapa = objEtapa.sCodigo
                            objRelMOUtiPer.iCodMO = objEtapaMO.iMaoDeObra
                            objRelMOUtiPer.sProjeto = objProjeto.sCodigo
                            objRelMOUtiPer.dtData = objEtapaMO.dtData
                            objRelMOUtiPer.sObs = objEtapaMO.sObservacao
                            
                            colMO.Add objRelMOUtiPer
                            
                        End If
                        
                        Select Case objEtapaMO.iTipo
                        
                            Case INDICE_INF_PREV
                                iPrevUtiCalc = DESMARCADO
                                objRelMOUtiPer.dCustoPrevInf = objRelMOUtiPer.dCustoPrevInf + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                                objRelMOUtiPer.dQuantidadePrevInf = objRelMOUtiPer.dQuantidadePrevInf + objEtapaMO.dHoras
                            
                            Case INDICE_CALC_PREV
                                objRelMOUtiPer.dCustoPrevCalc = objRelMOUtiPer.dCustoPrevCalc + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                                objRelMOUtiPer.dQuantidadePrevCalc = objRelMOUtiPer.dQuantidadePrevCalc + objEtapaMO.dHoras
                            
                            Case INDICE_INF_REAL
                                iRealUtiCalc = DESMARCADO
                                objRelMOUtiPer.dCustoRealInf = objRelMOUtiPer.dCustoRealInf + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                                objRelMOUtiPer.dQuantidadeRealInf = objRelMOUtiPer.dQuantidadeRealInf + objEtapaMO.dHoras
                            
                            Case INDICE_CALC_REAL
                                objRelMOUtiPer.dCustoRealCalc = objRelMOUtiPer.dCustoRealCalc + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                                objRelMOUtiPer.dQuantidadeRealCalc = objRelMOUtiPer.dQuantidadeRealCalc + objEtapaMO.dHoras

                        End Select
                    
                    End If
                
                Next
            
            End If
            
            For Each objRelMOUtiPer In colMO
    
                lErro = Comando_Executar(alComando(2), "INSERT INTO RelMOUtiPerPRJ( NumIntRel, Seq, CodMO, DescricaoMO, " & _
                        "CustoPrevInf, QuantidadePrevInf, CustoPrevCalc, QuantidadePrevCalc, CustoRealInf, " & _
                        "QuantidadeRealInf, CustoRealCalc, QuantidadeRealCalc,  NumIntDocEtapaPRJMO, " & _
                        "NumIntDocPRJ, Projeto, DescricaoPRJ, Cliente, NumIntDocEtapa, " & _
                        "Etapa, DescricaoEtapa,PrevUtiCalc,RealUtiCalc, Data, Perc, Obs) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelMOUtiPer.lNumIntRel, objRelMOUtiPer.iSeq, objRelMOUtiPer.iCodMO, objRelMOUtiPer.sDescricaoMO, _
                        objRelMOUtiPer.dCustoPrevInf, objRelMOUtiPer.dQuantidadePrevInf, objRelMOUtiPer.dCustoPrevCalc, objRelMOUtiPer.dQuantidadePrevCalc, objRelMOUtiPer.dCustoRealInf, _
                        objRelMOUtiPer.dQuantidadeRealInf, objRelMOUtiPer.dCustoRealCalc, objRelMOUtiPer.dQuantidadeRealCalc, objRelMOUtiPer.lNumIntDocEtapaPRJMO, _
                        objRelMOUtiPer.lNumIntDocPRJ, sProjetoMask, objRelMOUtiPer.sDescricaoPRJ, objRelMOUtiPer.lCliente, objRelMOUtiPer.lNumIntDocEtapa, _
                        objRelMOUtiPer.sEtapa, objRelMOUtiPer.sDescricaoEtapa, iPrevUtiCalc, iRealUtiCalc, objRelMOUtiPer.dtData, dPerc, objRelMOUtiPer.sObs)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194696

            Next

        Next
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194697
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194698

    RelMOUtiPerPRJ_Prepara = SUCESSO

    Exit Function

Erro_RelMOUtiPerPRJ_Prepara:

    RelMOUtiPerPRJ_Prepara = gErr

    Select Case gErr
    
        Case 194690
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 194691
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 194692, 194695
        
        Case 194693, 194694, 194697
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 194696
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMOUTIPERPRJ", gErr)

        Case 194698
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194699)

    End Select
    
    Call Transacao_Rollback

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function RelAcompPRJ_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal sProjetoDe As String, ByVal sProjetoAte As String) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim objProjeto As ClassProjetos
Dim iSeq As Integer
Dim iSeq2 As Integer
Dim objEtapa As ClassPRJEtapas
Dim objEtapaMO As ClassPRJEtapaMO
Dim objEtapaMaq As ClassPRJEtapaMaquinas
Dim objEtapaMP As ClassPRJEtapaMateriais
Dim dPercProjeto As Double
Dim dPercEtapa As Double
Dim dCustoPrevInf As Double
Dim dCustoPrevCalc As Double
Dim dCustoRealInf As Double
Dim dCustoRealCalc As Double
Dim colAcomp As New Collection
Dim objRelAcomp As ClassRelAcompPRJ
Dim objRelAcompDet As ClassRelAcompPRJDet
Dim bAchou As Boolean
Dim iPrevUtiCalc As Integer
Dim iRealUtiCalc As Integer
Dim sProjetoMask As String
Dim sProdutoMask As String

On Error GoTo Erro_RelAcompPRJ_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 194700

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194701
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_ACOMP_PRJ", lNumIntRel)
    If lErro <> SUCESSO Then gError 194702
    
    iSeq = 0
    
    sProjeto = String(STRING_PRJ_CODIGO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Codigo FROM Projetos WHERE (FilialEmpresa = ? OR FilialEmpresa = ?) AND (Codigo >= ? OR ?= '') AND (Codigo <= ? OR ?= '') ", lNumIntDocPRJ, sProjeto, iFilialEmpresa, 0, sProjetoDe, sProjetoDe, sProjetoAte, sProjetoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 194703

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194704
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objProjeto = New ClassProjetos
        
        objProjeto.sCodigo = sProjeto
        objProjeto.lNumIntDoc = lNumIntDocPRJ
        objProjeto.iFilialEmpresa = iFilialEmpresa

        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194705

        lErro = CF("PRJEtapas_Le_Projeto", objProjeto, True)
        If lErro <> SUCESSO Then gError 194706
        
        Call objProjeto.Calcula_Dados_Calculados
        
        Call Retorno_Projeto_Tela2(objProjeto.sCodigo, sProjetoMask)
        
        If objProjeto.dPercentualComplet <> 0 Then
            dPercProjeto = objProjeto.dPercentualComplet
        Else
            dPercProjeto = objProjeto.objDadosCalc.dPercentualComplet
        End If
            
        For Each objEtapa In objProjeto.colEtapas
        
            Set objRelAcomp = New ClassRelAcompPRJ
            
            iSeq = iSeq + 1
            iSeq2 = 0
            
            If objEtapa.dPercentualComplet <> 0 Then
                dPercEtapa = objEtapa.dPercentualComplet
            Else
                dPercEtapa = objEtapa.objDadosCalc.dPercentualComplet
            End If
            
            objRelAcomp.dPercCompEtapa = dPercEtapa
            objRelAcomp.dPercCompPRJ = dPercProjeto
            objRelAcomp.dtDataFimPrevEtapa = IIf(objEtapa.dtDataFim = DATA_NULA, objEtapa.objDadosCalc.dtDataFimPrev, objEtapa.dtDataFim)
            objRelAcomp.dtDataFimPrevPRJ = IIf(objProjeto.dtDataFim = DATA_NULA, objProjeto.objDadosCalc.dtDataFimPrev, objProjeto.dtDataFim)
            objRelAcomp.dtDataFimRealEtapa = IIf(objEtapa.dtDataFimReal = DATA_NULA, objEtapa.objDadosCalc.dtDataFimReal, objEtapa.dtDataFimReal)
            objRelAcomp.dtDataFimRealPRJ = IIf(objProjeto.dtDataFimReal = DATA_NULA, objProjeto.objDadosCalc.dtDataFimReal, objProjeto.dtDataFimReal)
            objRelAcomp.dtDataIniPrevEtapa = IIf(objEtapa.dtDataInicio = DATA_NULA, objEtapa.objDadosCalc.dtDataIniPrev, objEtapa.dtDataInicio)
            objRelAcomp.dtDataIniPrevPRJ = IIf(objProjeto.dtDataInicio = DATA_NULA, objProjeto.objDadosCalc.dtDataIniPrev, objProjeto.dtDataInicio)
            objRelAcomp.dtDataIniRealEtapa = IIf(objEtapa.dtDataInicioReal = DATA_NULA, objEtapa.objDadosCalc.dtDataIniReal, objEtapa.dtDataInicioReal)
            objRelAcomp.dtDataIniRealPRJ = IIf(objProjeto.dtDataInicioReal = DATA_NULA, objProjeto.objDadosCalc.dtDataIniReal, objProjeto.dtDataInicioReal)
            objRelAcomp.iSeq = iSeq
            objRelAcomp.lCliente = objProjeto.lCliente
            objRelAcomp.lNumIntDocEtapa = objEtapa.lNumIntDoc
            objRelAcomp.lNumIntDocPRJ = objEtapa.lNumIntDocPRJ
            objRelAcomp.lNumIntRel = lNumIntRel
            objRelAcomp.sDescricaoEtapa = objEtapa.sDescricao
            objRelAcomp.sDescricaoPRJ = objProjeto.sDescricao
            objRelAcomp.sEtapa = objEtapa.sCodigo
            objRelAcomp.sProjeto = objProjeto.sCodigo

            dCustoPrevInf = 0
            dCustoPrevCalc = 0
            dCustoRealInf = 0
            dCustoRealCalc = 0

            iPrevUtiCalc = MARCADO
            iRealUtiCalc = MARCADO
            
            For Each objEtapaMP In objEtapa.colMateriaPrima
                                        
                Select Case objEtapaMP.iTipo
                
                    Case INDICE_INF_PREV
                        iPrevUtiCalc = DESMARCADO
                        dCustoPrevInf = dCustoPrevInf + objEtapaMP.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_CALC_PREV
                        dCustoPrevCalc = dCustoPrevCalc + objEtapaMP.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_INF_REAL
                        iRealUtiCalc = DESMARCADO
                        dCustoRealInf = dCustoRealInf + objEtapaMP.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_CALC_REAL
                        dCustoRealCalc = dCustoRealCalc + objEtapaMP.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)

                End Select
            
            Next
            
            For Each objEtapaMP In objEtapa.colMateriaPrima
            
                Call Mascara_RetornaProdutoTela(objEtapaMP.sProduto, sProdutoMask)
            
                bAchou = False
                For Each objRelAcompDet In objRelAcomp.colDet
                    If objRelAcompDet.sTipoTexto = "Produtos" And objRelAcompDet.iSeqAcomp = objRelAcomp.iSeq And objRelAcompDet.sItem = sProdutoMask And objRelAcompDet.sObs = objEtapaMP.sObservacao Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    Set objRelAcompDet = New ClassRelAcompPRJDet
                
                    objRelAcompDet.iSeqAcomp = objRelAcomp.iSeq
                    objRelAcompDet.iTipo = 1
                    objRelAcompDet.lNumIntRel = objRelAcomp.lNumIntRel
                    objRelAcompDet.sDescricao = objEtapaMP.sDescricao
                    
                    objRelAcompDet.sItem = sProdutoMask
                    objRelAcompDet.sTipoTexto = "Produtos"
                    objRelAcompDet.sUM = objEtapaMP.sUM
                    objRelAcompDet.sObs = objEtapaMP.sObservacao
                End If
                
                If (iPrevUtiCalc = MARCADO And objEtapaMP.iTipo = INDICE_CALC_PREV) Or (iPrevUtiCalc = DESMARCADO And objEtapaMP.iTipo = INDICE_INF_PREV) Or _
                   (iRealUtiCalc = MARCADO And objEtapaMP.iTipo = INDICE_CALC_REAL) Or (iRealUtiCalc = DESMARCADO And objEtapaMP.iTipo = INDICE_INF_REAL) Then
                
                    If objEtapaMP.iTipo = INDICE_CALC_PREV Or objEtapaMP.iTipo = INDICE_INF_PREV Then
                        objRelAcompDet.dCustoPrev = objRelAcompDet.dCustoPrev + objEtapaMP.dCusto
                        objRelAcompDet.dQtdPrev = objRelAcompDet.dQtdPrev + objEtapaMP.dQuantidade
                    End If
                    
                    If objEtapaMP.iTipo = INDICE_CALC_REAL Or objEtapaMP.iTipo = INDICE_INF_REAL Then
                        objRelAcompDet.dCustoReal = objRelAcompDet.dCustoReal + objEtapaMP.dCusto
                        objRelAcompDet.dQtdReal = objRelAcompDet.dQtdReal + objEtapaMP.dQuantidade
                    End If
                
                    If Not bAchou Then
                        iSeq2 = iSeq2 + 1
                        objRelAcompDet.iSeq = iSeq2
                        
                        objRelAcomp.colDet.Add objRelAcompDet
                    End If
                End If
            
            Next
            
            objRelAcomp.dCustoPrev = objRelAcomp.dCustoPrev + IIf(iPrevUtiCalc = MARCADO, dCustoPrevCalc, dCustoPrevInf)
            objRelAcomp.dCustoReal = objRelAcomp.dCustoReal + IIf(iRealUtiCalc = MARCADO, dCustoRealCalc, dCustoPrevInf)

            dCustoPrevInf = 0
            dCustoPrevCalc = 0
            dCustoRealInf = 0
            dCustoRealCalc = 0

            iPrevUtiCalc = MARCADO
            iRealUtiCalc = MARCADO
            
            For Each objEtapaMO In objEtapa.colMaoDeObra
                                        
                Select Case objEtapaMO.iTipo
                
                    Case INDICE_INF_PREV
                        iPrevUtiCalc = DESMARCADO
                        dCustoPrevInf = dCustoPrevInf + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_CALC_PREV
                        dCustoPrevCalc = dCustoPrevCalc + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_INF_REAL
                        iRealUtiCalc = DESMARCADO
                        dCustoRealInf = dCustoRealInf + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_CALC_REAL
                        dCustoRealCalc = dCustoRealCalc + objEtapaMO.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)

                End Select
            
            Next
            
            For Each objEtapaMO In objEtapa.colMaoDeObra
                
                bAchou = False
                For Each objRelAcompDet In objRelAcomp.colDet
                    If objRelAcompDet.sTipoTexto = "Mão de Obra" And objRelAcompDet.iSeqAcomp = objRelAcomp.iSeq And objRelAcompDet.sItem = CStr(objEtapaMO.iMaoDeObra) And objRelAcompDet.sObs = objEtapaMO.sObservacao Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    Set objRelAcompDet = New ClassRelAcompPRJDet
                
                    objRelAcompDet.iSeqAcomp = objRelAcomp.iSeq
                    objRelAcompDet.iTipo = 2
                    objRelAcompDet.lNumIntRel = objRelAcomp.lNumIntRel
                    objRelAcompDet.sDescricao = objEtapaMO.sDescricao
                    objRelAcompDet.sItem = CStr(objEtapaMO.iMaoDeObra)
                    objRelAcompDet.sTipoTexto = "Mão de Obra"
                    objRelAcompDet.sUM = "Horas"
                    objRelAcompDet.sObs = objEtapaMO.sObservacao
                End If
                
                If (iPrevUtiCalc = MARCADO And objEtapaMO.iTipo = INDICE_CALC_PREV) Or (iPrevUtiCalc = DESMARCADO And objEtapaMO.iTipo = INDICE_INF_PREV) Or _
                   (iRealUtiCalc = MARCADO And objEtapaMO.iTipo = INDICE_CALC_REAL) Or (iRealUtiCalc = DESMARCADO And objEtapaMO.iTipo = INDICE_INF_REAL) Then
                
                    If objEtapaMO.iTipo = INDICE_CALC_PREV Or objEtapaMO.iTipo = INDICE_INF_PREV Then
                        objRelAcompDet.dCustoPrev = objRelAcompDet.dCustoPrev + objEtapaMO.dCusto
                        objRelAcompDet.dQtdPrev = objRelAcompDet.dQtdPrev + objEtapaMO.dHoras
                    End If
                    
                    If objEtapaMO.iTipo = INDICE_CALC_REAL Or objEtapaMO.iTipo = INDICE_INF_REAL Then
                        objRelAcompDet.dCustoReal = objRelAcompDet.dCustoReal + objEtapaMO.dCusto
                        objRelAcompDet.dQtdReal = objRelAcompDet.dQtdReal + objEtapaMO.dHoras
                    End If
                
                    If Not bAchou Then
                        iSeq2 = iSeq2 + 1
                        objRelAcompDet.iSeq = iSeq2
                        
                        objRelAcomp.colDet.Add objRelAcompDet
                    End If
                End If
            
            Next
            
            objRelAcomp.dCustoPrev = objRelAcomp.dCustoPrev + IIf(iPrevUtiCalc = MARCADO, dCustoPrevCalc, dCustoPrevInf)
            objRelAcomp.dCustoReal = objRelAcomp.dCustoReal + IIf(iRealUtiCalc = MARCADO, dCustoRealCalc, dCustoPrevInf)

            dCustoPrevInf = 0
            dCustoPrevCalc = 0
            dCustoRealInf = 0
            dCustoRealCalc = 0

            iPrevUtiCalc = MARCADO
            iRealUtiCalc = MARCADO
            
            For Each objEtapaMaq In objEtapa.colMaquinas
                                        
                Select Case objEtapaMaq.iTipo
                
                    Case INDICE_INF_PREV
                        iPrevUtiCalc = DESMARCADO
                        dCustoPrevInf = dCustoPrevInf + objEtapaMaq.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_CALC_PREV
                        dCustoPrevCalc = dCustoPrevCalc + objEtapaMaq.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_INF_REAL
                        iRealUtiCalc = DESMARCADO
                        dCustoRealInf = dCustoRealInf + objEtapaMaq.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)
                    
                    Case INDICE_CALC_REAL
                        dCustoRealCalc = dCustoRealCalc + objEtapaMaq.dCusto '(objEtapaMO.dQuantidade * objEtapaMO.dCusto)

                End Select
            
            Next
            
            For Each objEtapaMaq In objEtapa.colMaquinas
                
                bAchou = False
                For Each objRelAcompDet In objRelAcomp.colDet
                    If objRelAcompDet.sTipoTexto = "Máquinas" And objRelAcompDet.iSeqAcomp = objRelAcomp.iSeq And objRelAcompDet.sItem = CStr(objEtapaMaq.iMaquina) And objRelAcompDet.sObs = objEtapaMaq.sObservacao Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    Set objRelAcompDet = New ClassRelAcompPRJDet
                
                    objRelAcompDet.iSeqAcomp = objRelAcomp.iSeq
                    objRelAcompDet.iTipo = 3
                    objRelAcompDet.lNumIntRel = objRelAcomp.lNumIntRel
                    objRelAcompDet.sDescricao = objEtapaMaq.sDescricao
                    objRelAcompDet.sItem = CStr(objEtapaMaq.iMaquina)
                    objRelAcompDet.sTipoTexto = "Máquinas"
                    objRelAcompDet.sUM = "Horas"
                    objRelAcompDet.sObs = objEtapaMaq.sObservacao
                End If
                
                If (iPrevUtiCalc = MARCADO And objEtapaMaq.iTipo = INDICE_CALC_PREV) Or (iPrevUtiCalc = DESMARCADO And objEtapaMaq.iTipo = INDICE_INF_PREV) Or _
                   (iRealUtiCalc = MARCADO And objEtapaMaq.iTipo = INDICE_CALC_REAL) Or (iRealUtiCalc = DESMARCADO And objEtapaMaq.iTipo = INDICE_INF_REAL) Then
                
                    If objEtapaMaq.iTipo = INDICE_CALC_PREV Or objEtapaMaq.iTipo = INDICE_INF_PREV Then
                        objRelAcompDet.dCustoPrev = objRelAcompDet.dCustoPrev + objEtapaMaq.dCusto
                        objRelAcompDet.dQtdPrev = objRelAcompDet.dQtdPrev + objEtapaMaq.dHoras
                    End If
                    
                    If objEtapaMaq.iTipo = INDICE_CALC_REAL Or objEtapaMaq.iTipo = INDICE_INF_REAL Then
                        objRelAcompDet.dCustoReal = objRelAcompDet.dCustoReal + objEtapaMaq.dCusto
                        objRelAcompDet.dQtdReal = objRelAcompDet.dQtdReal + objEtapaMaq.dHoras
                    End If
                
                    If Not bAchou Then
                        iSeq2 = iSeq2 + 1
                        objRelAcompDet.iSeq = iSeq2
                        
                        objRelAcomp.colDet.Add objRelAcompDet
                    End If
                End If
            
            Next
            
            objRelAcomp.dCustoPrev = objRelAcomp.dCustoPrev + IIf(iPrevUtiCalc = MARCADO, dCustoPrevCalc, dCustoPrevInf)
            objRelAcomp.dCustoReal = objRelAcomp.dCustoReal + IIf(iRealUtiCalc = MARCADO, dCustoRealCalc, dCustoPrevInf)
            
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelAcompPRJ( NumIntRel, Seq, NumIntDocPRJ, Projeto, " & _
                "DescricaoPRJ, Cliente, NomeCliente, NumIntDocEtapa, Etapa, " & _
                "DescricaoEtapa, CustoPrev, CustoReal, PercCompPRJ, PercCompEtapa, " & _
                "DataIniPrevEtapa, DataFimPrevEtapa, DataIniRealEtapa, DataFimRealEtapa, DataIniPrevPRJ, " & _
                "DataFimPrevPRJ, DataIniRealPRJ, DataFimRealPRJ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objRelAcomp.lNumIntRel, objRelAcomp.iSeq, objRelAcomp.lNumIntDocPRJ, sProjetoMask, _
                objRelAcomp.sDescricaoPRJ, objRelAcomp.lCliente, objRelAcomp.sNomeCliente, objRelAcomp.lNumIntDocEtapa, objRelAcomp.sEtapa, _
                objRelAcomp.sDescricaoEtapa, objRelAcomp.dCustoPrev, objRelAcomp.dCustoReal, objRelAcomp.dPercCompPRJ, objRelAcomp.dPercCompEtapa, _
                objRelAcomp.dtDataIniPrevEtapa, objRelAcomp.dtDataFimPrevEtapa, objRelAcomp.dtDataIniRealEtapa, objRelAcomp.dtDataFimRealEtapa, objRelAcomp.dtDataIniPrevPRJ, _
                objRelAcomp.dtDataFimPrevPRJ, objRelAcomp.dtDataIniRealPRJ, objRelAcomp.dtDataFimRealPRJ)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194707
            
            For Each objRelAcompDet In objRelAcomp.colDet
            
                lErro = Comando_Executar(alComando(3), "INSERT INTO RelAcompPRJDet( NumIntRel, SeqAcomp, Seq, Tipo, " & _
                        "TipoTexto, Item, Descricao, CustoPrev, CustoReal, " & _
                        "UM, QtdPrev, QtdReal, Obs) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelAcompDet.lNumIntRel, objRelAcompDet.iSeqAcomp, objRelAcompDet.iSeq, objRelAcompDet.iTipo, _
                        objRelAcompDet.sTipoTexto, objRelAcompDet.sItem, objRelAcompDet.sDescricao, objRelAcompDet.dCustoPrev, objRelAcompDet.dCustoReal, _
                        objRelAcompDet.sUM, objRelAcompDet.dQtdPrev, objRelAcompDet.dQtdReal, objRelAcompDet.sObs)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194711
            
            Next

        Next
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194708
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 194709

    RelAcompPRJ_Prepara = SUCESSO

    Exit Function

Erro_RelAcompPRJ_Prepara:

    RelAcompPRJ_Prepara = gErr

    Select Case gErr
    
        Case 194700
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 194701
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 194702, 194705, 194706
        
        Case 194703, 194704, 194708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 194707, 194711
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELACOMPPRJ", gErr)

        Case 194709
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194710)

    End Select
    
    Call Transacao_Rollback

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Projetos_Grava_Customizado(ByVal objProjetos As ClassProjetos) As Long
    Projetos_Grava_Customizado = SUCESSO
End Function

Public Function PRJRelsEmExcel_Gera(ByVal objFluxo As ClassFluxoPRJ) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCRFATConfig As ClassCRFATConfig

On Error GoTo Erro_PRJRelsEmExcel_Gera

    Set objCRFATConfig = New ClassCRFATConfig
    objCRFATConfig.iFilialEmpresa = EMPRESA_TODA
    objCRFATConfig.iTipo = 0
    objCRFATConfig.sCodigo = "PRJ_EXPORTA_EXCEL_DIR"
    objCRFATConfig.sConteudo = objFluxo.sDiretorio
    objCRFATConfig.sDescricao = "Diretório de Exportação dos relatórios de projeto em excel"
   
    lErro = CF("CRFATConfig_Grava2", objCRFATConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Set objCRFATConfig = New ClassCRFATConfig
    objCRFATConfig.iFilialEmpresa = EMPRESA_TODA
    objCRFATConfig.iTipo = 0
    objCRFATConfig.sCodigo = "PRJ_EXPORTA_EXCEL_MODELO" & CStr(objFluxo.iTipo)
    objCRFATConfig.sConteudo = objFluxo.sModelo
    objCRFATConfig.sDescricao = "Modelo usado para relatórios de projeto em excel"
   
    lErro = CF("CRFATConfig_Grava2", objCRFATConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Abertura de Comando
    For iIndice = 0 To 6
        objFluxo.alComando(iIndice) = Comando_Abrir()
        If objFluxo.alComando(iIndice) = 0 Then gError 213117
    Next
   
    Select Case objFluxo.iTipo
    
        Case 1
            lErro = PRJRelsEmExcel_Gera_Tipo1(objFluxo)
        
        Case Else
            gError 213118
            
    End Select
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    'Fecha Comando
    For iIndice = 0 To 25
        Call Comando_Fechar(objFluxo.alComando(iIndice))
    Next
    
    PRJRelsEmExcel_Gera = SUCESSO

    Exit Function

Erro_PRJRelsEmExcel_Gera:

    PRJRelsEmExcel_Gera = gErr

    Select Case gErr

        Case 213117
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 213118 'ERRO_TIPO_CAMPO_INVALIDO
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213119)

    End Select

    'Fecha Comando
    For iIndice = 0 To 25
        Call Comando_Fechar(objFluxo.alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function PRJRelsEmExcel_Gera_Tipo1(ByVal objFluxo As ClassFluxoPRJ) As Long

Dim lErro As Long
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde
Dim lCount As Long, tMov As typeMovContaCorrente
Dim objPastaTrabalho As Object 'Excel.Workbook
Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objPlanilhaRange As Object 'Excel.Range
Dim objTabelaDinamica As Object 'Excel.PivotTable
Dim sTipoDesc As String, iCredito As Integer
Dim objItem As ClassItensPRJCR, colItens As Collection
Dim lCol As Long, lLin As Long, lCliente As Long
Dim objProjeto As ClassProjetos, objEtapa As ClassPRJEtapas
Dim objCliente As ClassCliente, objNatureza As ClassNatMovCta
Dim objCcl As ClassCcl, bAchou As Boolean
Dim colCli As New Collection, colCcl As New Collection
Dim colNat As New Collection, colPrj As New Collection
Dim iTipoPRJ As Integer, lNumIntDoc As Long, dValor As Double
Dim sNatureza As String, sCcl As String, sSigla As String, lNumTitulo As Long
Dim lCliForn As Long, iParcela As Integer
Dim sCclMask As String, sNatMask As String, sPRJMask As String
Dim sNomeCliForn As String, sNomeRedCliForn As String
Dim dtDataEmissao As Date, dtDataVencimento As Date

On Error GoTo Erro_PRJRelsEmExcel_Gera_Tipo1
  
    'Abre o excel
    lErro = CF("Excel_Abrir")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Desabilita as mensagens do Excel
    GL_objExcelSist.DisplayAlerts = False
        
    Call GL_objExcelSist.Workbooks.Open(objFluxo.sModelo)
    
    Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
    
    'Seleciona a planilha ativa na pasta de trabalho criada
    Set objPlanilhaExcel = objPastaTrabalho.Sheets(1)
    
    lCol = 0
    lLin = 0
    lLin = lLin + 1
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Data"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Valor"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Tipo de Movto"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Número do Movto"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Valor do Movto"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Detalhe do Movto"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Projeto"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Descrição do Projeto"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Etapa"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Descrição da Etapa"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Responsável Prj"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Cliente Prj"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Nome Cliente Prj"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Natureza"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Descrição Natureza"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Ccl"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Descrição Centro de Custo/Lucro"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Tipo Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Núm.Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Tipo Cli/Forn Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Cli/Forn Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Nome Red.Cli/Forn Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Nome Cli/Forn Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Data Emi.Doc"
    lCol = lCol + 1
    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
    objPlanilhaRange.Value = "Data Vencimento Doc"
    
    'Le os movimentos
    lErro = Comando_Executar(objFluxo.alComando(0), "SELECT COUNT(*) FROM MovimentosContaCorrente WHERE FilialEmpresa = ? AND DataMovimento BETWEEN ? AND ? AND Excluido = 0 ", lCount, objFluxo.iFilialEmpresa, objFluxo.dtDataDe, objFluxo.dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 213120

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(objFluxo.alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213121
   
    objFrmAguarde.iTotalItens = lCount
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)

    'Alocação de espaço no buffer
    tMov.sHistorico = String(STRING_MAXIMO, 0)
    tMov.sNumRefExterna = String(STRING_MAXIMO, 0)
    tMov.sNatureza = String(STRING_MAXIMO, 0)
    sTipoDesc = String(STRING_MAXIMO, 0)
    sCcl = String(STRING_MAXIMO, 0)
    
    'Le os vouchers
    lErro = Comando_Executar(objFluxo.alComando(1), "SELECT M.NumMovto, M.CodConta, M.Tipo, M.TipoMeioPagto, M.Numero, M.DataMovimento, M.Valor, M.Historico, M.NumRefExterna, M.NumRefInterna, M.Natureza, T.NomeReduzido, T.Credito, M.Ccl FROM MovimentosContaCorrente As M, TiposMovtoCtaCorrente As T, ContasCorrentesInternas AS C WHERE M.Tipo = T.Codigo AND M.CodConta = C.Codigo AND C.ChequePre = 0 AND M.FilialEmpresa = ? AND M.DataMovimento BETWEEN ? AND ? AND M.Excluido = 0 ORDER BY M.DataMovimento, M.NumMovto ", _
    tMov.lNumMovto, tMov.iCodConta, tMov.iTipo, tMov.iTipoMeioPagto, tMov.lNumero, tMov.dtDataMovimento, tMov.dValor, tMov.sHistorico, tMov.sNumRefExterna, tMov.lNumRefInterna, tMov.sNatureza, sTipoDesc, iCredito, sCcl, _
    objFluxo.iFilialEmpresa, objFluxo.dtDataDe, objFluxo.dtDataAte)
    If lErro <> AD_SQL_SUCESSO Then gError 213122

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(objFluxo.alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213123
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set colItens = New Collection
    
        Select Case tMov.iTipo
        
            Case MOVCCI_SAQUE, MOVCCI_DEPOSITO
            
                Set objItem = New ClassItensPRJCR
                
                objItem.sNatureza = tMov.sNatureza
                objItem.dValorOriginal = tMov.dValor
                objItem.sDescricao = sTipoDesc
                objItem.sCcl = sCcl
                objItem.dtDataEmissao = tMov.dtDataMovimento
                objItem.dtDataVencimento = tMov.dtDataMovimento
                
                If iCredito = DESMARCADO Then objItem.iNegativo = MARCADO
                
                If tMov.iTipo = MOVCCI_DEPOSITO Then
                    iTipoPRJ = PRJ_CR_TIPO_DEPOSITO
                    objItem.sTipoDoc = "Depósito"
                Else
                    iTipoPRJ = PRJ_CR_TIPO_SAQUE
                    objItem.sTipoDoc = "Saque"
                End If
                lNumIntDoc = tMov.lNumMovto
                
                lErro = PRJRelsEmExcel_Obtem_PRJ(objFluxo, iTipoPRJ, lNumIntDoc, objItem, colItens)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MOVCCI_RECEBIMENTO_TITULO, MOVCCI_EXTRATO_CARTAO_CRED, MOVCCI_BORDERO_CHEQUE_PRE, MOVCCI_DEP_DESCONTO_CHEQUE, MOVCCI_SAQ_DESCONTO_CHEQUE
            
                sNatureza = String(STRING_MAXIMO, 0)
                sCcl = String(STRING_MAXIMO, 0)
                sSigla = String(STRING_MAXIMO, 0)
                sNomeCliForn = String(STRING_MAXIMO, 0)
                sNomeRedCliForn = String(STRING_MAXIMO, 0)
            
                lErro = Comando_Executar(objFluxo.alComando(2), "SELECT P.NumParcela, T.DataEmissao, P.DataVencimento, P.NumIntTitulo, T.Natureza, T.Ccl, T.SiglaDocumento, T.NumTitulo, SUM(BP.ValorRecebido), T.Cliente, C.NomeReduzido, C.RazaoSocial FROM BaixasRec AS B, BaixasParcRec As BP, ParcelasRecTodas AS P, TitulosRecTodos As T, Clientes AS C WHERE T.NumIntDoc = P.NumIntTitulo AND B.NumMovCta = ? AND B.NumIntBaixa = BP.NumIntBaixa AND BP.NumIntParcela = P.NumIntDoc AND B.Status <> 5 AND BP.Status <> 5 AND C.Codigo = T.Cliente GROUP BY P.NumParcela, T.DataEmissao, P.DataVencimento, P.NumIntTitulo, T.Natureza, T.Ccl, T.SiglaDocumento, T.NumTitulo, T.Cliente, C.NomeReduzido, C.RazaoSocial ", _
                iParcela, dtDataEmissao, dtDataVencimento, lNumIntDoc, sNatureza, sCcl, sSigla, lNumTitulo, dValor, lCliForn, sNomeRedCliForn, sNomeCliForn, tMov.lNumMovto)
                If lErro <> AD_SQL_SUCESSO Then gError 213124
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(objFluxo.alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213125
                
                Do While lErro <> AD_SQL_SEM_DADOS
                
                    Set objItem = New ClassItensPRJCR
                    
                    objItem.sNatureza = sNatureza
                    objItem.dValorOriginal = dValor
                    objItem.sDescricao = sTipoDesc
                    objItem.sCcl = sCcl
                    objItem.sObservacao = sSigla & " " & CStr(lNumTitulo) & "-" & CStr(iParcela)
                    objItem.dtDataEmissao = dtDataEmissao
                    objItem.dtDataVencimento = dtDataVencimento
                    If iCredito = DESMARCADO Then objItem.iNegativo = MARCADO
                    
                    objItem.sTipoDoc = sSigla
                    objItem.sNumDoc = CStr(lNumTitulo) & "-" & CStr(iParcela)
                    objItem.sTipoCliForn = "Cliente"
                    objItem.sNomeCliForn = sNomeCliForn
                    objItem.sNomeRedCliForn = sNomeRedCliForn
                    objItem.sCliForn = CStr(lCliForn)

                    iTipoPRJ = PRJ_CR_TIPO_TITREC
                    
                    lErro = PRJRelsEmExcel_Obtem_PRJ(objFluxo, iTipoPRJ, lNumIntDoc, objItem, colItens)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    lErro = Comando_BuscarProximo(objFluxo.alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213126
                
                Loop
            
            Case MOVCCI_PAGTO_TITULO_POR_CHEQUE, MOVCCI_PAGTO_TITULO_POR_DINHEIRO, MOVCCI_PAGTO_TITULO_POR_BORDERO
            
                sNatureza = String(STRING_MAXIMO, 0)
                sCcl = String(STRING_MAXIMO, 0)
                sSigla = String(STRING_MAXIMO, 0)
                sNomeCliForn = String(STRING_MAXIMO, 0)
                sNomeRedCliForn = String(STRING_MAXIMO, 0)
                
                lErro = Comando_Executar(objFluxo.alComando(3), "SELECT P.NumParcela, T.DataEmissao, P.DataVencimento, P.NumIntTitulo, T.Natureza, T.Ccl, T.SiglaDocumento, T.NumTitulo, SUM(BP.ValorBaixado+BP.ValorMulta+BP.ValorJuros-BP.ValorDesconto), T.Fornecedor, F.NomeReduzido, F.RazaoSocial FROM BaixasPag AS B, BaixasParcPag As BP, ParcelasPagTodas AS P, TitulosPagTodos As T, Fornecedores AS F WHERE T.NumIntDoc = P.NumIntTitulo AND B.NumMovCta = ? AND B.NumIntBaixa = BP.NumIntBaixa AND BP.NumIntParcela = P.NumIntDoc AND B.Status <> 5 AND BP.Status <> 5 AND F.Codigo = T.Fornecedor GROUP BY P.NumParcela, T.DataEmissao, P.DataVencimento, P.NumIntTitulo, T.Natureza, T.Ccl, T.SiglaDocumento, T.NumTitulo, T.Fornecedor, F.NomeReduzido, F.RazaoSocial ", _
                iParcela, dtDataEmissao, dtDataVencimento, lNumIntDoc, sNatureza, sCcl, sSigla, lNumTitulo, dValor, lCliForn, sNomeRedCliForn, sNomeCliForn, tMov.lNumMovto)
                If lErro <> AD_SQL_SUCESSO Then gError 213127
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(objFluxo.alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213128
                
                Do While lErro <> AD_SQL_SEM_DADOS
                
                    Set objItem = New ClassItensPRJCR
                    
                    objItem.sNatureza = sNatureza
                    objItem.dValorOriginal = dValor
                    objItem.sDescricao = sTipoDesc
                    objItem.sCcl = sCcl
                    objItem.sObservacao = sSigla & " " & CStr(lNumTitulo) & "-" & CStr(iParcela)
                    If iCredito = DESMARCADO Then objItem.iNegativo = MARCADO

                    objItem.sTipoDoc = sSigla
                    objItem.sNumDoc = CStr(lNumTitulo) & "-" & CStr(iParcela)
                    objItem.sTipoCliForn = "Fornecedor"
                    objItem.sNomeCliForn = sNomeCliForn
                    objItem.sNomeRedCliForn = sNomeRedCliForn
                    objItem.sCliForn = CStr(lCliForn)
                    objItem.dtDataEmissao = dtDataEmissao
                    objItem.dtDataVencimento = dtDataVencimento
                    
                    iTipoPRJ = PRJ_CR_TIPO_TITPAG
                    
                    lErro = PRJRelsEmExcel_Obtem_PRJ(objFluxo, iTipoPRJ, lNumIntDoc, objItem, colItens)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    lErro = Comando_BuscarProximo(objFluxo.alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213129
                
                Loop
            
            Case MOVCCI_CANC_PAGTO 'NumIntDoc da BaixaParcPag CP
                'IGNORAR POIS, PARA SIMPLIFICAR, NÃO SERÁ CONSIDERADO A SAIDA DO DINHEIRO E NEM A VOLTA
            
            Case MOVCCI_RECEBTO_CANCELA, MOVCCI_DEP_DEVOLUCAO_CHEQUE, MOVCCI_SAQ_DEVOLUCAO_CHEQUE 'NumIntDoc de BaixaParcRecCanc CR
                'IGNORAR POIS, PARA SIMPLIFICAR, NÃO SERÁ CONSIDERADO A ENTRADA DO DINHEIRO E NEM A SAIDA
            
            Case MOVCCI_ENVIO_CHEQUES_TERCEIROS
                'NÃO ESTÁ SENDO TRATADO
            
            Case MOVCCI_DEVOLUCAO_CHEQUES_TERCEIROS
                'NÃO ESTÁ SENDO TRATADO
            
            Case MOVCCI_CHEQUEPRE_BACKOFFICE
                'NÃO ESTÁ CONSIDERANDO MOVIMENTAÇÕES NA CONTA DE CHEQUE PRÉ
                
            Case MOVCCI_SAIDA_TRANSFERENCIA, MOVCCI_ENTRADA_TRANSFERENCIA
                'IGNORAR
                
            Case Else
            
                Set objItem = New ClassItensPRJCR
                colItens.Add objItem
                
                objItem.sNatureza = tMov.sNatureza
                objItem.dValorOriginal = tMov.dValor
                objItem.sDescricao = sTipoDesc
                objItem.sCcl = ""
                objItem.dtDataEmissao = tMov.dtDataMovimento
                objItem.dtDataVencimento = tMov.dtDataMovimento
                
                If iCredito = DESMARCADO Then objItem.iNegativo = MARCADO
        
        End Select
                
        For Each objItem In colItens
        
            If (objFluxo.lNumIntPRJ = 0 Or objFluxo.lNumIntPRJ = objItem.lNumIntPRJ) And (objFluxo.lNumIntEtapa = 0 Or objFluxo.lNumIntEtapa = objItem.lNumIntEtapa) Then
            
                bAchou = False
                If Len(Trim(objItem.sNatureza)) > 0 Then
                    For Each objNatureza In colNat
                        If objNatureza.sCodigo = objItem.sNatureza Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Then
                    Set objNatureza = New ClassNatMovCta
                    objNatureza.sCodigo = objItem.sNatureza
                    If Len(Trim(objNatureza.sCodigo)) > 0 Then
                        lErro = CF("NatMovCta_Le", objNatureza)
                        If lErro <> SUCESSO And lErro <> 122786 Then gError ERRO_SEM_MENSAGEM
                        colNat.Add objNatureza
                    End If
                End If
                
                bAchou = False
                If Len(Trim(objItem.sCcl)) > 0 Then
                    For Each objCcl In colCcl
                        If objCcl.sCcl = objItem.sCcl Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Then
                    Set objCcl = New ClassCcl
                    objCcl.sCcl = objItem.sCcl
                    If Len(Trim(objCcl.sCcl)) > 0 Then
                        lErro = CF("Ccl_Le", objCcl)
                        If lErro <> SUCESSO And lErro <> 5599 Then gError ERRO_SEM_MENSAGEM
                        colCcl.Add objCcl
                    End If
                End If
                
                bAchou = False
                If Len(Trim(objItem.lNumIntPRJ)) > 0 Then
                    For Each objProjeto In colPrj
                        If objProjeto.lNumIntDoc = objItem.lNumIntPRJ Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Then
                    Set objProjeto = New ClassProjetos
                    objProjeto.lNumIntDoc = objItem.lNumIntPRJ
                    If objProjeto.lNumIntDoc <> 0 Then
                        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                        
                        lErro = CF("PRJEtapas_Le_Projeto", objProjeto)
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                        
                        colPrj.Add objProjeto
                    End If
                End If
                If objItem.lNumIntEtapa = 0 Then
                    Set objEtapa = New ClassPRJEtapas
                Else
                    For Each objEtapa In objProjeto.colEtapas
                        If objEtapa.lNumIntDoc = objItem.lNumIntEtapa Then
                            Exit For
                        End If
                    Next
                End If
                
                If objEtapa.lCliente <> 0 Then
                    lCliente = objEtapa.lCliente
                Else
                    lCliente = objProjeto.lCliente
                End If
                
                bAchou = False
                If lCliente > 0 Then
                    For Each objCliente In colCli
                        If objCliente.lCodigo = lCliente Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Then
                    Set objCliente = New ClassCliente
                    objCliente.lCodigo = lCliente
                    If lCliente <> 0 Then
                        lErro = CF("Cliente_Le", objCliente)
                        If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                        colCli.Add objCliente
                    End If
                End If
            
                lCol = 0
                lLin = lLin + 1
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = tMov.dtDataMovimento
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.dValorOriginal * IIf(objItem.iNegativo = MARCADO, -1, 1)
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = sTipoDesc
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = tMov.lNumMovto
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = tMov.dValor * IIf(objItem.iNegativo = MARCADO, -1, 1)
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sObservacao
                
                If Len(Trim(objProjeto.sCodigo)) > 0 Then
                    lErro = Mascara_RetornaItemTela(SEGMENTO_PROJETO, objProjeto.sCodigo, sPRJMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                Else
                    sPRJMask = ""
                End If
                
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = sPRJMask
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objProjeto.sDescricao
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objEtapa.sCodigo
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objEtapa.sDescricao
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = IIf(Len(Trim(objEtapa.sResponsavel)) > 0, objEtapa.sResponsavel, objProjeto.sResponsavel)
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = IIf(objCliente.lCodigo = 0, "", objCliente.lCodigo)
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objCliente.sRazaoSocial
                
                If Len(Trim(objItem.sNatureza)) > 0 Then
                    lErro = Mascara_RetornaItemTela(SEGMENTO_NATMOVCTA, objItem.sNatureza, sNatMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                Else
                    sNatMask = ""
                End If
                
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = sNatMask
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objNatureza.sDescricao
                
                If Len(Trim(objItem.sCcl)) > 0 Then
                    lErro = Mascara_RetornaItemTela(SEGMENTO_CCL, objItem.sCcl, sCclMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                Else
                    sCclMask = ""
                End If
                
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = sCclMask
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objCcl.sDescCcl
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sTipoDoc
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sNumDoc
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sTipoCliForn
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sCliForn
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sNomeRedCliForn
                lCol = lCol + 1
                Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                objPlanilhaRange.Value = objItem.sNomeCliForn
                lCol = lCol + 1
                If objItem.dtDataEmissao <> DATA_NULA Then
                    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                    objPlanilhaRange.Value = objItem.dtDataEmissao
                End If
                lCol = lCol + 1
                If objItem.dtDataVencimento <> DATA_NULA Then
                    Set objPlanilhaRange = objPlanilhaExcel.range(objPlanilhaExcel.Cells(lLin, lCol), objPlanilhaExcel.Cells(lLin, lCol))
                    objPlanilhaRange.Value = objItem.dtDataVencimento
                End If
            End If
        
        Next
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(objFluxo.alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213130
    
        Call objFrmAguardeTela.ProcessouItem
        
        If objFrmAguarde.iCancelar = MARCADO Then Exit Do

    Loop
    
    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 213131
    
    For Each objPlanilhaExcel In objPastaTrabalho.Sheets
        For Each objTabelaDinamica In objPlanilhaExcel.PivotTables
            Call objTabelaDinamica.PivotCache.Refresh
        Next
    Next
    
    objPastaTrabalho.SaveAS objFluxo.sDiretorio & objFluxo.sNomeArquivo & ".xls"
    GL_objExcelSist.Visible = True
       
    PRJRelsEmExcel_Gera_Tipo1 = SUCESSO

    Exit Function

Erro_PRJRelsEmExcel_Gera_Tipo1:

    PRJRelsEmExcel_Gera_Tipo1 = gErr

    Select Case gErr

        Case 213120 To 213123, 213130
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCONTACORRENTE", gErr)
    
        Case 213124 To 213126
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", gErr)
    
        Case 213127 To 213129
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 213131
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213132)
            
    End Select

    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    Call CF("Excel_Fechar")

    Exit Function

End Function

Private Function PRJRelsEmExcel_Obtem_PRJ(ByVal objFluxo As ClassFluxoPRJ, ByVal iTipoPRJ As Integer, ByVal lNumIntDoc As Long, ByVal objItem As ClassItensPRJCR, ByVal colItens As Collection) As Long

Dim lErro As Long, bAchou As Boolean
Dim objPRJInfo As ClassProjetoInfo, colPRJInfo As New Collection
Dim objPRJInfoRateio As New ClassProjetoInfo, dValorTotal As Double, dResto As Double
Dim lNumIntDocNF As Long, lNumIntDocNFPag As Long, iIndice As Integer
Dim tPRJInfo As typeProjetoInfo, dValor As Double, objItemRateio As ClassPRJCR
Dim objItemAux As ClassItensPRJCR

On Error GoTo Erro_PRJRelsEmExcel_Obtem_PRJ

    objPRJInfoRateio.iTipoOrigem = iTipoPRJ
    objPRJInfoRateio.lNumIntDocOrigem = lNumIntDoc

    Select Case iTipoPRJ
    
        Case PRJ_CR_TIPO_TITPAG
            
            lErro = Comando_Executar(objFluxo.alComando(4), "SELECT NumIntDoc FROM NFiscal WHERE ClasseDocCPR = ? AND NumIntDocCPR = ?", lNumIntDocNF, CPR_TITULO_PAGAR, lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 213133
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(objFluxo.alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213134
            
            If lErro = AD_SQL_SUCESSO Then
                Set objPRJInfo = New ClassProjetoInfo
                objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_NF
                objPRJInfo.lNumIntDocOrigem = lNumIntDocNF
                colPRJInfo.Add objPRJInfo
            Else
            
                lErro = Comando_Executar(objFluxo.alComando(5), "SELECT NumIntDoc FROM NFsPag_Todas WHERE NumIntTitPag = ?", lNumIntDocNFPag, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 213135
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(objFluxo.alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213136
                
                If lErro = AD_SQL_SUCESSO Then
                    Do While lErro <> AD_SQL_SUCESSO
                    
                        Set objPRJInfo = New ClassProjetoInfo
                        objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_NFPAG
                        objPRJInfo.lNumIntDocOrigem = lNumIntDocNFPag
                        colPRJInfo.Add objPRJInfo
                    
                        lErro = Comando_BuscarProximo(objFluxo.alComando(5))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213137
                    
                    Loop
                Else
                    Set objPRJInfo = New ClassProjetoInfo
                    objPRJInfo.iTipoOrigem = iTipoPRJ
                    objPRJInfo.lNumIntDocOrigem = lNumIntDoc
                    colPRJInfo.Add objPRJInfo
                End If
            
            End If
        
        Case PRJ_CR_TIPO_TITREC
        
            lErro = Comando_Executar(objFluxo.alComando(4), "SELECT NumIntDoc FROM NFiscal WHERE ClasseDocCPR = ? AND NumIntDocCPR = ?", lNumIntDocNF, CPR_TITULO_RECEBER, lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 213138
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(objFluxo.alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213139
            
            If lErro = AD_SQL_SUCESSO Then
                Set objPRJInfo = New ClassProjetoInfo
                objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_NF
                objPRJInfo.lNumIntDocOrigem = lNumIntDocNF
                colPRJInfo.Add objPRJInfo
            Else
                Set objPRJInfo = New ClassProjetoInfo
                objPRJInfo.iTipoOrigem = iTipoPRJ
                objPRJInfo.lNumIntDocOrigem = lNumIntDoc
                colPRJInfo.Add objPRJInfo
            End If
   
        Case Else
            Set objPRJInfo = New ClassProjetoInfo
            objPRJInfo.iTipoOrigem = iTipoPRJ
            objPRJInfo.lNumIntDocOrigem = lNumIntDoc
            colPRJInfo.Add objPRJInfo
    
    End Select
    
    For Each objPRJInfo In colPRJInfo
    
        lErro = Comando_Executar(objFluxo.alComando(6), "SELECT F.NumIntDocPRJ, F.NumIntDocEtapa, SUM( CASE WHEN F.TipoValor = 0 THEN F.Valor ELSE -F.Valor END) FROM PRJCustoReceitaReal AS F, ProjetoInfo AS I WHERE F.NumIntDocPRJInfo = I.NumIntDoc AND I.TipoOrigem = ? AND I.NumIntDocOrigem = ? GROUP BY F.NumIntDocPRJ, F.NumIntDocEtapa", _
        tPRJInfo.lNumIntDocPRJ, tPRJInfo.lNumIntDocEtapa, dValor, objPRJInfo.iTipoOrigem, objPRJInfo.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 213140
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(objFluxo.alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213141

        Do While lErro <> AD_SQL_SEM_DADOS
        
            dValorTotal = dValorTotal + dValor
            
            bAchou = False
            For Each objItemRateio In objPRJInfoRateio.colPRJCustoReceitaReal
                If objItemRateio.lNumIntDocPRJ = tPRJInfo.lNumIntDocPRJ And objItemRateio.lNumIntDocEtapa = tPRJInfo.lNumIntDocEtapa Then
                    objItemRateio.dValor = objItemRateio.dValor + dValor
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objItemRateio = New ClassPRJCR
                objItemRateio.lNumIntDocPRJ = tPRJInfo.lNumIntDocPRJ
                objItemRateio.lNumIntDocEtapa = tPRJInfo.lNumIntDocEtapa
                objItemRateio.dValor = dValor
                objPRJInfoRateio.colPRJCustoReceitaReal.Add objItemRateio
            End If
        
            lErro = Comando_BuscarProximo(objFluxo.alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213142
        
        Loop
        
    Next
    
    If objPRJInfoRateio.colPRJCustoReceitaReal.Count = 0 Or dValorTotal < DELTA_VALORMONETARIO Then
        colItens.Add objItem 'Adiciona o que veio, não tem projeto
    Else
        'Se tem projeto quebra o valor de acordo com o valor de cada projeto
        iIndice = 0
        dResto = objItem.dValorOriginal
        For Each objItemRateio In objPRJInfoRateio.colPRJCustoReceitaReal
            iIndice = iIndice + 1
            Set objItemAux = New ClassItensPRJCR
            If iIndice = objPRJInfoRateio.colPRJCustoReceitaReal.Count Then
                'Pro último coloca o que sobrar
                objItemAux.dValorOriginal = dResto
            Else
                objItemAux.dValorOriginal = Arredonda_Moeda(objItem.dValorOriginal * objItemRateio.dValor / dValorTotal)
            End If
            dResto = dResto - objItemAux.dValorOriginal
            objItemAux.lNumIntEtapa = objItemRateio.lNumIntDocEtapa
            objItemAux.lNumIntPRJ = objItemRateio.lNumIntDocPRJ
            objItemAux.lNumMovto = objItem.lNumMovto
            objItemAux.iNegativo = objItem.iNegativo
            objItemAux.sCcl = objItem.sCcl
            objItemAux.sDescricao = objItem.sDescricao
            objItemAux.sItem = objItem.sItem
            objItemAux.sNatureza = objItem.sNatureza
            objItemAux.sObservacao = objItem.sObservacao
            
            objItemAux.sCliForn = objItem.sCliForn
            objItemAux.sNumDoc = objItem.sNumDoc
            objItemAux.sTipoDoc = objItem.sTipoDoc
            objItemAux.sTipoCliForn = objItem.sTipoCliForn
            objItemAux.sNomeCliForn = objItem.sNomeCliForn
            objItemAux.sNomeRedCliForn = objItem.sNomeRedCliForn
            objItemAux.dtDataEmissao = objItem.dtDataEmissao
            objItemAux.dtDataVencimento = objItem.dtDataVencimento
            
            colItens.Add objItemAux
        Next
    End If

    PRJRelsEmExcel_Obtem_PRJ = SUCESSO

    Exit Function

Erro_PRJRelsEmExcel_Obtem_PRJ:

    PRJRelsEmExcel_Obtem_PRJ = gErr

    Select Case gErr

        Case 213133, 213134, 213138, 213139
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL4", gErr)

        Case 213135, 213136, 213137
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)
            
        Case 213140 To 213142
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213143)


    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PRJEtapaVistorias_Grava(ByVal objPRJEtapaVistorias As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJEtapaVistorias_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213613

    lErro = CF("PRJEtapaVistorias_Grava_EmTrans", objPRJEtapaVistorias)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213614

    PRJEtapaVistorias_Grava = SUCESSO

    Exit Function

Erro_PRJEtapaVistorias_Grava:

    PRJEtapaVistorias_Grava = gErr

    Select Case gErr

        Case 213613
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213614
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213615)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJEtapaVistorias_Grava_EmTrans(ByVal objPRJEtapaVistorias As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long, lCodigo As Long

On Error GoTo Erro_PRJEtapaVistorias_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213616
    Next

    'Le a tabelaPRJEtapaVistorias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapaVistorias WHERE NumIntPRJEtapa= ? AND Codigo= ? ", _
                0, lNumIntDoc, objPRJEtapaVistorias.lNumIntPRJEtapa, objPRJEtapaVistorias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213617

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213618

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 213619

        objPRJEtapaVistorias.lNumIntDoc = lNumIntDoc

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PRJEtapaVistorias SET Data= ? , DataValidade= ?, Responsavel= ? , Laudo= ? ", alComando(0), _
                objPRJEtapaVistorias.dtData, objPRJEtapaVistorias.dtDataValidade, objPRJEtapaVistorias.sResponsavel, objPRJEtapaVistorias.sLaudo)
        If lErro <> AD_SQL_SUCESSO Then gError 213620

    Else
    
        If objPRJEtapaVistorias.lCodigo <> 0 Then gError 213630
    
        'Le a tabelaPRJEtapaVistorias
        lErro = Comando_ExecutarPos(alComando(4), "SELECT ProxCodVistoria FROM PRJEtapas WHERE NumIntDoc = ?  ", _
                    0, lCodigo, objPRJEtapaVistorias.lNumIntPRJEtapa)
        If lErro <> AD_SQL_SUCESSO Then gError 213627
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213628
   
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE PRJEtapas SET ProxCodVistoria= ? ", alComando(4), lCodigo + 1)
        If lErro <> AD_SQL_SUCESSO Then gError 213629

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_INT_PROX_PRJETAPAVISTORIAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objPRJEtapaVistorias.lNumIntDoc = lNumIntDoc
        objPRJEtapaVistorias.lCodigo = lCodigo

        lErro = Comando_Executar(alComando(1), "INSERT INTO PRJEtapaVistorias(NumIntDoc, NumIntPRJEtapa, Codigo, Data, DataValidade, Responsavel, Laudo) VALUES (?,?,?,?,?,?,?)", _
                objPRJEtapaVistorias.lNumIntDoc, objPRJEtapaVistorias.lNumIntPRJEtapa, objPRJEtapaVistorias.lCodigo, objPRJEtapaVistorias.dtData, _
                objPRJEtapaVistorias.dtDataValidade, objPRJEtapaVistorias.sResponsavel, objPRJEtapaVistorias.sLaudo)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213621

    End If
    
    lErro = PRJEtapaVistorias_Grava_UltVist(objPRJEtapaVistorias.lNumIntPRJEtapa, alComando(2), alComando(3))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not (objPRJEtapaVistorias.objAnexos Is Nothing) Then
        objPRJEtapaVistorias.objAnexos.iTipoDoc = ANEXO_TIPO_VISTPRJ
        objPRJEtapaVistorias.objAnexos.lNumIntDoc = objPRJEtapaVistorias.lNumIntDoc
        lErro = CF("Anexos_Grava", objPRJEtapaVistorias.objAnexos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaVistorias_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaVistorias_Grava_EmTrans:

    PRJEtapaVistorias_Grava_EmTrans = gErr

    Select Case gErr

        Case 213616
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213617, 213618, 213627, 213628
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAVISTORIAS", gErr)

        Case 213619
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAVISTORIAS", gErr)

        Case 213620, 213629
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRJETAPAVISTORIAS", gErr)

        Case 213621
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRJETAPAVISTORIAS", gErr)
            
        Case 213630
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAVISTORIAS_NAO_CADASTRADO2", gErr, CStr(objPRJEtapaVistorias.lCodigo))

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213622)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PRJEtapaVistorias_Exclui(ByVal objPRJEtapaVistorias As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PRJEtapaVistorias_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213623

    lErro = CF("PRJEtapaVistorias_Exclui_EmTrans", objPRJEtapaVistorias)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213624

    PRJEtapaVistorias_Exclui = SUCESSO

    Exit Function

Erro_PRJEtapaVistorias_Exclui:

    PRJEtapaVistorias_Exclui = gErr

    Select Case gErr

        Case 213623
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213624
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213625)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PRJEtapaVistorias_Exclui_EmTrans(ByVal objPRJEtapaVistorias As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_PRJEtapaVistorias_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213626
    Next

    'Le a tabelaPRJEtapaVistorias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PRJEtapaVistorias WHERE NumIntPRJEtapa = ? AND Codigo = ?", _
                0, lNumIntDoc, objPRJEtapaVistorias.lNumIntPRJEtapa, objPRJEtapaVistorias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213627

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213628

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 213629

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 213630
    
    objPRJEtapaVistorias.lNumIntDoc = lNumIntDoc

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PRJEtapaVistorias", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 213631
    
    lErro = PRJEtapaVistorias_Grava_UltVist(objPRJEtapaVistorias.lNumIntPRJEtapa, alComando(2), alComando(3))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not (objPRJEtapaVistorias.objAnexos Is Nothing) Then
        objPRJEtapaVistorias.objAnexos.iTipoDoc = ANEXO_TIPO_VISTPRJ
        objPRJEtapaVistorias.objAnexos.lNumIntDoc = lNumIntDoc
        lErro = CF("Anexos_Exclui", objPRJEtapaVistorias.objAnexos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PRJEtapaVistorias_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PRJEtapaVistorias_Exclui_EmTrans:

    PRJEtapaVistorias_Exclui_EmTrans = gErr

    Select Case gErr

        Case 213626
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213627, 213628
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAVISTORIAS", gErr)

        Case 213629
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAVISTORIAS_NAO_CADASTRADO", gErr, objPRJEtapaVistorias.lCodigo)

        Case 213630
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRJETAPAVISTORIAS", gErr)

        Case 213631
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRJETAPAVISTORIAS", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213632)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function PRJEtapaVistorias_Grava_UltVist(ByVal lNumIntDocEtapa As Long, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim dtData As Date, dtValid As Date

On Error GoTo Erro_PRJEtapaVistorias_Grava_UltVist

    'Le a tabelaPRJEtapaVistorias
    lErro = Comando_Executar(lComando1, "SELECT Data, DataValidade FROM PRJEtapaVistorias WHERE NumIntPRJEtapa= ? ORDER BY Data DESC, NumIntDoc DESC ", _
                dtData, dtValid, lNumIntDocEtapa)
    If lErro <> AD_SQL_SUCESSO Then gError 213623

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213624

    If lErro <> AD_SQL_SUCESSO Then
        dtData = DATA_NULA
        dtValid = DATA_NULA
    End If

    lErro = Comando_Executar(lComando2, "UPDATE PRJEtapas SET DataVistoria= ? , ValidadeVistoria= ? WHERE NumIntDoc = ? ", dtData, dtValid, lNumIntDocEtapa)
    If lErro <> AD_SQL_SUCESSO Then gError 213625

    PRJEtapaVistorias_Grava_UltVist = SUCESSO

    Exit Function

Erro_PRJEtapaVistorias_Grava_UltVist:

    PRJEtapaVistorias_Grava_UltVist = gErr

    Select Case gErr
        
        Case 213623 To 213625
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAVISTORIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213626)

    End Select

    Exit Function

End Function

