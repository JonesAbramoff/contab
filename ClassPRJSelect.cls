VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Projetos_Le(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tProjetos As typeProjetos

On Error GoTo Erro_Projetos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181296

    'Alocação de espaço no buffer
    tProjetos.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tProjetos.sNomeReduzido = String(STRING_PRJ_NOMEREDUZIDO, 0)
    tProjetos.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tProjetos.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tProjetos.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tProjetos.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tProjetos.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)
    tProjetos.sSegmento = String(STRING_MAXIMO, 0)

    'Le a tabelaProjetos
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, DataCriacao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Segmento FROM Projetos WHERE Codigo= ? AND (FilialEmpresa = ? OR FilialEmpresa = ?) ", _
                tProjetos.lNumIntDoc, tProjetos.sCodigo, tProjetos.iFilialEmpresa, tProjetos.sNomeReduzido, _
                tProjetos.sDescricao, tProjetos.dtDataCriacao, tProjetos.lCliente, tProjetos.iFilialCliente, tProjetos.sResponsavel, _
                tProjetos.sObjetivo, tProjetos.sJustificativa, tProjetos.sObservacao, tProjetos.dtDataInicio, tProjetos.dtDataFim, _
                tProjetos.dtDataInicioReal, tProjetos.dtDataFimReal, tProjetos.dPercentualComplet, tProjetos.lNumIntDocEscopo, tProjetos.sSegmento, _
                objProjetos.sCodigo, objProjetos.iFilialEmpresa, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 181297

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181298

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objProjetos.lNumIntDoc = tProjetos.lNumIntDoc
    objProjetos.sCodigo = tProjetos.sCodigo
    objProjetos.iFilialEmpresa = tProjetos.iFilialEmpresa
    objProjetos.sNomeReduzido = tProjetos.sNomeReduzido
    objProjetos.sDescricao = tProjetos.sDescricao
    objProjetos.dtDataCriacao = tProjetos.dtDataCriacao
    objProjetos.lCliente = tProjetos.lCliente
    objProjetos.iFilialCliente = tProjetos.iFilialCliente
    objProjetos.sResponsavel = tProjetos.sResponsavel
    objProjetos.sObjetivo = tProjetos.sObjetivo
    objProjetos.sJustificativa = tProjetos.sJustificativa
    objProjetos.sObservacao = tProjetos.sObservacao
    objProjetos.dtDataInicio = tProjetos.dtDataInicio
    objProjetos.dtDataFim = tProjetos.dtDataFim
    objProjetos.dtDataInicioReal = tProjetos.dtDataInicioReal
    objProjetos.dtDataFimReal = tProjetos.dtDataFimReal
    objProjetos.dPercentualComplet = tProjetos.dPercentualComplet
    objProjetos.lNumIntDocEscopo = tProjetos.lNumIntDocEscopo
    objProjetos.sSegmento = tProjetos.sSegmento

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Projetos_Le = SUCESSO

    Exit Function

Erro_Projetos_Le:

    Projetos_Le = gErr

    Select Case gErr

        Case 181296
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181297, 181298
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Projetos_Le_NomeReduzido(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tProjetos As typeProjetos

On Error GoTo Erro_Projetos_Le_NomeReduzido

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181296

    'Alocação de espaço no buffer
    tProjetos.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tProjetos.sNomeReduzido = String(STRING_PRJ_NOMEREDUZIDO, 0)
    tProjetos.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tProjetos.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tProjetos.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tProjetos.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tProjetos.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)
    tProjetos.sSegmento = String(STRING_MAXIMO, 0)

    'Le a tabelaProjetos
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, DataCriacao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Segmento FROM Projetos WHERE NomeReduzido= ? AND (FilialEmpresa = ? OR FilialEmpresa = ?) ", _
                tProjetos.lNumIntDoc, tProjetos.sCodigo, tProjetos.iFilialEmpresa, tProjetos.sNomeReduzido, _
                tProjetos.sDescricao, tProjetos.dtDataCriacao, tProjetos.lCliente, tProjetos.iFilialCliente, tProjetos.sResponsavel, _
                tProjetos.sObjetivo, tProjetos.sJustificativa, tProjetos.sObservacao, tProjetos.dtDataInicio, tProjetos.dtDataFim, _
                tProjetos.dtDataInicioReal, tProjetos.dtDataFimReal, tProjetos.dPercentualComplet, tProjetos.lNumIntDocEscopo, tProjetos.sSegmento, _
                objProjetos.sNomeReduzido, objProjetos.iFilialEmpresa, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then gError 181297

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181298

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objProjetos.lNumIntDoc = tProjetos.lNumIntDoc
    objProjetos.sCodigo = tProjetos.sCodigo
    objProjetos.iFilialEmpresa = tProjetos.iFilialEmpresa
    objProjetos.sNomeReduzido = tProjetos.sNomeReduzido
    objProjetos.sDescricao = tProjetos.sDescricao
    objProjetos.dtDataCriacao = tProjetos.dtDataCriacao
    objProjetos.lCliente = tProjetos.lCliente
    objProjetos.iFilialCliente = tProjetos.iFilialCliente
    objProjetos.sResponsavel = tProjetos.sResponsavel
    objProjetos.sObjetivo = tProjetos.sObjetivo
    objProjetos.sJustificativa = tProjetos.sJustificativa
    objProjetos.sObservacao = tProjetos.sObservacao
    objProjetos.dtDataInicio = tProjetos.dtDataInicio
    objProjetos.dtDataFim = tProjetos.dtDataFim
    objProjetos.dtDataInicioReal = tProjetos.dtDataInicioReal
    objProjetos.dtDataFimReal = tProjetos.dtDataFimReal
    objProjetos.dPercentualComplet = tProjetos.dPercentualComplet
    objProjetos.lNumIntDocEscopo = tProjetos.lNumIntDocEscopo
    objProjetos.sSegmento = tProjetos.sSegmento

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Projetos_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Projetos_Le_NomeReduzido:

    Projetos_Le_NomeReduzido = gErr

    Select Case gErr

        Case 181296
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181297, 181298
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Projetos_Le_NumIntDoc(ByVal objProjetos As ClassProjetos, Optional ByVal lComando_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tProjetos As typeProjetos

On Error GoTo Erro_Projetos_Le_NumIntDoc

    'Executa a abertura do Comando
    If lComando_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lComando_P1
    End If
    If lComando = 0 Then gError 181296

    'Alocação de espaço no buffer
    tProjetos.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tProjetos.sNomeReduzido = String(STRING_PRJ_NOMEREDUZIDO, 0)
    tProjetos.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tProjetos.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tProjetos.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tProjetos.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tProjetos.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)
    tProjetos.sSegmento = String(STRING_MAXIMO, 0)

    'Le a tabelaProjetos
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, DataCriacao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Segmento FROM Projetos WHERE NumIntDoc= ? ", _
                tProjetos.lNumIntDoc, tProjetos.sCodigo, tProjetos.iFilialEmpresa, tProjetos.sNomeReduzido, _
                tProjetos.sDescricao, tProjetos.dtDataCriacao, tProjetos.lCliente, tProjetos.iFilialCliente, tProjetos.sResponsavel, _
                tProjetos.sObjetivo, tProjetos.sJustificativa, tProjetos.sObservacao, tProjetos.dtDataInicio, tProjetos.dtDataFim, _
                tProjetos.dtDataInicioReal, tProjetos.dtDataFimReal, tProjetos.dPercentualComplet, tProjetos.lNumIntDocEscopo, tProjetos.sSegmento, _
                objProjetos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181297

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181298

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objProjetos.lNumIntDoc = tProjetos.lNumIntDoc
    objProjetos.sCodigo = tProjetos.sCodigo
    objProjetos.iFilialEmpresa = tProjetos.iFilialEmpresa
    objProjetos.sNomeReduzido = tProjetos.sNomeReduzido
    objProjetos.sDescricao = tProjetos.sDescricao
    objProjetos.dtDataCriacao = tProjetos.dtDataCriacao
    objProjetos.lCliente = tProjetos.lCliente
    objProjetos.iFilialCliente = tProjetos.iFilialCliente
    objProjetos.sResponsavel = tProjetos.sResponsavel
    objProjetos.sObjetivo = tProjetos.sObjetivo
    objProjetos.sJustificativa = tProjetos.sJustificativa
    objProjetos.sObservacao = tProjetos.sObservacao
    objProjetos.dtDataInicio = tProjetos.dtDataInicio
    objProjetos.dtDataFim = tProjetos.dtDataFim
    objProjetos.dtDataInicioReal = tProjetos.dtDataInicioReal
    objProjetos.dtDataFimReal = tProjetos.dtDataFimReal
    objProjetos.dPercentualComplet = tProjetos.dPercentualComplet
    objProjetos.lNumIntDocEscopo = tProjetos.lNumIntDocEscopo
    objProjetos.sSegmento = tProjetos.sSegmento

    'Fecha Comando
    If lComando_P1 = 0 Then Call Comando_Fechar(lComando)

    Projetos_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_Projetos_Le_NumIntDoc:

    Projetos_Le_NumIntDoc = gErr

    Select Case gErr

        Case 181296
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181297, 181298
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select

    'Fecha Comando
    If lComando_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEscopo_Le(ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEscopo As typePRJEscopo

On Error GoTo Erro_PRJEscopo_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181459
    
    'Alocação de espaço no buffer
    tPRJEscopo.sDescricao = String(STRING_PRJ_ESC_DESCRICAO, 0)
    tPRJEscopo.sExpectativa = String(STRING_PRJ_ESC_EXPECTATIVA, 0)
    tPRJEscopo.sFatoresSucesso = String(STRING_PRJ_ESC_FATORES, 0)
    tPRJEscopo.sRestricoes = String(STRING_PRJ_ESC_RESTRICOES, 0)
    tPRJEscopo.sPremissas = String(STRING_PRJ_ESC_PREMISSAS, 0)
    tPRJEscopo.sExclusoesEspecificas = String(STRING_PRJ_ESC_EXCLUSOES, 0)

    'Le a tabelaPRJEscopo
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Descricao, Expectativa, FatoresSucesso, Restricoes, Premissas, " & _
                "ExclusoesEspecificas FROM PRJEscopo WHERE NumIntDoc= ? ", _
                tPRJEscopo.lNumIntDoc, tPRJEscopo.sDescricao, tPRJEscopo.sExpectativa, tPRJEscopo.sFatoresSucesso, _
                tPRJEscopo.sRestricoes, tPRJEscopo.sPremissas, tPRJEscopo.sExclusoesEspecificas, _
                objPRJEscopo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181460

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181461

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJEscopo.lNumIntDoc = tPRJEscopo.lNumIntDoc
    objPRJEscopo.sDescricao = tPRJEscopo.sDescricao
    objPRJEscopo.sExpectativa = tPRJEscopo.sExpectativa
    objPRJEscopo.sFatoresSucesso = tPRJEscopo.sFatoresSucesso
    objPRJEscopo.sRestricoes = tPRJEscopo.sRestricoes
    objPRJEscopo.sPremissas = tPRJEscopo.sPremissas
    objPRJEscopo.sExclusoesEspecificas = tPRJEscopo.sExclusoesEspecificas

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJEscopo_Le = SUCESSO

    Exit Function

Erro_PRJEscopo_Le:

    PRJEscopo_Le = gErr

    Select Case gErr

        Case 181459
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181460, 181461
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJESCOPO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181462)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function CamposCustomizados_Le(ByVal objCamposCustomizados As ClassCamposCustomizados) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim tCamposCustomizados As typeCamposCustomizados

On Error GoTo Erro_CamposCustomizados_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181464

    'Alocação de espaço no buffer
    For iIndice = 1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES
        tCamposCustomizados.sTexto(iIndice) = String(STRING_CAMPO_CUST_TEXTO, 0)
    Next

    'Le a tabelaCamposCustomizados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, TipoNumIntDocOrigem, Data1, Data2, Data3, " & _
                "Data4, Data5, Texto1, Texto2, Texto3, Texto4, Texto5, " & _
                "Numero1, Numero2, Numero3, Numero4, Numero5, Valor1, Valor2, " & _
                "Valor3, Valor4, Valor5 FROM CamposCustomizados WHERE NumIntDocOrigem= ?  AND TipoNumIntDocOrigem= ? ", _
                tCamposCustomizados.lNumIntDoc, tCamposCustomizados.lNumIntDocOrigem, tCamposCustomizados.iTipoNumIntDocOrigem, tCamposCustomizados.dtData(1), _
                tCamposCustomizados.dtData(2), tCamposCustomizados.dtData(3), tCamposCustomizados.dtData(4), tCamposCustomizados.dtData(5), tCamposCustomizados.sTexto(1), _
                tCamposCustomizados.sTexto(2), tCamposCustomizados.sTexto(3), tCamposCustomizados.sTexto(4), tCamposCustomizados.sTexto(5), tCamposCustomizados.lNumero(1), _
                tCamposCustomizados.lNumero(2), tCamposCustomizados.lNumero(3), tCamposCustomizados.lNumero(4), tCamposCustomizados.lNumero(5), tCamposCustomizados.dValor(1), _
                tCamposCustomizados.dValor(2), tCamposCustomizados.dValor(3), tCamposCustomizados.dValor(4), tCamposCustomizados.dValor(5), _
                objCamposCustomizados.lNumIntDocOrigem, objCamposCustomizados.iTipoNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 181465

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181466

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCamposCustomizados.lNumIntDoc = tCamposCustomizados.lNumIntDoc
    objCamposCustomizados.lNumIntDocOrigem = tCamposCustomizados.lNumIntDocOrigem
    objCamposCustomizados.iTipoNumIntDocOrigem = tCamposCustomizados.iTipoNumIntDocOrigem
    
    For iIndice = 1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES
        objCamposCustomizados.dtData(iIndice) = tCamposCustomizados.dtData(iIndice)
        objCamposCustomizados.sTexto(iIndice) = tCamposCustomizados.sTexto(iIndice)
        objCamposCustomizados.lNumero(iIndice) = tCamposCustomizados.lNumero(iIndice)
        objCamposCustomizados.dValor(iIndice) = tCamposCustomizados.dValor(iIndice)
    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CamposCustomizados_Le = SUCESSO

    Exit Function

Erro_CamposCustomizados_Le:

    CamposCustomizados_Le = gErr

    Select Case gErr

        Case 181464
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181465, 181466
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOSCUSTOMIZADOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181524)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TiposCamposCust_Le(ByVal objTiposCamposCust As ClassTiposCamposCust) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTiposCamposCust As typeTiposCamposCust

On Error GoTo Erro_TiposCamposCust_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181537

    'Alocação de espaço no buffer
    tTiposCamposCust.sNomeTela = String(STRING_TIPO_CAMPO_CUST_NOMETELA, 0)
    tTiposCamposCust.sNomeTabela = String(STRING_TIPO_CAMPO_CUST_NOMETABELA, 0)

    'Le a tabelaTiposCamposCust
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeTela, NomeTabela, DatasPreenchida, TextosPreenchidos, " & _
                "NumerosPreenchidos, ValoresPreenchidos FROM TiposCamposCust WHERE Codigo= ? ", _
                tTiposCamposCust.iCodigo, tTiposCamposCust.sNomeTela, tTiposCamposCust.sNomeTabela, _
                tTiposCamposCust.iDatasPreenchida, tTiposCamposCust.iTextosPreenchidos, tTiposCamposCust.iNumerosPreenchidos, tTiposCamposCust.iValoresPreenchidos, _
                objTiposCamposCust.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 181538

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181539

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTiposCamposCust.iCodigo = tTiposCamposCust.iCodigo
    objTiposCamposCust.sNomeTela = tTiposCamposCust.sNomeTela
    objTiposCamposCust.sNomeTabela = tTiposCamposCust.sNomeTabela
    objTiposCamposCust.iDatasPreenchida = tTiposCamposCust.iDatasPreenchida
    objTiposCamposCust.iTextosPreenchidos = tTiposCamposCust.iTextosPreenchidos
    objTiposCamposCust.iNumerosPreenchidos = tTiposCamposCust.iNumerosPreenchidos
    objTiposCamposCust.iValoresPreenchidos = tTiposCamposCust.iValoresPreenchidos

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TiposCamposCust_Le = SUCESSO

    Exit Function

Erro_TiposCamposCust_Le:

    TiposCamposCust_Le = gErr

    Select Case gErr

        Case 181537
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181538, 181539
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSCAMPOSCUST", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181562)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TiposCamposCust_Le_NomeTela(ByVal objTiposCamposCust As ClassTiposCamposCust) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTiposCamposCust As typeTiposCamposCust

On Error GoTo Erro_TiposCamposCust_Le_NomeTela

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181564

    'Alocação de espaço no buffer
    tTiposCamposCust.sNomeTela = String(STRING_TIPO_CAMPO_CUST_NOMETELA, 0)
    tTiposCamposCust.sNomeTabela = String(STRING_TIPO_CAMPO_CUST_NOMETABELA, 0)

    'Le a tabelaTiposCamposCust
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeTela, NomeTabela, DatasPreenchida, TextosPreenchidos, " & _
                "NumerosPreenchidos, ValoresPreenchidos FROM TiposCamposCust WHERE NomeTela= ? ", _
                tTiposCamposCust.iCodigo, tTiposCamposCust.sNomeTela, tTiposCamposCust.sNomeTabela, _
                tTiposCamposCust.iDatasPreenchida, tTiposCamposCust.iTextosPreenchidos, tTiposCamposCust.iNumerosPreenchidos, tTiposCamposCust.iValoresPreenchidos, _
                objTiposCamposCust.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 181565

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181566

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTiposCamposCust.iCodigo = tTiposCamposCust.iCodigo
    objTiposCamposCust.sNomeTela = tTiposCamposCust.sNomeTela
    objTiposCamposCust.sNomeTabela = tTiposCamposCust.sNomeTabela
    objTiposCamposCust.iDatasPreenchida = tTiposCamposCust.iDatasPreenchida
    objTiposCamposCust.iTextosPreenchidos = tTiposCamposCust.iTextosPreenchidos
    objTiposCamposCust.iNumerosPreenchidos = tTiposCamposCust.iNumerosPreenchidos
    objTiposCamposCust.iValoresPreenchidos = tTiposCamposCust.iValoresPreenchidos

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TiposCamposCust_Le_NomeTela = SUCESSO

    Exit Function

Erro_TiposCamposCust_Le_NomeTela:

    TiposCamposCust_Le_NomeTela = gErr

    Select Case gErr

        Case 181564
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181565, 181566
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSCAMPOSCUST", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181567)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ProjetoInfo_Le(ByVal objProjetoInfo As ClassProjetoInfo) As Long

Dim lErro As Long
Dim lComando As Long
Dim tProjetoInfo As typeProjetoInfo

On Error GoTo Erro_ProjetoInfo_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181765

    'Le a tabelaPRJCustoReceitaReal
    If objProjetoInfo.iTipoOrigem <> PRJ_CR_TIPO_OP Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, NumIntDocEtapa FROM ProjetoInfo WHERE TipoOrigem= ?  AND NumIntDocOrigem= ? ", _
                    tProjetoInfo.lNumIntDoc, tProjetoInfo.lNumIntDocPRJ, tProjetoInfo.lNumIntDocEtapa, objProjetoInfo.iTipoOrigem, objProjetoInfo.lNumIntDocOrigem)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, NumIntDocEtapa FROM ProjetoInfo WHERE TipoOrigem= ?  AND CodigoOP= ? AND FilialEmpresa = ? ", _
                    tProjetoInfo.lNumIntDoc, tProjetoInfo.lNumIntDocPRJ, tProjetoInfo.lNumIntDocEtapa, objProjetoInfo.iTipoOrigem, objProjetoInfo.sCodigoOP, objProjetoInfo.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 181766

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181767

    'Sem Dados
    If lErro <> AD_SQL_SEM_DADOS Then
    
        objProjetoInfo.lNumIntDoc = tProjetoInfo.lNumIntDoc
        objProjetoInfo.lNumIntDocPRJ = tProjetoInfo.lNumIntDocPRJ
        objProjetoInfo.lNumIntDocEtapa = tProjetoInfo.lNumIntDocEtapa
        
        lErro = PRJCustoReceitaReal_Le(objProjetoInfo)
        If lErro <> SUCESSO Then gError 181768
        
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ProjetoInfo_Le = SUCESSO

    Exit Function

Erro_ProjetoInfo_Le:

    ProjetoInfo_Le = gErr

    Select Case gErr

        Case 181765
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181766, 181767
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOINFO", gErr)
            
        Case 181768

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181769)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJCustoReceitaReal_Le(ByVal objProjetoInfo As ClassProjetoInfo) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJCR As typePRJCR
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_PRJCustoReceitaReal_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181641
    
    tPRJCR.sObservacao = String(STRING_PRJ_CR_OBSERVACAO, 0)
    tPRJCR.sItem = String(STRING_PRJ_CR_ITEM, 0)

    'Le a tabelaPRJCustoReceitaReal
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, NumIntDocEtapa, Valor, Quantidade, Percentual, " & _
                "TipoValor, Observacao, Item, CalcAuto FROM PRJCustoReceitaReal WHERE NumIntDocPRJInfo = ? ", _
                tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocEtapa, _
                tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, _
                tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                objProjetoInfo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181642

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181643

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJCR = New ClassPRJCR

        objPRJCR.lNumIntDoc = tPRJCR.lNumIntDoc
        objPRJCR.lNumIntDocPRJ = tPRJCR.lNumIntDocPRJ
        objPRJCR.lNumIntDocEtapa = tPRJCR.lNumIntDocEtapa
        objPRJCR.dValor = tPRJCR.dValor
        objPRJCR.dQuantidade = tPRJCR.dQuantidade
        objPRJCR.dPercent = tPRJCR.dPercentual
        objPRJCR.iTipoValor = tPRJCR.iTipoValor
        objPRJCR.sObservacao = tPRJCR.sObservacao
        objPRJCR.sItem = tPRJCR.sItem
        objPRJCR.iCalcAuto = tPRJCR.iCalcAuto
        
        objPRJCR.lNumIntDocPRJInfo = objProjetoInfo.lNumIntDocPRJ
        objPRJCR.iTipoOrigem = objProjetoInfo.iTipoOrigem
        objPRJCR.lNumIntDocOrigem = objProjetoInfo.lNumIntDocOrigem
        
        objProjetoInfo.colPRJCustoReceitaReal.Add objPRJCR
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181644
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJCustoReceitaReal_Le = SUCESSO

    Exit Function

Erro_PRJCustoReceitaReal_Le:

    PRJCustoReceitaReal_Le = gErr

    Select Case gErr

        Case 181641
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181642, 181643, 181644
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCUSTORECEITAREAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJCustoReceitaReal_Le_Etapa(ByVal objEtapa As ClassPRJEtapas, ByVal colPRJCustoReceitaReal As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJCR As typePRJCR
Dim objPRJCR As ClassPRJCR

On Error GoTo Erro_PRJCustoReceitaReal_Le_Etapa

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189257
    
    tPRJCR.sObservacao = String(STRING_PRJ_CR_OBSERVACAO, 0)
    tPRJCR.sItem = String(STRING_PRJ_CR_ITEM, 0)

    'Le a tabelaPRJCustoReceitaReal
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, NumIntDocEtapa, NumIntDocPRJInfo, Valor, Quantidade, Percentual, " & _
                "TipoValor, Observacao, Item, CalcAuto FROM PRJCustoReceitaReal WHERE NumIntDocEtapa = ? ", _
                tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocEtapa, tPRJCR.lNumIntDocPRJInfo, _
                tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, _
                tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189258

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189259

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJCR = New ClassPRJCR

        objPRJCR.lNumIntDoc = tPRJCR.lNumIntDoc
        objPRJCR.lNumIntDocPRJ = tPRJCR.lNumIntDocPRJ
        objPRJCR.lNumIntDocEtapa = tPRJCR.lNumIntDocEtapa
        objPRJCR.dValor = tPRJCR.dValor
        objPRJCR.dQuantidade = tPRJCR.dQuantidade
        objPRJCR.dPercent = tPRJCR.dPercentual
        objPRJCR.iTipoValor = tPRJCR.iTipoValor
        objPRJCR.sObservacao = tPRJCR.sObservacao
        objPRJCR.sItem = tPRJCR.sItem
        objPRJCR.iCalcAuto = tPRJCR.iCalcAuto
        objPRJCR.lNumIntDocPRJInfo = tPRJCR.lNumIntDocPRJInfo
        
        colPRJCustoReceitaReal.Add objPRJCR
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189260
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJCustoReceitaReal_Le_Etapa = SUCESSO

    Exit Function

Erro_PRJCustoReceitaReal_Le_Etapa:

    PRJCustoReceitaReal_Le_Etapa = gErr

    Select Case gErr

        Case 189257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189258, 189259, 189260
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCUSTORECEITAREAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189261)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapas_Le(ByVal objPRJEtapas As ClassPRJEtapas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapas As typePRJEtapas

On Error GoTo Erro_PRJEtapas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181708

    'Alocação de espaço no buffer
    tPRJEtapas.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tPRJEtapas.sReferencia = String(STRING_PRJ_REFERENCIA, 0)
    tPRJEtapas.sNomeReduzido = String(STRING_ETAPAPRJ_NOMEREDUZIDO, 0)
    tPRJEtapas.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tPRJEtapas.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tPRJEtapas.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tPRJEtapas.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tPRJEtapas.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Codigo, Referencia, NomeReduzido, Descricao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, NumIntDocEtapaPaiOrg, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Nivel, Seq, Posicao, DataVistoria, ValidadeVistoria FROM PRJEtapas WHERE Codigo = ? AND NumIntDocPRJ = ? ", _
                tPRJEtapas.lNumIntDoc, tPRJEtapas.lNumIntDocPRJ, tPRJEtapas.sCodigo, tPRJEtapas.sReferencia, _
                tPRJEtapas.sNomeReduzido, tPRJEtapas.sDescricao, tPRJEtapas.lCliente, tPRJEtapas.iFilialCliente, tPRJEtapas.sResponsavel, _
                tPRJEtapas.sObjetivo, tPRJEtapas.sJustificativa, tPRJEtapas.sObservacao, tPRJEtapas.dtDataInicio, tPRJEtapas.dtDataFim, _
                tPRJEtapas.lNumIntDocEtapaPaiOrg, tPRJEtapas.dtDataInicioReal, tPRJEtapas.dtDataFimReal, tPRJEtapas.dPercentualComplet, tPRJEtapas.lNumIntDocEscopo, tPRJEtapas.iNivel, tPRJEtapas.iSeq, tPRJEtapas.iPosicao, tPRJEtapas.dtDataVistoria, tPRJEtapas.dtValidadeVistoria, _
                objPRJEtapas.sCodigo, objPRJEtapas.lNumIntDocPRJ)
    If lErro <> AD_SQL_SUCESSO Then gError 181709

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181710

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJEtapas.lNumIntDoc = tPRJEtapas.lNumIntDoc
    objPRJEtapas.lNumIntDocPRJ = tPRJEtapas.lNumIntDocPRJ
    objPRJEtapas.sCodigo = tPRJEtapas.sCodigo
    objPRJEtapas.sReferencia = tPRJEtapas.sReferencia
    objPRJEtapas.sNomeReduzido = tPRJEtapas.sNomeReduzido
    objPRJEtapas.sDescricao = tPRJEtapas.sDescricao
    objPRJEtapas.lCliente = tPRJEtapas.lCliente
    objPRJEtapas.iFilialCliente = tPRJEtapas.iFilialCliente
    objPRJEtapas.sResponsavel = tPRJEtapas.sResponsavel
    objPRJEtapas.sObjetivo = tPRJEtapas.sObjetivo
    objPRJEtapas.sJustificativa = tPRJEtapas.sJustificativa
    objPRJEtapas.sObservacao = tPRJEtapas.sObservacao
    objPRJEtapas.dtDataInicio = tPRJEtapas.dtDataInicio
    objPRJEtapas.dtDataFim = tPRJEtapas.dtDataFim
    objPRJEtapas.lNumIntDocEtapaPaiOrg = tPRJEtapas.lNumIntDocEtapaPaiOrg
    objPRJEtapas.dtDataInicioReal = tPRJEtapas.dtDataInicioReal
    objPRJEtapas.dtDataFimReal = tPRJEtapas.dtDataFimReal
    objPRJEtapas.dPercentualComplet = tPRJEtapas.dPercentualComplet
    objPRJEtapas.lNumIntDocEscopo = tPRJEtapas.lNumIntDocEscopo
    objPRJEtapas.iNivel = tPRJEtapas.iNivel
    objPRJEtapas.iSeq = tPRJEtapas.iSeq
    objPRJEtapas.iPosicao = tPRJEtapas.iPosicao
    objPRJEtapas.dtDataVistoria = tPRJEtapas.dtDataVistoria
    objPRJEtapas.dtValidadeVistoria = tPRJEtapas.dtValidadeVistoria
        
    lErro = PRJEtapaItensProd_Le(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185370

    lErro = PRJEtapaMaquinas_Le(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185371

    lErro = PRJEtapaMateriais_Le(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185372

    lErro = PRJEtapaMO_Le(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185373

    lErro = PRJEtapasPredecessoras_Le(objPRJEtapas)
    If lErro <> SUCESSO Then gError 185374
    
    lErro = PRJEtapas_Le_TiposCalculados(objPRJEtapas)
    If lErro <> SUCESSO Then gError 189294

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJEtapas_Le = SUCESSO

    Exit Function

Erro_PRJEtapas_Le:

    PRJEtapas_Le = gErr

    Select Case gErr

        Case 181708
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181709, 181710
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 185370 To 185374, 189294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181735)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJEtapas_Le_NumIntDoc(ByVal objPRJEtapas As ClassPRJEtapas, Optional ByVal lCmd_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapas As typePRJEtapas

On Error GoTo Erro_PRJEtapas_Le_NumIntDoc

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 185475

    'Alocação de espaço no buffer
    tPRJEtapas.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tPRJEtapas.sReferencia = String(STRING_PRJ_REFERENCIA, 0)
    tPRJEtapas.sNomeReduzido = String(STRING_ETAPAPRJ_NOMEREDUZIDO, 0)
    tPRJEtapas.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tPRJEtapas.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tPRJEtapas.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tPRJEtapas.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tPRJEtapas.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Codigo, Referencia, NomeReduzido, Descricao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, NumIntDocEtapaPaiOrg, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Nivel, Seq, Posicao, DataVistoria, ValidadeVistoria FROM PRJEtapas WHERE NumIntDoc = ? ", _
                tPRJEtapas.lNumIntDoc, tPRJEtapas.lNumIntDocPRJ, tPRJEtapas.sCodigo, tPRJEtapas.sReferencia, _
                tPRJEtapas.sNomeReduzido, tPRJEtapas.sDescricao, tPRJEtapas.lCliente, tPRJEtapas.iFilialCliente, tPRJEtapas.sResponsavel, _
                tPRJEtapas.sObjetivo, tPRJEtapas.sJustificativa, tPRJEtapas.sObservacao, tPRJEtapas.dtDataInicio, tPRJEtapas.dtDataFim, _
                tPRJEtapas.lNumIntDocEtapaPaiOrg, tPRJEtapas.dtDataInicioReal, tPRJEtapas.dtDataFimReal, tPRJEtapas.dPercentualComplet, tPRJEtapas.lNumIntDocEscopo, tPRJEtapas.iNivel, tPRJEtapas.iSeq, tPRJEtapas.iPosicao, tPRJEtapas.dtDataVistoria, tPRJEtapas.dtValidadeVistoria, _
                objPRJEtapas.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185476

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185477

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJEtapas.lNumIntDoc = tPRJEtapas.lNumIntDoc
    objPRJEtapas.lNumIntDocPRJ = tPRJEtapas.lNumIntDocPRJ
    objPRJEtapas.sCodigo = tPRJEtapas.sCodigo
    objPRJEtapas.sReferencia = tPRJEtapas.sReferencia
    objPRJEtapas.sNomeReduzido = tPRJEtapas.sNomeReduzido
    objPRJEtapas.sDescricao = tPRJEtapas.sDescricao
    objPRJEtapas.lCliente = tPRJEtapas.lCliente
    objPRJEtapas.iFilialCliente = tPRJEtapas.iFilialCliente
    objPRJEtapas.sResponsavel = tPRJEtapas.sResponsavel
    objPRJEtapas.sObjetivo = tPRJEtapas.sObjetivo
    objPRJEtapas.sJustificativa = tPRJEtapas.sJustificativa
    objPRJEtapas.sObservacao = tPRJEtapas.sObservacao
    objPRJEtapas.dtDataInicio = tPRJEtapas.dtDataInicio
    objPRJEtapas.dtDataFim = tPRJEtapas.dtDataFim
    objPRJEtapas.lNumIntDocEtapaPaiOrg = tPRJEtapas.lNumIntDocEtapaPaiOrg
    objPRJEtapas.dtDataInicioReal = tPRJEtapas.dtDataInicioReal
    objPRJEtapas.dtDataFimReal = tPRJEtapas.dtDataFimReal
    objPRJEtapas.dPercentualComplet = tPRJEtapas.dPercentualComplet
    objPRJEtapas.lNumIntDocEscopo = tPRJEtapas.lNumIntDocEscopo
    objPRJEtapas.iNivel = tPRJEtapas.iNivel
    objPRJEtapas.iSeq = tPRJEtapas.iSeq
    objPRJEtapas.iPosicao = tPRJEtapas.iPosicao
    objPRJEtapas.dtDataVistoria = tPRJEtapas.dtDataVistoria
    objPRJEtapas.dtValidadeVistoria = tPRJEtapas.dtValidadeVistoria
        
    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    PRJEtapas_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_NumIntDoc:

    PRJEtapas_Le_NumIntDoc = gErr

    Select Case gErr

        Case 185475
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185476, 185477
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185478)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJEtapas_Le_Projeto(ByVal objProjeto As ClassProjetos, Optional ByVal bCompleto As Boolean = False, Optional ByVal bFluxo As Boolean = False, Optional ByVal lCmd_P1 As Long = 0, Optional ByVal lCmd_P2 As Long = 0, Optional ByVal lCmd_P3 As Long = 0, Optional ByVal lCmd_P4 As Long = 0, Optional ByVal lCmd_P5 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapas As typePRJEtapas
Dim objPRJEtapas As ClassPRJEtapas

On Error GoTo Erro_PRJEtapas_Le_Projeto

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 185001

    'Alocação de espaço no buffer
    tPRJEtapas.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tPRJEtapas.sReferencia = String(STRING_PRJ_REFERENCIA, 0)
    tPRJEtapas.sNomeReduzido = String(STRING_ETAPAPRJ_NOMEREDUZIDO, 0)
    tPRJEtapas.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tPRJEtapas.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tPRJEtapas.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tPRJEtapas.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tPRJEtapas.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Codigo, Referencia, NomeReduzido, Descricao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, NumIntDocEtapaPaiOrg, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Nivel, Seq, Posicao, DataVistoria, ValidadeVistoria FROM PRJEtapas WHERE NumIntDocPRJ = ? ORDER BY Posicao ", _
                tPRJEtapas.lNumIntDoc, tPRJEtapas.lNumIntDocPRJ, tPRJEtapas.sCodigo, tPRJEtapas.sReferencia, _
                tPRJEtapas.sNomeReduzido, tPRJEtapas.sDescricao, tPRJEtapas.lCliente, tPRJEtapas.iFilialCliente, tPRJEtapas.sResponsavel, _
                tPRJEtapas.sObjetivo, tPRJEtapas.sJustificativa, tPRJEtapas.sObservacao, tPRJEtapas.dtDataInicio, tPRJEtapas.dtDataFim, _
                tPRJEtapas.lNumIntDocEtapaPaiOrg, tPRJEtapas.dtDataInicioReal, tPRJEtapas.dtDataFimReal, tPRJEtapas.dPercentualComplet, tPRJEtapas.lNumIntDocEscopo, tPRJEtapas.iNivel, tPRJEtapas.iSeq, tPRJEtapas.iPosicao, tPRJEtapas.dtDataVistoria, tPRJEtapas.dtValidadeVistoria, _
                objProjeto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185002

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185003

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJEtapas = New ClassPRJEtapas
    
        objPRJEtapas.lNumIntDoc = tPRJEtapas.lNumIntDoc
        objPRJEtapas.lNumIntDocPRJ = tPRJEtapas.lNumIntDocPRJ
        objPRJEtapas.sCodigo = tPRJEtapas.sCodigo
        objPRJEtapas.sReferencia = tPRJEtapas.sReferencia
        objPRJEtapas.sNomeReduzido = tPRJEtapas.sNomeReduzido
        objPRJEtapas.sDescricao = tPRJEtapas.sDescricao
        objPRJEtapas.lCliente = tPRJEtapas.lCliente
        objPRJEtapas.iFilialCliente = tPRJEtapas.iFilialCliente
        objPRJEtapas.sResponsavel = tPRJEtapas.sResponsavel
        objPRJEtapas.sObjetivo = tPRJEtapas.sObjetivo
        objPRJEtapas.sJustificativa = tPRJEtapas.sJustificativa
        objPRJEtapas.sObservacao = tPRJEtapas.sObservacao
        objPRJEtapas.dtDataInicio = tPRJEtapas.dtDataInicio
        objPRJEtapas.dtDataFim = tPRJEtapas.dtDataFim
        objPRJEtapas.lNumIntDocEtapaPaiOrg = tPRJEtapas.lNumIntDocEtapaPaiOrg
        objPRJEtapas.dtDataInicioReal = tPRJEtapas.dtDataInicioReal
        objPRJEtapas.dtDataFimReal = tPRJEtapas.dtDataFimReal
        objPRJEtapas.dPercentualComplet = tPRJEtapas.dPercentualComplet
        objPRJEtapas.lNumIntDocEscopo = tPRJEtapas.lNumIntDocEscopo
        objPRJEtapas.iNivel = tPRJEtapas.iNivel
        objPRJEtapas.iSeq = tPRJEtapas.iSeq
        objPRJEtapas.iPosicao = tPRJEtapas.iPosicao
        objPRJEtapas.dtDataVistoria = tPRJEtapas.dtDataVistoria
        objPRJEtapas.dtValidadeVistoria = tPRJEtapas.dtValidadeVistoria
        
        If bCompleto Then
                
            lErro = PRJEtapaItensProd_Le(objPRJEtapas, lCmd_P2)
            If lErro <> SUCESSO Then gError 187480
        
            lErro = PRJEtapaMaquinas_Le(objPRJEtapas, lCmd_P3)
            If lErro <> SUCESSO Then gError 187481
        
            lErro = PRJEtapaMateriais_Le(objPRJEtapas, lCmd_P4)
            If lErro <> SUCESSO Then gError 187482
        
            lErro = PRJEtapaMO_Le(objPRJEtapas, lCmd_P5)
            If lErro <> SUCESSO Then gError 187483
            
            If Not bFluxo Then
            
                lErro = PRJEtapasPredecessoras_Le(objPRJEtapas)
                If lErro <> SUCESSO Then gError 187484
                
                lErro = PRJEtapas_Le_TiposCalculados(objPRJEtapas)
                If lErro <> SUCESSO Then gError 189293
                
            End If
    
        End If
        
        objProjeto.colEtapas.Add objPRJEtapas
    
        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185004

    Loop

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    PRJEtapas_Le_Projeto = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_Projeto:

    PRJEtapas_Le_Projeto = gErr

    Select Case gErr

        Case 185001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185002, 185003, 185004
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)
            
        Case 187480 To 187484, 189293
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181735)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Projeto_Le_DocInfo(ByVal lNumIntDocPRJInfo As Long, sDocOrigem As String, Optional objObjetoOrigem As Object, Optional sNomeTela As String, Optional sDocOrigemRed As String, Optional iFilialDoc As Integer) As Long

Dim lErro As Long
Dim objProjetoInfo As New ClassProjetoInfo
Dim objNFiscal As New ClassNFiscal
Dim objTituloRec As New ClassTituloReceber
Dim objTituloPag As New ClassTituloPagar
Dim objNFPag As New ClassNFsPag
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objMovContaSaque As New ClassMovContaCorrente
Dim objMovContaDeposito As New ClassMovContaCorrente
Dim objOP As New ClassOrdemDeProducao
Dim objMovEstoque As New ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objParcPV As ClassParcelaPedidoVenda
Dim iStatus As Integer
Dim objForn As New ClassFornecedor
Dim objCli As New ClassCliente
Dim sNome As String

On Error GoTo Erro_Projeto_Le_DocInfo

    iFilialDoc = 0

    objProjetoInfo.lNumIntDoc = lNumIntDocPRJInfo
    
    lErro = ProjetoInfo_Le_NumIntDoc(objProjetoInfo)
    If lErro <> SUCESSO Then gError 189640
    
    Select Case objProjetoInfo.iTipoOrigem
    
        Case PRJ_CR_TIPO_NF
        
            objNFiscal.lNumIntDoc = objProjetoInfo.lNumIntDocOrigem
            
            lErro = CF("NFiscal_Le", objNFiscal)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 182053
            
            If objNFiscal.lFornecedor <> 0 Then
                objForn.lCodigo = objNFiscal.lFornecedor
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 189640
                sNome = "Forn " & objForn.lCodigo & SEPARADOR & objForn.sNomeReduzido
            ElseIf objNFiscal.lCliente <> 0 Then
                objCli.lCodigo = objNFiscal.lCliente
                lErro = CF("Cliente_Le", objCli)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 189640
                sNome = "Cli " & objCli.lCodigo & SEPARADOR & objCli.sNomeReduzido
            Else
                sNome = ""
            End If
            
            sDocOrigem = "Nota fiscal " & CStr(objNFiscal.lNumNotaFiscal) & ", série " & objNFiscal.sSerie & ", filial empresa " & CStr(objNFiscal.iFilialEmpresa) & " e " & sNome
            sDocOrigemRed = "NF " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal) & " FE " & CStr(objNFiscal.iFilialEmpresa) & " " & sNome
            
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
            
            iFilialDoc = objNFiscal.iFilialEmpresa
    
            'lê o Tipo da Nota Fiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError 189641
    
            sNomeTela = objTipoDocInfo.sNomeTelaNFiscal
            Set objObjetoOrigem = objNFiscal
        
        Case PRJ_CR_TIPO_SAQUE
        
            objMovContaSaque.lNumMovto = objProjetoInfo.lNumIntDocOrigem
        
            lErro = CF("MovContaCorrente_Le", objMovContaSaque)
            If lErro <> SUCESSO And lErro <> 11893 Then gError 185030
        
            sDocOrigem = "Saque NumMovto " & CStr(objProjetoInfo.lNumIntDocOrigem)
            sDocOrigemRed = "Saque " & CStr(objProjetoInfo.lNumIntDocOrigem)

            sNomeTela = "Saque"
            Set objObjetoOrigem = objMovContaSaque

        Case PRJ_CR_TIPO_DEPOSITO
            
            objMovContaDeposito.lNumMovto = objProjetoInfo.lNumIntDocOrigem
            
            lErro = CF("MovContaCorrente_Le", objMovContaDeposito)
            If lErro <> SUCESSO And lErro <> 11893 Then gError 185030
        
            sDocOrigem = "Depósito NumMovto " & CStr(objProjetoInfo.lNumIntDocOrigem)
            sDocOrigemRed = "Depósito " & CStr(objProjetoInfo.lNumIntDocOrigem)

            sNomeTela = "Deposito"
            Set objObjetoOrigem = objMovContaDeposito
            
        Case PRJ_CR_TIPO_TITREC
        
            objTituloRec.lNumIntDoc = objProjetoInfo.lNumIntDocOrigem
        
            sNomeTela = "TituloReceber"
        
            lErro = CF("TituloReceber_Le", objTituloRec)
            If lErro <> SUCESSO And lErro <> 26061 Then gError 182054
            If lErro <> SUCESSO Then
                lErro = CF("TituloReceberBaixado_Le", objTituloRec)
                If lErro <> SUCESSO And lErro <> 56570 Then gError 182054
                sNomeTela = "TituloReceber_Consulta"
            End If
            
            If objTituloRec.lCliente <> 0 Then
                objCli.lCodigo = objTituloRec.lCliente
                lErro = CF("Cliente_Le", objCli)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 189640
                sNome = "Cli " & objCli.lCodigo & SEPARADOR & objCli.sNomeReduzido
            Else
                sNome = ""
            End If
            
            sDocOrigem = "Título a receber " & CStr(objTituloRec.lNumTitulo) & ", tipo " & objTituloRec.sSiglaDocumento & ", filial empresa " & objTituloRec.iFilialEmpresa & " e " & sNome
            sDocOrigemRed = "Tít.Rec " & objTituloRec.sSiglaDocumento & " " & CStr(objTituloRec.lNumTitulo) & " FilEmp " & objTituloRec.iFilialEmpresa & " " & sNome
    
            iFilialDoc = objTituloRec.iFilialEmpresa
    
            Set objObjetoOrigem = objTituloRec
            
        Case PRJ_CR_TIPO_TITPAG

            objTituloPag.lNumIntDoc = objProjetoInfo.lNumIntDocOrigem
        
            If objTituloPag.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
                sNomeTela = "FaturasPag"
            Else
                sNomeTela = "NFFatPag"
            End If
            
            lErro = CF("TituloPagar_Le", objTituloPag)
            If lErro <> SUCESSO And lErro <> 18372 Then gError 182055
            If lErro <> SUCESSO Then
                lErro = CF("TituloPagarBaixado_Le", objTituloPag)
                If lErro <> SUCESSO And lErro <> 56661 Then gError 182055
                If objTituloPag.sSiglaDocumento <> TIPODOC_FATURA_A_PAGAR Then
                    sNomeTela = "TituloPagar_Consulta"
                End If
            End If
            
            If objTituloPag.lFornecedor <> 0 Then
                objForn.lCodigo = objTituloPag.lFornecedor
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 189640
                sNome = "Forn " & objForn.lCodigo & SEPARADOR & objForn.sNomeReduzido
            Else
                sNome = ""
            End If
            
            iFilialDoc = objTituloPag.iFilialEmpresa

            sDocOrigem = "Título a pagar " & objTituloPag.lNumTitulo & ", tipo " & objTituloPag.sSiglaDocumento & ", filial empresa " & objTituloPag.iFilialEmpresa & " e " & sNome
            sDocOrigemRed = "Tít.Pag. " & objTituloPag.sSiglaDocumento & " " & objTituloPag.lNumTitulo & " FilEmp " & objTituloPag.iFilialEmpresa & " " & sNome

            Set objObjetoOrigem = objTituloPag

        Case PRJ_CR_TIPO_OV
        
            objOrcamentoVenda.lNumIntDoc = objProjetoInfo.lNumIntDocOrigem
        
            lErro = CF("OrcamentoVenda_Le_NumIntDoc", objOrcamentoVenda)
            If lErro <> SUCESSO And lErro <> 94462 Then gError 182056

            sDocOrigem = "Orçamento de venda " & CStr(objOrcamentoVenda.lCodigo) & " e filial empresa " & CStr(objOrcamentoVenda.iFilialEmpresa)
            sDocOrigemRed = "OV " & CStr(objOrcamentoVenda.lCodigo) & " FilEmp " & CStr(objOrcamentoVenda.iFilialEmpresa)

            iFilialDoc = objOrcamentoVenda.iFilialEmpresa

            sNomeTela = "OrcamentoVenda"
            Set objObjetoOrigem = objOrcamentoVenda

        Case PRJ_CR_TIPO_PV
        
            objPedidoVenda.lNumIntDoc = objProjetoInfo.lNumIntDocOrigem
            
            lErro = CF("PedidoVenda_Le_NumIntDoc", objPedidoVenda)
            If lErro <> SUCESSO Then gError 185028
            
            lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPedidoVenda, iStatus)
            If lErro <> SUCESSO Then gError 185028

            sDocOrigem = "Pedido de venda código " & CStr(objPedidoVenda.lCodigo) & " e filial empresa " & CStr(objPedidoVenda.iFilialEmpresa)
            sDocOrigemRed = "PV " & CStr(objPedidoVenda.lCodigo) & " FilEmp " & CStr(objPedidoVenda.iFilialEmpresa)

            iFilialDoc = objPedidoVenda.iFilialEmpresa

            sNomeTela = "PedidoVenda"
            Set objObjetoOrigem = objPedidoVenda

        Case PRJ_CR_TIPO_NFPAG
    
            objNFPag.lNumIntDoc = objProjetoInfo.lNumIntDocOrigem
        
            lErro = CF("NFPag_Le_Todos2", objNFPag)
            If lErro <> SUCESSO And lErro <> 60139 Then gError 182058

            sDocOrigem = "Nota Fiscal a Pagar número " & CStr(objNFPag.lNumNotaFiscal) & " e filial empresa " & CStr(objNFPag.iFilialEmpresa)
            sDocOrigemRed = "NF a Pagar " & CStr(objNFPag.lNumNotaFiscal) & " FilEmp " & CStr(objNFPag.iFilialEmpresa)
            
            iFilialDoc = objNFPag.iFilialEmpresa
            
            sNomeTela = "NFPag"
            Set objObjetoOrigem = objNFPag
            
        Case PRJ_CR_TIPO_OP
            
            sDocOrigem = "Ordem de produção código " & objProjetoInfo.sCodigoOP & " e filial empresa " & CStr(objProjetoInfo.iFilialEmpresa)
            sDocOrigemRed = "OP " & objProjetoInfo.sCodigoOP & " FilEmp " & CStr(objProjetoInfo.iFilialEmpresa)
    
            objOP.sCodigo = objProjetoInfo.sCodigoOP
            objOP.iFilialEmpresa = objProjetoInfo.iFilialEmpresa

            iFilialDoc = objProjetoInfo.iFilialEmpresa

            sNomeTela = "OrdemProducao"
            Set objObjetoOrigem = objOP
    
        Case PRJ_CR_TIPO_PRODENTRADA
            
            objMovEstoque.lCodigo = objProjetoInfo.lNumIntDocOrigem
        
            sDocOrigem = "Produção entrada código " & objProjetoInfo.lNumIntDocOrigem
            sDocOrigemRed = "Prod.Ent. " & objProjetoInfo.lNumIntDocOrigem
        
            sNomeTela = "ProducaoEntrada"
            Set objObjetoOrigem = objMovEstoque
        
        Case PRJ_CR_TIPO_REQPROD
        
            objMovEstoque.lCodigo = objProjetoInfo.lNumIntDocOrigem
            
            sDocOrigem = "Requisição para produção código " & objProjetoInfo.lNumIntDocOrigem
            sDocOrigemRed = "Req.Prod. " & objProjetoInfo.lNumIntDocOrigem
    
            sNomeTela = "ProducaoSaida"
            Set objObjetoOrigem = objMovEstoque
    
    End Select
 
    Projeto_Le_DocInfo = SUCESSO

    Exit Function

Erro_Projeto_Le_DocInfo:

    Projeto_Le_DocInfo = gErr

    Select Case gErr
    
        Case 182053 To 182058, 189640, 189641

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182059)

    End Select

    Exit Function

End Function

Public Function ProjetoInfo_Le_NumIntDoc(ByVal objProjetoInfo As ClassProjetoInfo, Optional ByVal bLeTudo As Boolean = True, Optional ByVal lCmd_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tProjetoInfo As typeProjetoInfo

On Error GoTo Erro_ProjetoInfo_Le_NumIntDoc

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 182060
    
    tProjetoInfo.sCodigoOP = String(STRING_OPCODIGO, 0)

    'Le a tabelaPRJCustoReceitaReal
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, NumIntDocEtapa, TipoOrigem, NumIntDocOrigem, CodigoOP, FilialEmpresa FROM ProjetoInfo WHERE NumIntDoc= ? ", _
                tProjetoInfo.lNumIntDoc, tProjetoInfo.lNumIntDocPRJ, tProjetoInfo.lNumIntDocEtapa, tProjetoInfo.iTipoOrigem, tProjetoInfo.lNumIntDocOrigem, tProjetoInfo.sCodigoOP, tProjetoInfo.iFilialEmpresa, objProjetoInfo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 182061

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182062

    'Sem Dados
    If lErro <> AD_SQL_SEM_DADOS Then
    
        objProjetoInfo.lNumIntDoc = tProjetoInfo.lNumIntDoc
        objProjetoInfo.lNumIntDocPRJ = tProjetoInfo.lNumIntDocPRJ
        objProjetoInfo.lNumIntDocEtapa = tProjetoInfo.lNumIntDocEtapa
        objProjetoInfo.iTipoOrigem = tProjetoInfo.iTipoOrigem
        objProjetoInfo.lNumIntDocOrigem = tProjetoInfo.lNumIntDocOrigem
        objProjetoInfo.sCodigoOP = tProjetoInfo.sCodigoOP
        objProjetoInfo.iFilialEmpresa = tProjetoInfo.iFilialEmpresa
        
        If bLeTudo Then
            lErro = PRJCustoReceitaReal_Le(objProjetoInfo)
            If lErro <> SUCESSO Then gError 182063
        End If
        
    End If

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    ProjetoInfo_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_ProjetoInfo_Le_NumIntDoc:

    ProjetoInfo_Le_NumIntDoc = gErr

    Select Case gErr

        Case 182060
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 182061, 182062
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOINFO", gErr)
            
        Case 182063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182064)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CarregaCombo_Etapas(ByVal objProjeto As ClassProjetos, ByVal objCombo As Object) As Long

Dim lErro As Long
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_CarregaCombo_Etapas

    lErro = PRJEtapas_Le_Projeto(objProjeto)
    If lErro <> SUCESSO Then gError 185028
    
    objCombo.Clear
    
    For Each objEtapa In objProjeto.colEtapas
        objCombo.AddItem objEtapa.sCodigo & SEPARADOR & objEtapa.sNomeReduzido
    Next

    CarregaCombo_Etapas = SUCESSO

    Exit Function

Erro_CarregaCombo_Etapas:

    CarregaCombo_Etapas = gErr

    Select Case gErr
    
        Case 185028

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185029)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapaItensProd_Le(ByVal objEtapa As ClassPRJEtapas, Optional ByVal lCmd_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapaItensProd As typePRJEtapaItensProd
Dim objPRJEtapaItensProd As ClassPRJEtapaItensProd

On Error GoTo Erro_PRJEtapaItensProd_Le

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 185329

    'Alocação de espaço no buffer
    tPRJEtapaItensProd.sProduto = String(STRING_PRODUTO, 0)
    tPRJEtapaItensProd.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tPRJEtapaItensProd.sVersao = String(STRING_KIT_VERSAO, 0)
    tPRJEtapaItensProd.sUM = String(STRING_UM_SIGLA, 0)
    tPRJEtapaItensProd.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaPRJEtapaItensProd
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocEtapaPRJ, Seq, Produto, Descricao, Versao, " & _
                "UM, Quantidade, Observacao FROM PRJEtapaItensProd WHERE NumIntDocEtapaPRJ= ? ", _
                tPRJEtapaItensProd.lNumIntDoc, tPRJEtapaItensProd.lNumIntDocEtapaPRJ, tPRJEtapaItensProd.iSeq, tPRJEtapaItensProd.sProduto, _
                tPRJEtapaItensProd.sDescricao, tPRJEtapaItensProd.sVersao, tPRJEtapaItensProd.sUM, tPRJEtapaItensProd.dQuantidade, tPRJEtapaItensProd.sObservacao, _
                objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185330

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185331

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJEtapaItensProd = New ClassPRJEtapaItensProd
    
        objPRJEtapaItensProd.lNumIntDoc = tPRJEtapaItensProd.lNumIntDoc
        objPRJEtapaItensProd.lNumIntDocEtapaPRJ = tPRJEtapaItensProd.lNumIntDocEtapaPRJ
        objPRJEtapaItensProd.iSeq = tPRJEtapaItensProd.iSeq
        objPRJEtapaItensProd.sProduto = tPRJEtapaItensProd.sProduto
        objPRJEtapaItensProd.sDescricao = tPRJEtapaItensProd.sDescricao
        objPRJEtapaItensProd.sVersao = tPRJEtapaItensProd.sVersao
        objPRJEtapaItensProd.sUM = tPRJEtapaItensProd.sUM
        objPRJEtapaItensProd.dQuantidade = tPRJEtapaItensProd.dQuantidade
        objPRJEtapaItensProd.sObservacao = tPRJEtapaItensProd.sObservacao
        
        objEtapa.colItensProduzidos.Add objPRJEtapaItensProd
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185332
        
    Loop

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    PRJEtapaItensProd_Le = SUCESSO

    Exit Function

Erro_PRJEtapaItensProd_Le:

    PRJEtapaItensProd_Le = gErr

    Select Case gErr

        Case 185329
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185330, 185331, 185332
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAITENSPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185404)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapaMaquinas_Le(ByVal objEtapa As ClassPRJEtapas, Optional ByVal lCmd_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapaMaquinas As typePRJEtapaMaquinas
Dim objPRJEtapaMaquinas As ClassPRJEtapaMaquinas

On Error GoTo Erro_PRJEtapaMaquinas_Le

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 185352

    'Alocação de espaço no buffer
    tPRJEtapaMaquinas.sDescricao = String(STRING_MAQUINA_DESCRICAO, 0)
    tPRJEtapaMaquinas.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapaMaquinas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocEtapa, Seq, NumIntDocMaq, Descricao, Quantidade, " & _
                "Horas, Custo, Tipo, Data, Observacao FROM PRJEtapaMaquinas WHERE NumIntDocEtapa= ? ", _
                tPRJEtapaMaquinas.lNumIntDoc, tPRJEtapaMaquinas.lNumIntDocEtapa, tPRJEtapaMaquinas.iSeq, tPRJEtapaMaquinas.lNumIntDocMaq, _
                tPRJEtapaMaquinas.sDescricao, tPRJEtapaMaquinas.iQuantidade, tPRJEtapaMaquinas.dHoras, tPRJEtapaMaquinas.dCusto, tPRJEtapaMaquinas.iTipo, tPRJEtapaMaquinas.dtData, tPRJEtapaMaquinas.sObservacao, _
                objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185353

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185354

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJEtapaMaquinas = New ClassPRJEtapaMaquinas

        objPRJEtapaMaquinas.lNumIntDoc = tPRJEtapaMaquinas.lNumIntDoc
        objPRJEtapaMaquinas.lNumIntDocEtapa = tPRJEtapaMaquinas.lNumIntDocEtapa
        objPRJEtapaMaquinas.iSeq = tPRJEtapaMaquinas.iSeq
        objPRJEtapaMaquinas.lNumIntDocMaq = tPRJEtapaMaquinas.lNumIntDocMaq
        objPRJEtapaMaquinas.sDescricao = tPRJEtapaMaquinas.sDescricao
        objPRJEtapaMaquinas.iQuantidade = tPRJEtapaMaquinas.iQuantidade
        objPRJEtapaMaquinas.dHoras = tPRJEtapaMaquinas.dHoras
        objPRJEtapaMaquinas.dCusto = tPRJEtapaMaquinas.dCusto
        objPRJEtapaMaquinas.iTipo = tPRJEtapaMaquinas.iTipo
        objPRJEtapaMaquinas.dtData = tPRJEtapaMaquinas.dtData
        objPRJEtapaMaquinas.sObservacao = tPRJEtapaMaquinas.sObservacao
        
        objEtapa.colMaquinas.Add objPRJEtapaMaquinas
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185355
        
    Loop

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    PRJEtapaMaquinas_Le = SUCESSO

    Exit Function

Erro_PRJEtapaMaquinas_Le:

    PRJEtapaMaquinas_Le = gErr

    Select Case gErr

        Case 185352
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185353, 185354, 185355
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185356)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapaMO_Le(ByVal objEtapa As ClassPRJEtapas, Optional ByVal lCmd_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapaMO As typePRJEtapaMO
Dim objPRJEtapaMO As ClassPRJEtapaMO

On Error GoTo Erro_PRJEtapaMO_Le

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 185375

    'Alocação de espaço no buffer
    tPRJEtapaMO.sDescricao = String(STRING_TIPO_MO_DESCRICAO, 0)
    tPRJEtapaMO.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapaMO
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocEtapa, Seq, MaoDeObra, Descricao, Quantidade, " & _
                "Horas, Custo, Tipo, Data, Observacao FROM PRJEtapaMO WHERE NumIntDocEtapa= ? ", _
                tPRJEtapaMO.lNumIntDoc, tPRJEtapaMO.lNumIntDocEtapa, tPRJEtapaMO.iSeq, tPRJEtapaMO.iMaoDeObra, _
                tPRJEtapaMO.sDescricao, tPRJEtapaMO.iQuantidade, tPRJEtapaMO.dHoras, tPRJEtapaMO.dCusto, tPRJEtapaMO.iTipo, tPRJEtapaMO.dtData, tPRJEtapaMO.sObservacao, _
                objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185376

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185377

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJEtapaMO = New ClassPRJEtapaMO

        objPRJEtapaMO.lNumIntDoc = tPRJEtapaMO.lNumIntDoc
        objPRJEtapaMO.lNumIntDocEtapa = tPRJEtapaMO.lNumIntDocEtapa
        objPRJEtapaMO.iSeq = tPRJEtapaMO.iSeq
        objPRJEtapaMO.iMaoDeObra = tPRJEtapaMO.iMaoDeObra
        objPRJEtapaMO.sDescricao = tPRJEtapaMO.sDescricao
        objPRJEtapaMO.iQuantidade = tPRJEtapaMO.iQuantidade
        objPRJEtapaMO.dHoras = tPRJEtapaMO.dHoras
        objPRJEtapaMO.dCusto = tPRJEtapaMO.dCusto
        objPRJEtapaMO.iTipo = tPRJEtapaMO.iTipo
        objPRJEtapaMO.dtData = tPRJEtapaMO.dtData
        objPRJEtapaMO.sObservacao = tPRJEtapaMO.sObservacao
        
        objEtapa.colMaoDeObra.Add objPRJEtapaMO
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185378
        
    Loop

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    PRJEtapaMO_Le = SUCESSO

    Exit Function

Erro_PRJEtapaMO_Le:

    PRJEtapaMO_Le = gErr

    Select Case gErr

        Case 185375
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185376, 185377, 185378
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185357)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapasPredecessoras_Le(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapasPredecessoras As typePRJEtapasPredecessoras
Dim objPRJEtapasPredecessoras As ClassPRJEtapasPredecessoras

On Error GoTo Erro_PRJEtapasPredecessoras_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185398

    'Le a tabelaPRJEtapasPredecessoras
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocEtapa, NumIntDocEtapaPre FROM PRJEtapasPredecessoras WHERE NumIntDocEtapa= ? ", _
                tPRJEtapasPredecessoras.lNumIntDoc, tPRJEtapasPredecessoras.lNumIntDocEtapa, tPRJEtapasPredecessoras.lNumIntDocEtapaPre, _
                objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185399

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185400

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJEtapasPredecessoras = New ClassPRJEtapasPredecessoras

        objPRJEtapasPredecessoras.lNumIntDoc = tPRJEtapasPredecessoras.lNumIntDoc
        objPRJEtapasPredecessoras.lNumIntDocEtapa = tPRJEtapasPredecessoras.lNumIntDocEtapa
        objPRJEtapasPredecessoras.lNumIntDocEtapaPre = tPRJEtapasPredecessoras.lNumIntDocEtapaPre

        objEtapa.colPredecessores.Add objPRJEtapasPredecessoras

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185401

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJEtapasPredecessoras_Le = SUCESSO

    Exit Function

Erro_PRJEtapasPredecessoras_Le:

    PRJEtapasPredecessoras_Le = gErr

    Select Case gErr

        Case 185398
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185399, 185400, 185401
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPASPREDECESSORAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185358)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapaMateriais_Le(ByVal objEtapa As ClassPRJEtapas, Optional ByVal lCmd_P1 As Long = 0) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapaMateriais As typePRJEtapaMateriais
Dim objPRJEtapaMateriais As ClassPRJEtapaMateriais

On Error GoTo Erro_PRJEtapaMateriais_Le

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        lComando = Comando_Abrir()
    Else
        lComando = lCmd_P1
    End If
    If lComando = 0 Then gError 185421

    'Alocação de espaço no buffer
    tPRJEtapaMateriais.sProduto = String(STRING_PRODUTO, 0)
    tPRJEtapaMateriais.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tPRJEtapaMateriais.sUM = String(STRING_UM_SIGLA, 0)
    tPRJEtapaMateriais.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapaMateriais
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocEtapa, Seq, Produto,  Descricao, " & _
                "UM, Quantidade, Custo, Tipo, Data, Observacao FROM PRJEtapaMateriais WHERE NumIntDocEtapa= ? ", _
                tPRJEtapaMateriais.lNumIntDoc, tPRJEtapaMateriais.lNumIntDocEtapa, tPRJEtapaMateriais.iSeq, tPRJEtapaMateriais.sProduto, _
                 tPRJEtapaMateriais.sDescricao, tPRJEtapaMateriais.sUM, tPRJEtapaMateriais.dQuantidade, tPRJEtapaMateriais.dCusto, _
                tPRJEtapaMateriais.iTipo, tPRJEtapaMateriais.dtData, tPRJEtapaMateriais.sObservacao, _
                objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185422

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185423

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJEtapaMateriais = New ClassPRJEtapaMateriais

        objPRJEtapaMateriais.lNumIntDoc = tPRJEtapaMateriais.lNumIntDoc
        objPRJEtapaMateriais.lNumIntDocEtapa = tPRJEtapaMateriais.lNumIntDocEtapa
        objPRJEtapaMateriais.iSeq = tPRJEtapaMateriais.iSeq
        objPRJEtapaMateriais.sProduto = tPRJEtapaMateriais.sProduto
        objPRJEtapaMateriais.sVersao = tPRJEtapaMateriais.sVersao
        objPRJEtapaMateriais.sDescricao = tPRJEtapaMateriais.sDescricao
        objPRJEtapaMateriais.sUM = tPRJEtapaMateriais.sUM
        objPRJEtapaMateriais.dQuantidade = tPRJEtapaMateriais.dQuantidade
        objPRJEtapaMateriais.dCusto = tPRJEtapaMateriais.dCusto
        objPRJEtapaMateriais.iTipo = tPRJEtapaMateriais.iTipo
        objPRJEtapaMateriais.dtData = tPRJEtapaMateriais.dtData
        objPRJEtapaMateriais.sObservacao = tPRJEtapaMateriais.sObservacao
        
        objEtapa.colMateriaPrima.Add objPRJEtapaMateriais
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185424
        
    Loop

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    PRJEtapaMateriais_Le = SUCESSO

    Exit Function

Erro_PRJEtapaMateriais_Le:

    PRJEtapaMateriais_Le = gErr

    Select Case gErr

        Case 185421
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185422, 185423, 185424
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAMATERIAIS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185425)

    End Select

    'Fecha Comando
    If lCmd_P1 = 0 Then Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJEtapas_Le_Descendentes(ByVal objPRJEtapas As ClassPRJEtapas, ByVal colEtapasDescendentes As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapas As typePRJEtapas
Dim objEtapaDesc As ClassPRJEtapas

On Error GoTo Erro_PRJEtapas_Le_Descendentes

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185536

    'Alocação de espaço no buffer
    tPRJEtapas.sCodigo = String(STRING_PRJ_CODIGO, 0)
    tPRJEtapas.sReferencia = String(STRING_PRJ_REFERENCIA, 0)
    tPRJEtapas.sNomeReduzido = String(STRING_ETAPAPRJ_NOMEREDUZIDO, 0)
    tPRJEtapas.sDescricao = String(STRING_PRJ_DESCRICAO, 0)
    tPRJEtapas.sResponsavel = String(STRING_PRJ_RESPONSAVEL, 0)
    tPRJEtapas.sObjetivo = String(STRING_PRJ_OBJETIVO, 0)
    tPRJEtapas.sJustificativa = String(STRING_PRJ_JUSTIFICATIVA, 0)
    tPRJEtapas.sObservacao = String(STRING_PRJ_OBSERVACAO, 0)

    'Le a tabelaPRJEtapas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Codigo, Referencia, NomeReduzido, Descricao, " & _
                "Cliente, FilialCliente, Responsavel, Objetivo, Justificativa, Observacao, DataInicio, " & _
                "DataFim, NumIntDocEtapaPaiOrg, DataInicioReal, DataFimReal, PercentualComplet, NumIntDocEscopo, Nivel, Seq, Posicao FROM PRJEtapas WHERE Codigo LIKE ? AND Codigo <> ? AND NumIntDocPRJ = ? ", _
                tPRJEtapas.lNumIntDoc, tPRJEtapas.lNumIntDocPRJ, tPRJEtapas.sCodigo, tPRJEtapas.sReferencia, _
                tPRJEtapas.sNomeReduzido, tPRJEtapas.sDescricao, tPRJEtapas.lCliente, tPRJEtapas.iFilialCliente, tPRJEtapas.sResponsavel, _
                tPRJEtapas.sObjetivo, tPRJEtapas.sJustificativa, tPRJEtapas.sObservacao, tPRJEtapas.dtDataInicio, tPRJEtapas.dtDataFim, _
                tPRJEtapas.lNumIntDocEtapaPaiOrg, tPRJEtapas.dtDataInicioReal, tPRJEtapas.dtDataFimReal, tPRJEtapas.dPercentualComplet, tPRJEtapas.lNumIntDocEscopo, tPRJEtapas.iNivel, tPRJEtapas.iSeq, tPRJEtapas.iPosicao, _
                objPRJEtapas.sCodigo + "%", objPRJEtapas.sCodigo, objPRJEtapas.lNumIntDocPRJ)
    If lErro <> AD_SQL_SUCESSO Then gError 185537

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185538

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objEtapaDesc = New ClassPRJEtapas

        objEtapaDesc.lNumIntDoc = tPRJEtapas.lNumIntDoc
        objEtapaDesc.lNumIntDocPRJ = tPRJEtapas.lNumIntDocPRJ
        objEtapaDesc.sCodigo = tPRJEtapas.sCodigo
        objEtapaDesc.sReferencia = tPRJEtapas.sReferencia
        objEtapaDesc.sNomeReduzido = tPRJEtapas.sNomeReduzido
        objEtapaDesc.sDescricao = tPRJEtapas.sDescricao
        objEtapaDesc.lCliente = tPRJEtapas.lCliente
        objEtapaDesc.iFilialCliente = tPRJEtapas.iFilialCliente
        objEtapaDesc.sResponsavel = tPRJEtapas.sResponsavel
        objEtapaDesc.sObjetivo = tPRJEtapas.sObjetivo
        objEtapaDesc.sJustificativa = tPRJEtapas.sJustificativa
        objEtapaDesc.sObservacao = tPRJEtapas.sObservacao
        objEtapaDesc.dtDataInicio = tPRJEtapas.dtDataInicio
        objEtapaDesc.dtDataFim = tPRJEtapas.dtDataFim
        objEtapaDesc.lNumIntDocEtapaPaiOrg = tPRJEtapas.lNumIntDocEtapaPaiOrg
        objEtapaDesc.dtDataInicioReal = tPRJEtapas.dtDataInicioReal
        objEtapaDesc.dtDataFimReal = tPRJEtapas.dtDataFimReal
        objEtapaDesc.dPercentualComplet = tPRJEtapas.dPercentualComplet
        objEtapaDesc.lNumIntDocEscopo = tPRJEtapas.lNumIntDocEscopo
        objEtapaDesc.iNivel = tPRJEtapas.iNivel
        objEtapaDesc.iSeq = tPRJEtapas.iSeq
        objEtapaDesc.iPosicao = tPRJEtapas.iPosicao
        
        colEtapasDescendentes.Add objEtapaDesc
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185539

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJEtapas_Le_Descendentes = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_Descendentes:

    PRJEtapas_Le_Descendentes = gErr

    Select Case gErr

        Case 185536
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185537 To 185539
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185540)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJRecebPagto_Le(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJRecebPagto As typePRJRecebPagto

On Error GoTo Erro_PRJRecebPagto_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185554

    'Le a tabelaPRJRecebPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Tipo, Numero, Valor, CliForn, " & _
                "Filial, NumIntDocProposta,NumIntDocContrato, IncluiCFF, FilialEmpresa FROM PRJRecebPagto WHERE Numero= ?  AND FilialEmpresa= ? AND Tipo = ?", _
                tPRJRecebPagto.lNumIntDoc, tPRJRecebPagto.lNumIntDocPRJ, tPRJRecebPagto.iTipo, tPRJRecebPagto.lNumero, _
                tPRJRecebPagto.dValor, tPRJRecebPagto.lCliForn, tPRJRecebPagto.iFilial, tPRJRecebPagto.lNumIntDocProposta, tPRJRecebPagto.lNumIntDocContrato, tPRJRecebPagto.iIncluiCFF, _
                tPRJRecebPagto.iFilialEmpresa, _
                objPRJRecebPagto.lNumero, objPRJRecebPagto.iFilialEmpresa, objPRJRecebPagto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 185555

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185556

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJRecebPagto.lNumIntDoc = tPRJRecebPagto.lNumIntDoc
    objPRJRecebPagto.lNumIntDocPRJ = tPRJRecebPagto.lNumIntDocPRJ
    objPRJRecebPagto.lNumIntDocContrato = tPRJRecebPagto.lNumIntDocContrato
    objPRJRecebPagto.iTipo = tPRJRecebPagto.iTipo
    objPRJRecebPagto.lNumero = tPRJRecebPagto.lNumero
    objPRJRecebPagto.dValor = tPRJRecebPagto.dValor
    objPRJRecebPagto.lCliForn = tPRJRecebPagto.lCliForn
    objPRJRecebPagto.iFilial = tPRJRecebPagto.iFilial
    objPRJRecebPagto.lNumIntDocProposta = tPRJRecebPagto.lNumIntDocProposta
    objPRJRecebPagto.iIncluiCFF = tPRJRecebPagto.iIncluiCFF
    objPRJRecebPagto.iFilialEmpresa = tPRJRecebPagto.iFilialEmpresa
    
    lErro = PRJRecebPagtoRegras_Le(objPRJRecebPagto)
    If lErro <> SUCESSO Then gError 185716

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJRecebPagto_Le = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Le:

    PRJRecebPagto_Le = gErr

    Select Case gErr

        Case 185554
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185555, 185556
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 185716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185557)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJRecebPagto_Le_CFF(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJRecebPagto As typePRJRecebPagto

On Error GoTo Erro_PRJRecebPagto_Le_CFF

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187542

    'Le a tabelaPRJRecebPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Tipo, Numero, Valor, CliForn, " & _
                "Filial, NumIntDocProposta,NumIntDocContrato, IncluiCFF, FilialEmpresa FROM PRJRecebPagto WHERE NumIntDocPRJ= ? AND IncluiCFF= ? AND Tipo = ?", _
                tPRJRecebPagto.lNumIntDoc, tPRJRecebPagto.lNumIntDocPRJ, tPRJRecebPagto.iTipo, tPRJRecebPagto.lNumero, _
                tPRJRecebPagto.dValor, tPRJRecebPagto.lCliForn, tPRJRecebPagto.iFilial, tPRJRecebPagto.lNumIntDocProposta, tPRJRecebPagto.lNumIntDocContrato, tPRJRecebPagto.iIncluiCFF, _
                tPRJRecebPagto.iFilialEmpresa, _
                objPRJRecebPagto.lNumIntDocPRJ, MARCADO, objPRJRecebPagto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 187543

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187544

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJRecebPagto.lNumIntDoc = tPRJRecebPagto.lNumIntDoc
    objPRJRecebPagto.lNumIntDocPRJ = tPRJRecebPagto.lNumIntDocPRJ
    objPRJRecebPagto.iTipo = tPRJRecebPagto.iTipo
    objPRJRecebPagto.lNumero = tPRJRecebPagto.lNumero
    objPRJRecebPagto.dValor = tPRJRecebPagto.dValor
    objPRJRecebPagto.lCliForn = tPRJRecebPagto.lCliForn
    objPRJRecebPagto.iFilial = tPRJRecebPagto.iFilial
    objPRJRecebPagto.lNumIntDocProposta = tPRJRecebPagto.lNumIntDocProposta
    objPRJRecebPagto.iIncluiCFF = tPRJRecebPagto.iIncluiCFF
    objPRJRecebPagto.iFilialEmpresa = tPRJRecebPagto.iFilialEmpresa
    objPRJRecebPagto.lNumIntDocContrato = tPRJRecebPagto.lNumIntDocContrato

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJRecebPagto_Le_CFF = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Le_CFF:

    PRJRecebPagto_Le_CFF = gErr

    Select Case gErr

        Case 187542
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187543, 187544
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187545)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJRecebPagtoRegras_Le(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJRecebPagtoRegras As typePRJRecebPagtoRegras
Dim objPRJRecebPagtoRegras As ClassPRJRecebPagtoRegras

On Error GoTo Erro_PRJRecebPagtoRegras_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185577

    'Alocação de espaço no buffer
    tPRJRecebPagtoRegras.sRegra = String(STRING_REGRAWFW_REGRA, 0)
    tPRJRecebPagtoRegras.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaPRJRecebPagtoRegras
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocRecebPagto, Regra, Percentual, CondPagto, Observacao FROM PRJRecebPagtoRegras WHERE NumIntDocRecebPagto= ? ", _
                tPRJRecebPagtoRegras.lNumIntDoc, tPRJRecebPagtoRegras.lNumIntDocRecebPagto, tPRJRecebPagtoRegras.sRegra, tPRJRecebPagtoRegras.dPercentual, _
                tPRJRecebPagtoRegras.iCondPagto, tPRJRecebPagtoRegras.sObservacao, _
                objPRJRecebPagto.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 185578

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185579

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJRecebPagtoRegras = New ClassPRJRecebPagtoRegras

        objPRJRecebPagtoRegras.lNumIntDoc = tPRJRecebPagtoRegras.lNumIntDoc
        objPRJRecebPagtoRegras.lNumIntDocRecebPagto = tPRJRecebPagtoRegras.lNumIntDocRecebPagto
        objPRJRecebPagtoRegras.sRegra = tPRJRecebPagtoRegras.sRegra
        objPRJRecebPagtoRegras.dPercentual = tPRJRecebPagtoRegras.dPercentual
        objPRJRecebPagtoRegras.iCondPagto = tPRJRecebPagtoRegras.iCondPagto
        objPRJRecebPagtoRegras.sObservacao = tPRJRecebPagtoRegras.sObservacao
    
        objPRJRecebPagto.colRegras.Add objPRJRecebPagtoRegras
    
        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185580

    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJRecebPagtoRegras_Le = SUCESSO

    Exit Function

Erro_PRJRecebPagtoRegras_Le:

    PRJRecebPagtoRegras_Le = gErr

    Select Case gErr

        Case 185577
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185578, 185579, 185580
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTOREGRAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185715)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MnemonicoPRJ_Le(ByVal colMnemonico As Collection) As Long
'retorna os mnemonicos

Dim lComando As Long
Dim tMnemonicoPRJ As typeMnemonicoPRJ
Dim lErro As Long
Dim objMnemonico As ClassMnemonicoPRJ

On Error GoTo Erro_MnemonicoPRJ_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187109
    
    tMnemonicoPRJ.sMnemonico = String(STRING_MNEMONICOCTB_MNEMONICO, 0)
    tMnemonicoPRJ.sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoPRJ.sMnemonicoCombo = String(STRING_MNEMONICOCTB_MNEMONICOCOMBO, 0)
    tMnemonicoPRJ.sMnemonicoDesc = String(STRING_MNEMONICOCTB_MNEMONICODESC, 0)
    
    'seleciona os mnemonicos
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoCombo, MnemonicoDesc FROM MnemonicoPRJ", _
    tMnemonicoPRJ.sMnemonico, tMnemonicoPRJ.iTipo, tMnemonicoPRJ.iNumParam, tMnemonicoPRJ.iParam1, tMnemonicoPRJ.iParam2, tMnemonicoPRJ.iParam3, tMnemonicoPRJ.sNomeGrid, tMnemonicoPRJ.sMnemonicoCombo, tMnemonicoPRJ.sMnemonicoDesc)
    If lErro <> AD_SQL_SUCESSO Then gError 187110
    
    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187111
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objMnemonico = New ClassMnemonicoPRJ
    
        objMnemonico.sMnemonico = tMnemonicoPRJ.sMnemonico
        objMnemonico.iTipo = tMnemonicoPRJ.iTipo
        objMnemonico.iNumParam = tMnemonicoPRJ.iNumParam
        objMnemonico.iParam1 = tMnemonicoPRJ.iParam1
        objMnemonico.iParam2 = tMnemonicoPRJ.iParam2
        objMnemonico.iParam3 = tMnemonicoPRJ.iParam3
        objMnemonico.sGrid = tMnemonicoPRJ.sNomeGrid
        objMnemonico.sMnemonicoCombo = tMnemonicoPRJ.sMnemonicoCombo
        objMnemonico.sMnemonicoDesc = tMnemonicoPRJ.sMnemonicoDesc
        
        colMnemonico.Add objMnemonico
        
        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187112
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    MnemonicoPRJ_Le = SUCESSO

    Exit Function

Erro_MnemonicoPRJ_Le:

    MnemonicoPRJ_Le = gErr
    
    Select Case gErr
    
        Case 187109
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 187110 To 187112
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187073)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MnemonicoPRJ_Le_Mnemonico(ByVal objMnemonico As ClassMnemonicoPRJ) As Long
'retorna o mnemonico em questao

Dim lComando As Long
Dim tMnemonicoPRJ As typeMnemonicoPRJ
Dim lErro As Long

On Error GoTo Erro_MnemonicoPRJ_Le_Mnemonico

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187113
    
    tMnemonicoPRJ.sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoPRJ.sMnemonicoCombo = String(STRING_MNEMONICOCTB_MNEMONICOCOMBO, 0)
    tMnemonicoPRJ.sMnemonicoDesc = String(STRING_MNEMONICOCTB_MNEMONICODESC, 0)
    
    'seleciona os mnemonicos do Modulo/Transacao passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoCombo, MnemonicoDesc FROM MnemonicoPRJ WHERE MnemonicoCombo=?", _
    tMnemonicoPRJ.iTipo, tMnemonicoPRJ.iNumParam, tMnemonicoPRJ.iParam1, tMnemonicoPRJ.iParam2, tMnemonicoPRJ.iParam3, tMnemonicoPRJ.sNomeGrid, tMnemonicoPRJ.sMnemonicoCombo, tMnemonicoPRJ.sMnemonicoDesc, objMnemonico.sMnemonicoCombo)
    If lErro <> AD_SQL_SUCESSO Then gError 187114
    
    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187115
    
    If lErro = AD_SQL_SEM_DADOS Then gError 187116
    
    objMnemonico.iTipo = tMnemonicoPRJ.iTipo
    objMnemonico.iNumParam = tMnemonicoPRJ.iNumParam
    objMnemonico.iParam1 = tMnemonicoPRJ.iParam1
    objMnemonico.iParam2 = tMnemonicoPRJ.iParam2
    objMnemonico.iParam3 = tMnemonicoPRJ.iParam3
    objMnemonico.sGrid = tMnemonicoPRJ.sNomeGrid
    objMnemonico.sMnemonicoCombo = tMnemonicoPRJ.sMnemonicoCombo
    objMnemonico.sMnemonicoDesc = tMnemonicoPRJ.sMnemonicoDesc
        
    Call Comando_Fechar(lComando)
    
    MnemonicoPRJ_Le_Mnemonico = SUCESSO

    Exit Function

Erro_MnemonicoPRJ_Le_Mnemonico:

    MnemonicoPRJ_Le_Mnemonico = gErr
    
    Select Case gErr
    
        Case 187113
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 187114, 187115
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOPRJ1", gErr, objMnemonico.sMnemonico)
            
        Case 187116
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187098)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function PRJRecebPagto_Le_Projeto(ByVal objPRJ As ClassProjetos, ByVal colPRJRecebPagto As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim objPRJRecebPagto As ClassPRJRecebPagto
Dim tPRJRecebPagto As typePRJRecebPagto

On Error GoTo Erro_PRJRecebPagto_Le_Projeto

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187448

    'Le a tabelaPRJRecebPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Tipo, Numero, Valor, CliForn, " & _
                "Filial, NumIntDocProposta,NumIntDocContrato, IncluiCFF, FilialEmpresa FROM PRJRecebPagto WHERE NumIntDocPRJ= ? AND IncluiCFF <> 0", _
                tPRJRecebPagto.lNumIntDoc, tPRJRecebPagto.lNumIntDocPRJ, tPRJRecebPagto.iTipo, tPRJRecebPagto.lNumero, _
                tPRJRecebPagto.dValor, tPRJRecebPagto.lCliForn, tPRJRecebPagto.iFilial, tPRJRecebPagto.lNumIntDocProposta, tPRJRecebPagto.lNumIntDocContrato, tPRJRecebPagto.iIncluiCFF, _
                tPRJRecebPagto.iFilialEmpresa, _
                objPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187449

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187450

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJRecebPagto = New ClassPRJRecebPagto

        objPRJRecebPagto.lNumIntDoc = tPRJRecebPagto.lNumIntDoc
        objPRJRecebPagto.lNumIntDocPRJ = tPRJRecebPagto.lNumIntDocPRJ
        objPRJRecebPagto.iTipo = tPRJRecebPagto.iTipo
        objPRJRecebPagto.lNumero = tPRJRecebPagto.lNumero
        objPRJRecebPagto.dValor = tPRJRecebPagto.dValor
        objPRJRecebPagto.lCliForn = tPRJRecebPagto.lCliForn
        objPRJRecebPagto.iFilial = tPRJRecebPagto.iFilial
        objPRJRecebPagto.lNumIntDocProposta = tPRJRecebPagto.lNumIntDocProposta
        objPRJRecebPagto.iIncluiCFF = tPRJRecebPagto.iIncluiCFF
        objPRJRecebPagto.iFilialEmpresa = tPRJRecebPagto.iFilialEmpresa
        objPRJRecebPagto.lNumIntDocContrato = tPRJRecebPagto.lNumIntDocContrato
        
        lErro = PRJRecebPagtoRegras_Le(objPRJRecebPagto)
        If lErro <> SUCESSO Then gError 187452
        
        colPRJRecebPagto.Add objPRJRecebPagto
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187451
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJRecebPagto_Le_Projeto = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Le_Projeto:

    PRJRecebPagto_Le_Projeto = gErr

    Select Case gErr

        Case 187448
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187449 To 187451
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)
 
        Case 187452

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187453)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJPropostas_Le(ByVal objPRJPropostas As ClassPRJPropostas, Optional ByVal bLeItens As Boolean = True, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJPropostas As typePRJPropostas

On Error GoTo Erro_PRJPropostas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185857

    'Alocação de espaço no buffer
    tPRJPropostas.sCodigo = String(STRING_PROPOSTAPRJ_CODIGO, 0)
    tPRJPropostas.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPRJPropostas.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

    'Le a tabelaPRJPropostas
    lErro = Comando_Executar(lComando, "SELECT NaturezaOP, NumIntDoc, NumIntDocPRJ, Codigo, Data, Observacao, Cliente, " & _
                "FilialCliente, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, " & _
                "CustoInformado, CustoCalculado, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco, NumIntDocContrato FROM PRJPropostas WHERE NumIntDocPRJ= ?  AND Codigo= ? ", _
                tPRJPropostas.sNaturezaOp, tPRJPropostas.lNumIntDoc, tPRJPropostas.lNumIntDocPRJ, tPRJPropostas.sCodigo, tPRJPropostas.dtData, _
                tPRJPropostas.sObservacao, tPRJPropostas.lCliente, tPRJPropostas.iFilialCliente, tPRJPropostas.dValorTotal, tPRJPropostas.dValorProdutos, _
                tPRJPropostas.dValorFrete, tPRJPropostas.dValorDesconto, tPRJPropostas.dValorSeguro, tPRJPropostas.dValorOutrasDespesas, tPRJPropostas.dCustoInformado, _
                tPRJPropostas.dCustoCalculado, tPRJPropostas.iExibirProdutos, tPRJPropostas.iExibirCustoCalc, tPRJPropostas.iExibirCustoInfo, tPRJPropostas.iExibirPreco, tPRJPropostas.lNumIntDocContrato, _
                objPRJPropostas.lNumIntDocPRJ, objPRJPropostas.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185858

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185859

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJPropostas.lNumIntDoc = tPRJPropostas.lNumIntDoc
    objPRJPropostas.lNumIntDocPRJ = tPRJPropostas.lNumIntDocPRJ
    objPRJPropostas.sCodigo = tPRJPropostas.sCodigo
    objPRJPropostas.dtData = tPRJPropostas.dtData
    objPRJPropostas.sObservacao = tPRJPropostas.sObservacao
    objPRJPropostas.lCliente = tPRJPropostas.lCliente
    objPRJPropostas.iFilialCliente = tPRJPropostas.iFilialCliente
    objPRJPropostas.dValorTotal = tPRJPropostas.dValorTotal
    objPRJPropostas.dValorProdutos = tPRJPropostas.dValorProdutos
    objPRJPropostas.dValorFrete = tPRJPropostas.dValorFrete
    objPRJPropostas.dValorDesconto = tPRJPropostas.dValorDesconto
    objPRJPropostas.dValorSeguro = tPRJPropostas.dValorSeguro
    objPRJPropostas.dValorOutrasDespesas = tPRJPropostas.dValorOutrasDespesas
    objPRJPropostas.dCustoInformado = tPRJPropostas.dCustoInformado
    objPRJPropostas.dCustoCalculado = tPRJPropostas.dCustoCalculado
    objPRJPropostas.iExibirCustoCalc = tPRJPropostas.iExibirCustoCalc
    objPRJPropostas.iExibirCustoInfo = tPRJPropostas.iExibirCustoInfo
    objPRJPropostas.iExibirPreco = tPRJPropostas.iExibirPreco
    objPRJPropostas.iExibirProdutos = tPRJPropostas.iExibirProdutos
    objPRJPropostas.sNaturezaOp = tPRJPropostas.sNaturezaOp
    objPRJPropostas.lNumIntDocContrato = tPRJPropostas.lNumIntDocContrato
    
    If bLeItens Then
    
        lErro = CF("PRJPropostaItens_Le", objPRJPropostas, bLeTrib)
        If lErro <> SUCESSO Then gError 187515
    
        lErro = CF("PRJPropostaEtapa_Le", objPRJPropostas)
        If lErro <> SUCESSO Then gError 187525
        
    End If
    
    If bLeTrib Then
    
        lErro = CF("PRJPropostas_Le_Tributacao", objPRJPropostas)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187599
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJPropostas_Le = SUCESSO

    Exit Function

Erro_PRJPropostas_Le:

    PRJPropostas_Le = gErr

    Select Case gErr

        Case 185857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185858, 185859
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)
            
        Case 187515, 187525, 187599

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185860)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJPropostas_Le_NumIntDoc(ByVal objPRJPropostas As ClassPRJPropostas, Optional ByVal bLeItens As Boolean = True, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJPropostas As typePRJPropostas

On Error GoTo Erro_PRJPropostas_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187997

    'Alocação de espaço no buffer
    tPRJPropostas.sCodigo = String(STRING_PROPOSTAPRJ_CODIGO, 0)
    tPRJPropostas.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPRJPropostas.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

    'Le a tabelaPRJPropostas
    lErro = Comando_Executar(lComando, "SELECT NaturezaOP, NumIntDoc, NumIntDocPRJ, Codigo, Data, Observacao, Cliente, " & _
                "FilialCliente, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, " & _
                "CustoInformado, CustoCalculado, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco, NumIntDocContrato FROM PRJPropostas WHERE NumIntDoc= ? ", _
                tPRJPropostas.sNaturezaOp, tPRJPropostas.lNumIntDoc, tPRJPropostas.lNumIntDocPRJ, tPRJPropostas.sCodigo, tPRJPropostas.dtData, _
                tPRJPropostas.sObservacao, tPRJPropostas.lCliente, tPRJPropostas.iFilialCliente, tPRJPropostas.dValorTotal, tPRJPropostas.dValorProdutos, _
                tPRJPropostas.dValorFrete, tPRJPropostas.dValorDesconto, tPRJPropostas.dValorSeguro, tPRJPropostas.dValorOutrasDespesas, tPRJPropostas.dCustoInformado, _
                tPRJPropostas.dCustoCalculado, tPRJPropostas.iExibirProdutos, tPRJPropostas.iExibirCustoCalc, tPRJPropostas.iExibirCustoInfo, tPRJPropostas.iExibirPreco, tPRJPropostas.lNumIntDocContrato, _
                objPRJPropostas.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187998

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187999

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJPropostas.lNumIntDoc = tPRJPropostas.lNumIntDoc
    objPRJPropostas.lNumIntDocPRJ = tPRJPropostas.lNumIntDocPRJ
    objPRJPropostas.sCodigo = tPRJPropostas.sCodigo
    objPRJPropostas.dtData = tPRJPropostas.dtData
    objPRJPropostas.sObservacao = tPRJPropostas.sObservacao
    objPRJPropostas.lCliente = tPRJPropostas.lCliente
    objPRJPropostas.iFilialCliente = tPRJPropostas.iFilialCliente
    objPRJPropostas.dValorTotal = tPRJPropostas.dValorTotal
    objPRJPropostas.dValorProdutos = tPRJPropostas.dValorProdutos
    objPRJPropostas.dValorFrete = tPRJPropostas.dValorFrete
    objPRJPropostas.dValorDesconto = tPRJPropostas.dValorDesconto
    objPRJPropostas.dValorSeguro = tPRJPropostas.dValorSeguro
    objPRJPropostas.dValorOutrasDespesas = tPRJPropostas.dValorOutrasDespesas
    objPRJPropostas.dCustoInformado = tPRJPropostas.dCustoInformado
    objPRJPropostas.dCustoCalculado = tPRJPropostas.dCustoCalculado
    objPRJPropostas.iExibirCustoCalc = tPRJPropostas.iExibirCustoCalc
    objPRJPropostas.iExibirCustoInfo = tPRJPropostas.iExibirCustoInfo
    objPRJPropostas.iExibirPreco = tPRJPropostas.iExibirPreco
    objPRJPropostas.iExibirProdutos = tPRJPropostas.iExibirProdutos
    objPRJPropostas.sNaturezaOp = tPRJPropostas.sNaturezaOp
    objPRJPropostas.lNumIntDocContrato = tPRJPropostas.lNumIntDocContrato
    
    If bLeItens Then
    
        lErro = CF("PRJPropostaItens_Le", objPRJPropostas, bLeTrib)
        If lErro <> SUCESSO Then gError 189000
    
        lErro = CF("PRJPropostaEtapa_Le", objPRJPropostas)
        If lErro <> SUCESSO Then gError 189001
        
    End If
    
    If bLeTrib Then
    
        lErro = CF("PRJPropostas_Le_NumIntDoc_Tributacao", objPRJPropostas)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189002
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJPropostas_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_PRJPropostas_Le_NumIntDoc:

    PRJPropostas_Le_NumIntDoc = gErr

    Select Case gErr

        Case 187997
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187998, 187999
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)
            
        Case 189000, 189001, 189002

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189003)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJRecebPagto_Le_Proposta(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJRecebPagto As typePRJRecebPagto

On Error GoTo Erro_PRJRecebPagto_Le_Proposta

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187488

    'Le a tabelaPRJRecebPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Tipo, Numero, Valor, CliForn, " & _
                "Filial, NumIntDocProposta,NumIntDocContrato, IncluiCFF, FilialEmpresa FROM PRJRecebPagto WHERE NumIntDocPRJ= ?  AND NumIntDocProposta= ? AND Tipo = ?", _
                tPRJRecebPagto.lNumIntDoc, tPRJRecebPagto.lNumIntDocPRJ, tPRJRecebPagto.iTipo, tPRJRecebPagto.lNumero, _
                tPRJRecebPagto.dValor, tPRJRecebPagto.lCliForn, tPRJRecebPagto.iFilial, tPRJRecebPagto.lNumIntDocProposta, tPRJRecebPagto.iIncluiCFF, _
                tPRJRecebPagto.iFilialEmpresa, _
                objPRJRecebPagto.lNumIntDocPRJ, objPRJRecebPagto.lNumIntDocProposta, tPRJRecebPagto.lNumIntDocContrato, objPRJRecebPagto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 187489

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187490

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJRecebPagto.lNumIntDoc = tPRJRecebPagto.lNumIntDoc
    objPRJRecebPagto.lNumIntDocPRJ = tPRJRecebPagto.lNumIntDocPRJ
    objPRJRecebPagto.iTipo = tPRJRecebPagto.iTipo
    objPRJRecebPagto.lNumero = tPRJRecebPagto.lNumero
    objPRJRecebPagto.dValor = tPRJRecebPagto.dValor
    objPRJRecebPagto.lCliForn = tPRJRecebPagto.lCliForn
    objPRJRecebPagto.iFilial = tPRJRecebPagto.iFilial
    objPRJRecebPagto.lNumIntDocProposta = tPRJRecebPagto.lNumIntDocProposta
    objPRJRecebPagto.iIncluiCFF = tPRJRecebPagto.iIncluiCFF
    objPRJRecebPagto.iFilialEmpresa = tPRJRecebPagto.iFilialEmpresa
    objPRJRecebPagto.lNumIntDocContrato = tPRJRecebPagto.lNumIntDocContrato

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJRecebPagto_Le_Proposta = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Le_Proposta:

    PRJRecebPagto_Le_Proposta = gErr

    Select Case gErr

        Case 187488
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187489, 187490
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187491)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJPropostaItens_Le(ByVal objPRJProposta As ClassPRJPropostas, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJPropostaItem As typePRJPropostaItem
Dim objPRJPropostaItem As ClassPRJPropostaItem

On Error GoTo Erro_PRJPropostaItens_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187510

    'Alocação de espaço no buffer
    tPRJPropostaItem.sProduto = String(STRING_PRODUTO, 0)
    tPRJPropostaItem.sDescProd = String(STRING_DESCRICAO_ITEM, 0)
    tPRJPropostaItem.sDescEtapa = String(STRING_DESCRICAO_ITEM, 0)
    tPRJPropostaItem.sUM = String(STRING_UM_SIGLA, 0)
    tPRJPropostaItem.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaPRJPropostaItens
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Item, NumIntDocProposta, Produto, DescProd, " & _
                "NumIntDocEtapa, DescEtapa, Quantidade, PrecoUnitario, PrecoTotal, UnidadeMed, ValorDesconto, " & _
                "DataEntrega, Observacao FROM PRJPropostaItens WHERE NumIntDocProposta= ? ", _
                tPRJPropostaItem.lNumIntDoc, tPRJPropostaItem.iFilialEmpresa, tPRJPropostaItem.iItem, tPRJPropostaItem.lNumIntDocProposta, _
                tPRJPropostaItem.sProduto, tPRJPropostaItem.sDescProd, tPRJPropostaItem.lNumIntDocEtapa, tPRJPropostaItem.sDescEtapa, tPRJPropostaItem.dQuantidade, _
                tPRJPropostaItem.dPrecoUnitario, tPRJPropostaItem.dPrecoTotal, tPRJPropostaItem.sUM, tPRJPropostaItem.dValorDesconto, tPRJPropostaItem.dtDataEntrega, _
                tPRJPropostaItem.sObservacao, _
                objPRJProposta.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187511

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187512

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJPropostaItem = New ClassPRJPropostaItem

        objPRJPropostaItem.lNumIntDoc = tPRJPropostaItem.lNumIntDoc
        objPRJPropostaItem.iFilialEmpresa = tPRJPropostaItem.iFilialEmpresa
        objPRJPropostaItem.iItem = tPRJPropostaItem.iItem
        objPRJPropostaItem.lNumIntDocProposta = tPRJPropostaItem.lNumIntDocProposta
        objPRJPropostaItem.sProduto = tPRJPropostaItem.sProduto
        objPRJPropostaItem.sDescProd = tPRJPropostaItem.sDescProd
        objPRJPropostaItem.lNumIntDocEtapa = tPRJPropostaItem.lNumIntDocEtapa
        objPRJPropostaItem.sDescEtapa = tPRJPropostaItem.sDescEtapa
        objPRJPropostaItem.dQuantidade = tPRJPropostaItem.dQuantidade
        objPRJPropostaItem.dPrecoUnitario = tPRJPropostaItem.dPrecoUnitario
        objPRJPropostaItem.dPrecoTotal = tPRJPropostaItem.dPrecoTotal
        objPRJPropostaItem.sUM = tPRJPropostaItem.sUM
        objPRJPropostaItem.dValorDesconto = tPRJPropostaItem.dValorDesconto
        objPRJPropostaItem.dtDataEntrega = tPRJPropostaItem.dtDataEntrega
        objPRJPropostaItem.sObservacao = tPRJPropostaItem.sObservacao
        
        objPRJProposta.colItens.Add objPRJPropostaItem
        
        If bLeTrib Then
        
            lErro = PRJPropostaItem_Le_Tributacao(objPRJPropostaItem)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187605
        
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187513
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJPropostaItens_Le = SUCESSO

    Exit Function

Erro_PRJPropostaItens_Le:

    PRJPropostaItens_Le = gErr

    Select Case gErr

        Case 187510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187511, 187512, 187513
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAITENS", gErr)
            
        Case 187605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187514)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJPropostaEtapa_Le(ByVal objPRJProposta As ClassPRJPropostas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJPropostaEtapa As typePRJPropostaEtapa
Dim objPRJPropostaEtapa As ClassPRJPropostaEtapa

On Error GoTo Erro_PRJPropostaEtapa_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187516

    tPRJPropostaEtapa.sDescricao = String(STRING_PROPOSTAETAPA_DESCRICAO, 0)
    tPRJPropostaEtapa.sObservacao = String(STRING_PROPOSTAETAPA_OBSERVACAO, 0)

    'Le a tabelaPRJPropostaEtapa
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocProposta, NumIntDocEtapa, NumIntDocEtapaItemProd, CustoInformado, Preco, Selecionado,Imprimir,Descricao,Observacao FROM PRJPropostaEtapa WHERE NumIntDocProposta= ? ", _
                tPRJPropostaEtapa.lNumIntDoc, tPRJPropostaEtapa.lNumIntDocProposta, tPRJPropostaEtapa.lNumIntDocEtapa, tPRJPropostaEtapa.lNumIntDocEtapaItemProd, _
                tPRJPropostaEtapa.dCustoInformado, tPRJPropostaEtapa.dPreco, tPRJPropostaEtapa.iSelecionado, tPRJPropostaEtapa.iImprimir, tPRJPropostaEtapa.sDescricao, tPRJPropostaEtapa.sObservacao, _
                objPRJProposta.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187517

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187518

    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objPRJPropostaEtapa = New ClassPRJPropostaEtapa

        objPRJPropostaEtapa.lNumIntDoc = tPRJPropostaEtapa.lNumIntDoc
        objPRJPropostaEtapa.lNumIntDocProposta = tPRJPropostaEtapa.lNumIntDocProposta
        objPRJPropostaEtapa.lNumIntDocEtapa = tPRJPropostaEtapa.lNumIntDocEtapa
        objPRJPropostaEtapa.lNumIntDocEtapaItemProd = tPRJPropostaEtapa.lNumIntDocEtapaItemProd
        objPRJPropostaEtapa.iSelecionado = tPRJPropostaEtapa.iSelecionado
        objPRJPropostaEtapa.dCustoInformado = tPRJPropostaEtapa.dCustoInformado
        objPRJPropostaEtapa.dPreco = tPRJPropostaEtapa.dPreco
        objPRJPropostaEtapa.iImprimir = tPRJPropostaEtapa.iImprimir
        objPRJPropostaEtapa.sDescricao = tPRJPropostaEtapa.sDescricao
        objPRJPropostaEtapa.sObservacao = tPRJPropostaEtapa.sObservacao
        
        objPRJProposta.colEtapas.Add objPRJPropostaEtapa

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187519

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJPropostaEtapa_Le = SUCESSO

    Exit Function

Erro_PRJPropostaEtapa_Le:

    PRJPropostaEtapa_Le = gErr

    Select Case gErr

        Case 187516
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187517, 187518, 187519
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAETAPA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187520)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJPropostas_Le_Tributacao(ByVal objProposta As ClassPRJPropostas) As Long
'Le a Tributacao da proposta de Venda Passado como parametro.

Dim lErro As Long
'Dim lComando As Long
'Dim tTributacaoPRJProp As typeTributacaoPRJProp
'Dim objTributacaoPRJProp As New ClassTributacaoPRJProp

On Error GoTo Erro_PRJPropostas_Le_Tributacao

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 187585
'
'    With tTributacaoPRJProp
'
'        .sCodProposta = String(STRING_PROPOSTAPRJ_CODIGO, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, CodProposta, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoPRJProp WHERE FilialEmpresa = ? AND CodProposta = ?", _
'            .iFilialEmpresa, .sCodProposta, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual, _
'            objProposta.iFilialEmpresa, objProposta.sCodigo)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187586
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187587
'
'    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    With objTributacaoPRJProp
'        .dICMSBase = tTributacaoPRJProp.dICMSBase
'        .dICMSSubstBase = tTributacaoPRJProp.dICMSSubstBase
'        .dICMSSubstValor = tTributacaoPRJProp.dICMSSubstValor
'        .dICMSValor = tTributacaoPRJProp.dICMSValor
'        .dIPIBase = tTributacaoPRJProp.dIPIBase
'        .dIPIValor = tTributacaoPRJProp.dIPIValor
'        .dIRRFAliquota = tTributacaoPRJProp.dIRRFAliquota
'        .dIRRFBase = tTributacaoPRJProp.dIRRFBase
'        .dIRRFValor = tTributacaoPRJProp.dIRRFValor
'        .dISSAliquota = tTributacaoPRJProp.dISSAliquota
'        .dISSBase = tTributacaoPRJProp.dISSBase
'        .dISSValor = tTributacaoPRJProp.dISSValor
'        .iFilialEmpresa = tTributacaoPRJProp.iFilialEmpresa
'        .iICMSBaseManual = tTributacaoPRJProp.iICMSBaseManual
'        .iICMSSubstBaseManual = tTributacaoPRJProp.iICMSSubstBaseManual
'        .iICMSSubstValorManual = tTributacaoPRJProp.iICMSSubstValorManual
'        .iICMSValorManual = tTributacaoPRJProp.iICMSValorManual
'        .iIPIBaseManual = tTributacaoPRJProp.iIPIBaseManual
'        .iIPIValorManual = tTributacaoPRJProp.iIPIValorManual
'        .iIRRFAliquotaManual = tTributacaoPRJProp.iIRRFAliquotaManual
'        .iIRRFValorManual = tTributacaoPRJProp.iIRRFValorManual
'        .iISSIncluso = tTributacaoPRJProp.iISSIncluso
'        .iISSInclusoManual = tTributacaoPRJProp.iISSValorManual
'        .iISSValorManual = tTributacaoPRJProp.iISSValorManual
'        .iTaxacaoAutomatica = tTributacaoPRJProp.iTaxacaoAutomatica
'        .iTipoTributacao = tTributacaoPRJProp.iTipoTributacao
'        .iTipoTributacaoManual = tTributacaoPRJProp.iTipoTributacaoManual
'        .sCodProposta = tTributacaoPRJProp.sCodProposta
'        .iPISRetidoManual = tTributacaoPRJProp.iPISRetidoManual
'        .iISSRetidoManual = tTributacaoPRJProp.iISSRetidoManual
'        .iCOFINSRetidoManual = tTributacaoPRJProp.iCOFINSRetidoManual
'        .iCSLLRetidoManual = tTributacaoPRJProp.iCSLLRetidoManual
'        .dPISRetido = tTributacaoPRJProp.dPISRetido
'        .dISSRetido = tTributacaoPRJProp.dISSRetido
'        .dCOFINSRetido = tTributacaoPRJProp.dCOFINSRetido
'        .dCSLLRetido = tTributacaoPRJProp.dCSLLRetido
'
'        .dPISCredito = tTributacaoPRJProp.dPISCredito
'        .iPISCreditoManual = tTributacaoPRJProp.iPISCreditoManual
'        .dCOFINSCredito = tTributacaoPRJProp.dCOFINSCredito
'        .iCOFINSCreditoManual = tTributacaoPRJProp.iCOFINSCreditoManual
'        .dIPICredito = tTributacaoPRJProp.dIPICredito
'        .iIPICreditoManual = tTributacaoPRJProp.iIPICreditoManual
'        .dICMSCredito = tTributacaoPRJProp.dICMSCredito
'        .iICMSCreditoManual = tTributacaoPRJProp.iICMSCreditoManual
'        .dValorINSS = tTributacaoPRJProp.dValorINSS
'        .iINSSValorManual = tTributacaoPRJProp.iINSSValorManual
'        .iINSSRetido = tTributacaoPRJProp.iINSSRetido
'        .iINSSRetidoManual = tTributacaoPRJProp.iINSSRetidoManual
'        .dINSSBase = tTributacaoPRJProp.dINSSBase
'        .iINSSBaseManual = tTributacaoPRJProp.iINSSBaseManual
'        .dINSSDeducoes = tTributacaoPRJProp.dINSSDeducoes
'        .iINSSDeducoesManual = tTributacaoPRJProp.iINSSDeducoesManual
'    End With
'
'    Set objProposta.objTributacaoPRJProp = objTributacaoPRJProp
'
'    lErro = CF("PRJPropostas_Le_Complementos", objProposta)
'    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187589
'
'    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187590
'
'    Call Comando_Fechar(lComando)

    objProposta.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_PROP
    objProposta.objTributacao.lNumIntDoc = objProposta.lNumIntDoc

    lErro = CF("TributacaoDoc_Le", objProposta.objTributacao)
    If lErro <> SUCESSO Then gError 187589
    
    PRJPropostas_Le_Tributacao = SUCESSO

    Exit Function

Erro_PRJPropostas_Le_Tributacao:

    PRJPropostas_Le_Tributacao = gErr

    Select Case gErr

'        Case 187585
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 187586, 187587
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBPROPOSTA", gErr)
'
'        Case ERRO_LEITURA_SEM_DADOS
        
        Case 187589
'
'        Case 187590
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_PROPOSTA_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187591)

    End Select

    'Call Comando_Fechar(lComando)

    Exit Function

End Function

'Public Function PRJPropostas_Le_Complementos(ByVal objProposta As ClassPRJPropostas) As Long
'
'Dim lComando As Long, objTributacaoPRJProp As ClassTributacaoPRJProp
'Dim tTributacaoComplPRJProp As typeTributacaoComplPV, objTributacaoItemPRJProp As ClassTributacaoItemPV
'Dim lErro As Long
'
'On Error GoTo Erro_PRJPropostas_Le_Complementos
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 187592
'
'    Set objTributacaoPRJProp = objProposta.objTributacaoPRJProp
'
'    tTributacaoComplPRJProp.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'    'ler dados de frete, seguro, despesas, descontos da tabela TributacaoComplPRJProp
'    lErro = Comando_Executar(lComando, "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, " & _
'        "IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, " & _
'        "ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoComplPRJProp WHERE FilialEmpresa = ? AND CodProposta = ?", _
'        tTributacaoComplPRJProp.iItem, tTributacaoComplPRJProp.sNaturezaOp, tTributacaoComplPRJProp.iNaturezaOpManual, tTributacaoComplPRJProp.iTipoTributacao, tTributacaoComplPRJProp.iTipoTributacaoManual, tTributacaoComplPRJProp.iIPITipo, tTributacaoComplPRJProp.iIPITipoManual, tTributacaoComplPRJProp.dIPIBaseCalculo, tTributacaoComplPRJProp.iIPIBaseManual, tTributacaoComplPRJProp.dIPIPercRedBase, _
'        tTributacaoComplPRJProp.iIPIPercRedBaseManual, tTributacaoComplPRJProp.dIPIAliquota, tTributacaoComplPRJProp.iIPIAliquotaManual, tTributacaoComplPRJProp.dIPIValor, tTributacaoComplPRJProp.iIPIValorManual, tTributacaoComplPRJProp.iICMSTipo, tTributacaoComplPRJProp.iICMSTipoManual, tTributacaoComplPRJProp.dICMSBase, tTributacaoComplPRJProp.iICMSBaseManual, tTributacaoComplPRJProp.dICMSPercRedBase, _
'        tTributacaoComplPRJProp.iICMSPercRedBaseManual, tTributacaoComplPRJProp.dICMSAliquota, tTributacaoComplPRJProp.iICMSAliquotaManual, tTributacaoComplPRJProp.dICMSValor, tTributacaoComplPRJProp.iICMSValorManual, tTributacaoComplPRJProp.dICMSSubstBase, tTributacaoComplPRJProp.iICMSSubstBaseManual, _
'        tTributacaoComplPRJProp.dICMSSubstAliquota, tTributacaoComplPRJProp.iICMSSubstAliquotaManual, tTributacaoComplPRJProp.dICMSSubstValor, tTributacaoComplPRJProp.iICMSSubstValorManual, tTributacaoComplPRJProp.dPISCredito, tTributacaoComplPRJProp.dCOFINSCredito, tTributacaoComplPRJProp.dIPICredito, tTributacaoComplPRJProp.dICMSCredito, objProposta.iFilialEmpresa, objProposta.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 187593
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemPRJProp = New ClassTributacaoItemPV
'
'        With objTributacaoItemPRJProp
'
'            .lNumIntDoc = tTributacaoComplPRJProp.iItem
'            .sNaturezaOp = tTributacaoComplPRJProp.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoComplPRJProp.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoComplPRJProp.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoComplPRJProp.iTipoTributacaoManual
'            .iIPITipo = tTributacaoComplPRJProp.iIPITipo
'            .iIPITipoManual = tTributacaoComplPRJProp.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoComplPRJProp.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoComplPRJProp.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoComplPRJProp.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoComplPRJProp.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoComplPRJProp.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoComplPRJProp.iIPIAliquotaManual
'            .dIPIValor = tTributacaoComplPRJProp.dIPIValor
'            .iIPIValorManual = tTributacaoComplPRJProp.iIPIValorManual
'            .iICMSTipo = tTributacaoComplPRJProp.iICMSTipo
'            .iICMSTipoManual = tTributacaoComplPRJProp.iICMSTipoManual
'            .dICMSBase = tTributacaoComplPRJProp.dICMSBase
'            .iICMSBaseManual = tTributacaoComplPRJProp.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoComplPRJProp.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoComplPRJProp.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoComplPRJProp.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoComplPRJProp.iICMSAliquotaManual
'            .dICMSValor = tTributacaoComplPRJProp.dICMSValor
'            .iICMSValorManual = tTributacaoComplPRJProp.iICMSValorManual
'            .dICMSSubstBase = tTributacaoComplPRJProp.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoComplPRJProp.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoComplPRJProp.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoComplPRJProp.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoComplPRJProp.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoComplPRJProp.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoComplPRJProp.dICMSCredito
'            .dIPICredito = tTributacaoComplPRJProp.dIPICredito
'            .dCOFINSCredito = tTributacaoComplPRJProp.dCOFINSCredito
'            .dPISCredito = tTributacaoComplPRJProp.dPISCredito
'
'        End With
'
'        Select Case tTributacaoComplPRJProp.iItem
'
'            Case ITEM_TIPO_FRETE
'                Set objTributacaoPRJProp.objTributacaoFrete = objTributacaoItemPRJProp
'
'            Case ITEM_TIPO_SEGURO
'                Set objTributacaoPRJProp.objTributacaoSeguro = objTributacaoItemPRJProp
'
'            Case ITEM_TIPO_DESCONTO
'                Set objTributacaoPRJProp.objTributacaoDesconto = objTributacaoItemPRJProp
'
'            Case ITEM_TIPO_OUTRAS_DESP
'                Set objTributacaoPRJProp.objTributacaoOutras = objTributacaoItemPRJProp
'
'            Case Else 'item invalido
'                gError 187596
'
'        End Select
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187597
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    PRJPropostas_Le_Complementos = SUCESSO
'
'    Exit Function
'
'Erro_PRJPropostas_Le_Complementos:
'
'    PRJPropostas_Le_Complementos = gErr
'
'    Select Case gErr
'
'        Case 187592
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 187593, ERRO_LEITURA_SEM_DADOS, 187597
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJPROPOSTAS", gErr)
'
'        Case 187596
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPL_PROPOSTA_TIPO_INVALIDO", gErr)
'
'        Case ERRO_LEITURA_SEM_DADOS
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187598)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Public Function PRJPropostaItem_Le_Tributacao(ByVal objPropostaItem As ClassPRJPropostaItem) As Long

Dim lErro As Long
'Dim lComando As Long
'Dim tTributacaoItemPRJProp As typeTributacaoItemPV
'Dim objTributacaoItemPRJProp As ClassTributacaoPRJPropItem

On Error GoTo Erro_PRJPropostaItem_Le_Tributacao

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 187600
'
'    With tTributacaoItemPRJProp
'
'        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, " & _
'            "ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito " & _
'            "FROM TributacaoPRJPropItem WHERE NumIntDoc = ? ORDER BY NumIntDoc", _
'            .sNaturezaOp, .iNaturezaOpManual, .iTipoTributacao, .iTipoTributacaoManual, .iIPITipo, .iIPITipoManual, .dIPIBaseCalculo, .iIPIBaseManual, .dIPIPercRedBase, .iIPIPercRedBaseManual, .dIPIAliquota, .iIPIAliquotaManual, .dIPIValor, _
'            .iIPIValorManual, .iICMSTipo, .iICMSTipoManual, .dICMSBase, .iICMSBaseManual, .dICMSPercRedBase, .iICMSPercRedBaseManual, .dICMSAliquota, .iICMSAliquotaManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstAliquota, .iICMSSubstAliquotaManual, .dICMSSubstValor, .iICMSSubstValorManual, .dPISCredito, .dCOFINSCredito, .dIPICredito, .dICMSCredito, objPropostaItem.lNumIntDoc)
'        If lErro <> AD_SQL_SUCESSO Then gError 187601
'
'    End With
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187602
'
'    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemPRJProp = New ClassTributacaoPRJPropItem
'
'        With objTributacaoItemPRJProp
'
'            .lNumIntDoc = tTributacaoItemPRJProp.lNumIntDoc
'            .sNaturezaOp = tTributacaoItemPRJProp.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoItemPRJProp.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoItemPRJProp.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoItemPRJProp.iTipoTributacaoManual
'            .iIPITipo = tTributacaoItemPRJProp.iIPITipo
'            .iIPITipoManual = tTributacaoItemPRJProp.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoItemPRJProp.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoItemPRJProp.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoItemPRJProp.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoItemPRJProp.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoItemPRJProp.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoItemPRJProp.iIPIAliquotaManual
'            .dIPIValor = tTributacaoItemPRJProp.dIPIValor
'            .iIPIValorManual = tTributacaoItemPRJProp.iIPIValorManual
'            .iICMSTipo = tTributacaoItemPRJProp.iICMSTipo
'            .iICMSTipoManual = tTributacaoItemPRJProp.iICMSTipoManual
'            .dICMSBase = tTributacaoItemPRJProp.dICMSBase
'            .iICMSBaseManual = tTributacaoItemPRJProp.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoItemPRJProp.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoItemPRJProp.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoItemPRJProp.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoItemPRJProp.iICMSAliquotaManual
'            .dICMSValor = tTributacaoItemPRJProp.dICMSValor
'            .iICMSValorManual = tTributacaoItemPRJProp.iICMSValorManual
'            .dICMSSubstBase = tTributacaoItemPRJProp.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoItemPRJProp.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoItemPRJProp.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoItemPRJProp.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoItemPRJProp.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoItemPRJProp.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoItemPRJProp.dICMSCredito
'            .dIPICredito = tTributacaoItemPRJProp.dIPICredito
'            .dCOFINSCredito = tTributacaoItemPRJProp.dCOFINSCredito
'            .dPISCredito = tTributacaoItemPRJProp.dPISCredito
'        End With
'
'        Set objPropostaItem.objTributacaoPRJPropItem = objTributacaoItemPRJProp
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187603
'
'    Loop
'
'    Call Comando_Fechar(lComando)

    objPropostaItem.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_PROP
    objPropostaItem.objTributacao.lNumIntDocItem = objPropostaItem.lNumIntDoc
    
    lErro = CF("TributacaoDocItem_Le", objPropostaItem.objTributacao)
    If lErro <> SUCESSO Then gError 187601
    
    PRJPropostaItem_Le_Tributacao = SUCESSO

    Exit Function

Erro_PRJPropostaItem_Le_Tributacao:

    PRJPropostaItem_Le_Tributacao = gErr

    Select Case gErr
'
'        Case 187600
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
        Case 187601 ', 187602, 187603
'
'        Case ERRO_LEITURA_SEM_DADOS
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187604)

    End Select

    'Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function MnemonicoMalaDireta_Le(ByVal objMnemonicoMalaDireta As ClassMnemonicoMalaDireta) As Long

Dim lErro As Long
Dim lComando As Long
Dim tMnemonicoMalaDireta As typeMnemonicoMalaDireta

On Error GoTo Erro_MnemonicoMalaDireta_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187591

    'Alocação de espaço no buffer
    tMnemonicoMalaDireta.sMnemonico = String(STRING_MNEMONICO_MALADIRETA, 0)
    tMnemonicoMalaDireta.sDescricao = String(STRING_MNEMONICO_MALADIRETA_DESCRICAO, 0)
    tMnemonicoMalaDireta.sNomeCampoObj = String(STRING_MNEMONICO_MALADIRETA_NOMECAMPOOBJ, 0)

    'Le a tabelaMnemonicoMalaDireta
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Descricao, TipoObj, NomeCampoObj, Tipo FROM MnemonicoMalaDireta WHERE Mnemonico= ? AND Tipo = ?", _
                tMnemonicoMalaDireta.sMnemonico, tMnemonicoMalaDireta.sDescricao, tMnemonicoMalaDireta.iTipoObj, tMnemonicoMalaDireta.sNomeCampoObj, _
                tMnemonicoMalaDireta.iTipo, _
                objMnemonicoMalaDireta.sMnemonico, objMnemonicoMalaDireta.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 187592

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187593

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objMnemonicoMalaDireta.sMnemonico = tMnemonicoMalaDireta.sMnemonico
    objMnemonicoMalaDireta.sDescricao = tMnemonicoMalaDireta.sDescricao
    objMnemonicoMalaDireta.iTipoObj = tMnemonicoMalaDireta.iTipoObj
    objMnemonicoMalaDireta.sNomeCampoObj = tMnemonicoMalaDireta.sNomeCampoObj
    objMnemonicoMalaDireta.iTipo = tMnemonicoMalaDireta.iTipo

    'Fecha Comando
    Call Comando_Fechar(lComando)

    MnemonicoMalaDireta_Le = SUCESSO

    Exit Function

Erro_MnemonicoMalaDireta_Le:

    MnemonicoMalaDireta_Le = gErr

    Select Case gErr

        Case 187591
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187592, 187593
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOMALADIRETA", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187594)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Projeto_ExisteProjeto(iExiste As Integer) As Long
'verifica se tem algum Projeto cadastrado.
'se existir iExiste = MARCADO , caso contrario iExiste = DESMARCADO

Dim lComando As Long
Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_Projeto_ExisteProjeto

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189154

    'Verifica se o Projeto existe
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM Projetos", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189155

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189156
        
    'Ccl está cadastrado
    If lErro = AD_SQL_SUCESSO Then
        iExiste = MARCADO
    Else
        iExiste = DESMARCADO
    End If

    Call Comando_Fechar(lComando)

    Projeto_ExisteProjeto = SUCESSO

    Exit Function

Erro_Projeto_ExisteProjeto:

    Projeto_ExisteProjeto = gErr
    
    Select Case gErr

        Case 189154
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189155, 189156
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189157)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RefEtapa_ExisteRefEtapa(iExiste As Integer) As Long
'verifica se tem algum RefEtapa cadastrado.
'se existir iExiste = MARCADO , caso contrario iExiste = DESMARCADO

Dim lComando As Long
Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_RefEtapa_ExisteRefEtapa

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189159

    'Verifica se o RefEtapa existe
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM PRJEtapas WHERE Referencia <> ?", lNumIntDoc, "")
    If lErro <> AD_SQL_SUCESSO Then gError 189160

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189161
        
    'Ccl está cadastrado
    If lErro = AD_SQL_SUCESSO Then
        iExiste = MARCADO
    Else
        iExiste = DESMARCADO
    End If

    Call Comando_Fechar(lComando)

    RefEtapa_ExisteRefEtapa = SUCESSO

    Exit Function

Erro_RefEtapa_ExisteRefEtapa:

    RefEtapa_ExisteRefEtapa = gErr
    
    Select Case gErr

        Case 189159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189160, 189161
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RefEtapaS", gErr)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189162)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJEtapas_Le_TiposCalculados(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim colPRJCR As New Collection
Dim colPRJCRUsados As New Collection
Dim objPRJCR As ClassPRJCR
Dim objPRJInfo As ClassProjetoInfo
Dim colPRJInfo As New Collection
Dim bAchou As Boolean
Dim objOP As ClassOrdemDeProducao
Dim colOP As New Collection
Dim objItemOP As ClassItemOP
Dim sProduto As String
Dim iItem As Integer
Dim lNumIntDocItemOP As Long
Dim dQuantidade As Double
Dim iSeqMaq As Integer
Dim iSeqMO As Integer
Dim iSeqMP As Integer
Dim objEtapaAux As New ClassPRJEtapas
Dim objEtapaMaq As ClassPRJEtapaMaquinas, objEtapaMO As ClassPRJEtapaMO, objEtapaMP As ClassPRJEtapaMateriais
Dim objEtapaMaqAux As ClassPRJEtapaMaquinas, objEtapaMOAux As ClassPRJEtapaMO, objEtapaMPAux As ClassPRJEtapaMateriais
Dim colEtapaMaqAux As New Collection, colEtapaMOAux As New Collection, colEtapaMPAux As New Collection
Dim colSaida As Collection
Dim colCampos As Collection
Dim sUM As String
Dim dQtd As Double
Dim dCusto As Double
Dim sObs As String
Dim dtData As Date
Dim sDescricao As String
Dim iMaquina As Integer
Dim iMO As Integer
Dim iQtd As Integer
Dim dHoras As Double
Dim lNumIntDocMaq As Long

On Error GoTo Erro_PRJEtapas_Le_TiposCalculados

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189262
    Next
    
    objEtapa.lNumIntDoc = objEtapa.lNumIntDoc
    objEtapa.lNumIntDocPRJ = objEtapa.lNumIntDocPRJ
    
    lErro = CF("PRJCustoReceitaReal_Le_Etapa", objEtapa, colPRJCR)
    If lErro <> SUCESSO Then gError 189263
    
    iIndice = 0
    'Para cada documento associado ao projeto
    For Each objPRJCR In colPRJCR
    
        iIndice = iIndice + 1
        bAchou = False
        For Each objPRJInfo In colPRJInfo
            If objPRJInfo.lNumIntDoc = objPRJCR.lNumIntDocPRJInfo Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objPRJInfo = New ClassProjetoInfo
            
            objPRJInfo.lNumIntDoc = objPRJCR.lNumIntDocPRJInfo
            
            lErro = CF("ProjetoInfo_Le_NumIntDoc", objPRJInfo)
            If lErro <> SUCESSO Then gError 189264
        End If
        
        'Pega as OPs
        If objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_OP Then
        
            Set objOP = New ClassOrdemDeProducao
        
            objOP.sCodigo = objPRJInfo.sCodigoOP
            objOP.iFilialEmpresa = objPRJInfo.iFilialEmpresa
        
            colOP.Add objOP
        
        End If
        
        colPRJInfo.Add objPRJInfo
    Next
    
    For Each objOP In colOP
    
        sProduto = String(STRING_PRODUTO, 0)
    
        'Le a OP para ver se ela não foi gerada automaticamente e le seus itens
        lErro = Comando_Executar(alComando(0), "SELECT IOP.Item,IOP.Produto, IOP.NumIntDoc, IOP.Quantidade FROM OrdensProducaoTodas AS OP, ItensOPTodos IOP WHERE OP.OPGeradora= '' AND OP.Codigo = IOP.Codigo AND OP.FilialEmpresa = IOP.FilialEmpresa AND OP.Codigo =? AND OP.FilialEmpresa =?", _
                  iItem, sProduto, lNumIntDocItemOP, dQuantidade, objOP.sCodigo, objOP.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 189265
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189266
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objItemOP = New ClassItemOP
        
            objItemOP.iItem = iItem
            objItemOP.sProduto = sProduto
            objItemOP.lNumIntDoc = lNumIntDocItemOP
            objItemOP.sCodigo = objOP.sCodigo
            objItemOP.iFilialEmpresa = objOP.iFilialEmpresa
            objItemOP.dQuantidade = dQuantidade
            
            lErro = PRJEtapas_Le_TiposCalculados_Recursivo(objEtapaAux, objItemOP, colOP, colPRJCR, colPRJInfo, iSeqMaq, iSeqMO, iSeqMP)
            If lErro <> SUCESSO Then gError 189267
            
            objOP.colItens.Add objItemOP
        
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189268
        
        Loop
        
    Next
    
    sProduto = String(STRING_PRODUTO, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    'Le a tabelaApontProdPRJ
    lErro = Comando_Executar(alComando(1), "SELECT AP.Produto, AP.UM, AP.Qtd, AP.Custo, AP.Obs, A.Data, P.Descricao FROM ApontProdPRJ AS AP, Produtos AS P, ApontPRJ AS A WHERE AP.NumIntDocApont= A.NumIntDoc AND AP.Produto = P.Codigo AND A.NumIntDocPRJ = ? AND A.NumIntDocEtapa = ? ", _
                sProduto, sUM, dQtd, dCusto, sObs, dtData, sDescricao, objEtapa.lNumIntDocPRJ, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194515

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194516

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objEtapaMP = New ClassPRJEtapaMateriais
        
        objEtapaMP.dCusto = dCusto
        objEtapaMP.dQuantidade = dQtd
        objEtapaMP.dtData = dtData
        objEtapaMP.iTipo = INDICE_CALC_REAL
        objEtapaMP.sDescricao = sDescricao
        objEtapaMP.sObservacao = sObs
        objEtapaMP.sProduto = sProduto
        objEtapaMP.sUM = sUM
        
        objEtapaAux.colMateriaPrima.Add objEtapaMP
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194517
        
    Loop
    
    sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    sDescricao = String(STRING_MAQUINA_DESCRICAO, 0)
    
    'Le a tabelaApontProdPRJ
    lErro = Comando_Executar(alComando(2), "SELECT M.NumIntDoc, AM.CodMaq, AM.Horas, AM.Qtd, AM.Custo, AM.Obs, A.Data, M.Descricao FROM ApontMaqPRJ AS AM, Maquinas AS M, ApontPRJ AS A WHERE AM.NumIntDocApont= A.NumIntDoc AND AM.CodMaq = M.Codigo AND A.NumIntDocPRJ = ? AND A.NumIntDocEtapa = ?", _
                lNumIntDocMaq, iMaquina, dHoras, iQtd, dCusto, sObs, dtData, sDescricao, objEtapa.lNumIntDocPRJ, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194518

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194519

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objEtapaMaq = New ClassPRJEtapaMaquinas
        
        objEtapaMaq.dCusto = dCusto
        objEtapaMaq.iQuantidade = iQtd
        objEtapaMaq.dtData = dtData
        objEtapaMaq.iTipo = INDICE_CALC_REAL
        objEtapaMaq.sDescricao = sDescricao
        objEtapaMaq.sObservacao = sObs
        objEtapaMaq.iMaquina = iMaquina
        objEtapaMaq.dHoras = dHoras
        objEtapaMaq.lNumIntDocMaq = lNumIntDocMaq
        
        objEtapaAux.colMaquinas.Add objEtapaMaq
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194520
        
    Loop
    
    sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    sDescricao = String(STRING_MAQUINA_DESCRICAO, 0)
    
    'Le a tabelaApontProdPRJ
    lErro = Comando_Executar(alComando(3), "SELECT AM.CodMO, AM.Horas, AM.Qtd, AM.Custo, AM.Obs, A.Data, M.Descricao FROM ApontMOPRJ AS AM, TiposDeMaoDeObra AS M, ApontPRJ AS A WHERE AM.NumIntDocApont= A.NumIntDoc AND AM.CodMO = M.Codigo AND A.NumIntDocPRJ = ? AND A.NumIntDocEtapa = ?", _
                iMO, dHoras, iQtd, dCusto, sObs, dtData, sDescricao, objEtapa.lNumIntDocPRJ, objEtapa.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194521

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194522

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objEtapaMO = New ClassPRJEtapaMO
        
        objEtapaMO.dCusto = dCusto
        objEtapaMO.iQuantidade = iQtd
        objEtapaMO.dtData = dtData
        objEtapaMO.iTipo = INDICE_CALC_REAL
        objEtapaMO.sDescricao = sDescricao
        objEtapaMO.sObservacao = sObs
        objEtapaMO.iMaoDeObra = iMO
        objEtapaMO.dHoras = dHoras
        
        objEtapaAux.colMaoDeObra.Add objEtapaMO
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194523
        
    Loop
    
    For Each objEtapaMO In objEtapaAux.colMaoDeObra
        bAchou = False
'        For Each objEtapaMOAux In colEtapaMOAux
'            If objEtapaMO.dtData = objEtapaMOAux.dtData And objEtapaMO.iMaoDeObra = objEtapaMOAux.iMaoDeObra And Abs(objEtapaMOAux.dCusto / objEtapaMOAux.dHoras - objEtapaMO.dCusto / objEtapaMO.dHoras) < QTDE_ESTOQUE_DELTA Then
'                bAchou = True
'                objEtapaMOAux.dCusto = objEtapaMOAux.dCusto + objEtapaMO.dCusto
'                objEtapaMOAux.dHoras = objEtapaMOAux.dHoras + objEtapaMO.dHoras
'                If objEtapaMOAux.iQuantidade < objEtapaMO.iQuantidade Then
'                    objEtapaMOAux.iQuantidade = objEtapaMO.iQuantidade
'                End If
'                Exit For
'            End If
'        Next
        If Not bAchou Then
            colEtapaMOAux.Add objEtapaMO
        End If
    Next
    For Each objEtapaMaq In objEtapaAux.colMaquinas
        bAchou = False
'        For Each objEtapaMaqAux In colEtapaMaqAux
'            If objEtapaMaq.dtData = objEtapaMaqAux.dtData And objEtapaMaq.lNumIntDocMaq = objEtapaMaqAux.lNumIntDocMaq And Abs(objEtapaMaqAux.dCusto / objEtapaMaqAux.dHoras - objEtapaMaq.dCusto / objEtapaMaq.dHoras) < QTDE_ESTOQUE_DELTA Then
'                bAchou = True
'                objEtapaMaqAux.dCusto = objEtapaMaqAux.dCusto + objEtapaMaq.dCusto
'                objEtapaMaqAux.dHoras = objEtapaMaqAux.dHoras + objEtapaMaq.dHoras
'                If objEtapaMaqAux.iQuantidade < objEtapaMaq.iQuantidade Then
'                    objEtapaMaqAux.iQuantidade = objEtapaMaq.iQuantidade
'                End If
'                Exit For
'            End If
'        Next
        If Not bAchou Then
            colEtapaMaqAux.Add objEtapaMaq
        End If
    Next
    For Each objEtapaMP In objEtapaAux.colMateriaPrima
        bAchou = False
'        For Each objEtapaMPAux In colEtapaMPAux
'            If objEtapaMP.dtData = objEtapaMPAux.dtData And objEtapaMPAux.sProduto = objEtapaMP.sProduto And Abs(objEtapaMPAux.dCusto / objEtapaMPAux.dQuantidade - objEtapaMP.dCusto / objEtapaMP.dQuantidade) < QTDE_ESTOQUE_DELTA Then
'                bAchou = True
'                objEtapaMPAux.dCusto = objEtapaMPAux.dCusto + objEtapaMP.dCusto
'                objEtapaMPAux.dQuantidade = objEtapaMPAux.dQuantidade + objEtapaMP.dQuantidade
'                Exit For
'            End If
'        Next
        If Not bAchou Then
            
            colEtapaMPAux.Add objEtapaMP
        End If
    Next
    
    Set colSaida = New Collection
    Set colCampos = New Collection
    
    colCampos.Add "dtData"
    colCampos.Add "iMaoDeObra"
    
    Call Ordena_Colecao(colEtapaMOAux, colSaida, colCampos)
    
    For Each objEtapaMO In colSaida
        objEtapa.colMaoDeObra.Add objEtapaMO
    Next
    
    Set colSaida = New Collection
    Set colCampos = New Collection
    
    colCampos.Add "dtData"
    colCampos.Add "iMaquina"
    
    Call Ordena_Colecao(colEtapaMaqAux, colSaida, colCampos)
    
    For Each objEtapaMaq In colSaida
        objEtapa.colMaquinas.Add objEtapaMaq
    Next
    
    Set colSaida = New Collection
    Set colCampos = New Collection
    
    colCampos.Add "dtData"
    colCampos.Add "sProduto"
    
    Call Ordena_Colecao(colEtapaMPAux, colSaida, colCampos)
    
    For Each objEtapaMP In colSaida
        objEtapa.colMateriaPrima.Add objEtapaMP
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PRJEtapas_Le_TiposCalculados = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_TiposCalculados:

    PRJEtapas_Le_TiposCalculados = gErr

    Select Case gErr
    
        Case 189262
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 189263, 189264, 189267
        
        Case 189265, 189266, 189268
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)
            
        Case 194515 To 194517
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTPRODPRJ", gErr)
        
        Case 194518 To 194520
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTMAQPRJ", gErr)
        
        Case 194521 To 194523
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTMOPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189269)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function PRJEtapas_Le_TiposCalculados_Recursivo(ByVal objEtapa As ClassPRJEtapas, ByVal objItemOP As ClassItemOP, ByVal colOP As Collection, ByVal colPRJCR As Collection, ByVal colPRJInfo As Collection, iSeqMaq As Integer, iSeqMO As Integer, iSeqMP As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 6) As Long
Dim sSQL1 As String
Dim lNumIntDoc1 As Long, lNumIntDoc2 As Long
Dim lNumIntDoc1Ant As Long, lNumIntDocMaq As Long
Dim dHorasMaq As Double, dHorasMO As Double
Dim iQtdMaq As Integer, iQtdMO As Integer, dQuantidade As Double
Dim iMO As Integer, iMaq As Integer
Dim objEtapaMaq As ClassPRJEtapaMaquinas, objEtapaMO As ClassPRJEtapaMO
Dim dCustoMaq As Double, dCustoMO As Double
Dim sDescMaq As String, sDescMO As String
Dim objPRJInfo As ClassProjetoInfo
Dim objPRJCR As ClassPRJCR
Dim dPercentual As Double
Dim bAchou As Boolean
Dim dtData As Date

On Error GoTo Erro_PRJEtapas_Le_TiposCalculados_Recursivo

    bAchou = False
    iIndice = 0
    For Each objPRJCR In colPRJCR
        iIndice = iIndice + 1
        Set objPRJInfo = colPRJInfo.Item(iIndice)
        If objPRJInfo.sCodigoOP = objItemOP.sCodigo And objPRJInfo.iFilialEmpresa = objItemOP.iFilialEmpresa Then
            If objPRJCR.sItem = CStr(objItemOP.iItem) And objPRJCR.iUsado = DESMARCADO Then
                dQuantidade = dQuantidade + objPRJCR.dQuantidade - objPRJCR.dQtdUsado
                objPRJCR.iUsado = MARCADO
                bAchou = True
            End If
        End If
    Next
    
    'Se esse item da OP não faz parte dessa etapa -> Sai fora
    If Not bAchou Then Exit Function

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189270
    Next
    
    dPercentual = dQuantidade / objItemOP.dQuantidade
    
    sDescMaq = String(STRING_MAQUINA_DESCRICAO, 0)
    sDescMO = String(STRING_TIPO_MO_DESCRICAO, 0)

    'Le as máquinas e mão de Obra usados nesse item de OP
    sSQL1 = "SELECT MAQ.Data, MAQ.NumIntDoc, MAQ.NumIntDocMaq, MAQ.Horas, MAQ.Quantidade,M.CustoHora, M.Codigo,M.Descricao, MO.NumIntDoc, MO.TipoMO, MO.Horas, MO.Quantidade, O.CustoHora, O.Descricao FROM OrdemProducaoOperacoes AS OPER,ItemOPOperacoesMaquinas AS MAQ, ItemOPOperacoesMO AS MO, Maquinas AS M, TiposDeMaodeObra AS O " & _
        "WHERE M.NumIntDoc = MAQ.NumIntDocMaq AND O.Codigo = MO.TipoMO AND  OPER.NumIntDoc = MAQ.NumIntDocOper AND MAQ.NumIntDoc = MO.NumIntDocItemOPMaq AND OPER.NumIntDocItemOP = ? " & _
        "ORDER BY MAQ.NumIntDoc, MO.NumIntDoc"

    lErro = Comando_Executar(alComando(1), sSQL1, dtData, lNumIntDoc1, lNumIntDocMaq, dHorasMaq, iQtdMaq, dCustoMaq, iMaq, sDescMaq, lNumIntDoc2, iMO, dHorasMO, iQtdMO, dCustoMO, sDescMO, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189271

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189272
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If lNumIntDoc1 <> lNumIntDoc1Ant Then
        
            'Carrega as máquinas usadas (Real)
            iSeqMaq = iSeqMaq + 1
            Set objEtapaMaq = New ClassPRJEtapaMaquinas
            
            objEtapaMaq.dCusto = dHorasMaq * dCustoMaq * dPercentual
            objEtapaMaq.dHoras = dHorasMaq * dPercentual
            objEtapaMaq.iMaquina = iMaq
            objEtapaMaq.iQuantidade = iQtdMaq
            objEtapaMaq.iSeq = iSeqMaq
            objEtapaMaq.iTipo = INDICE_CALC_REAL
            objEtapaMaq.lNumIntDocEtapa = objEtapa.lNumIntDoc
            objEtapaMaq.lNumIntDocMaq = lNumIntDocMaq
            objEtapaMaq.sDescricao = sDescMaq
            objEtapaMaq.dtData = dtData
            objEtapaMaq.sObservacao = "Apontamento OP " & objItemOP.sCodigo
            
            objEtapa.colMaquinas.Add objEtapaMaq
        
            lNumIntDoc1Ant = lNumIntDoc1
        End If

        'Carrega a mão de Obra usada (real)
        iSeqMO = iSeqMO + 1
        Set objEtapaMO = New ClassPRJEtapaMO
        
        objEtapaMO.dCusto = dHorasMO * dCustoMO * dPercentual
        objEtapaMO.dHoras = dHorasMO * dPercentual
        objEtapaMO.iMaoDeObra = iMO
        objEtapaMO.iQuantidade = iQtdMO
        objEtapaMO.iSeq = iSeqMO
        objEtapaMO.iTipo = INDICE_CALC_REAL
        objEtapaMO.lNumIntDocEtapa = objEtapa.lNumIntDoc
        objEtapaMO.sDescricao = sDescMO
        objEtapaMO.dtData = dtData
        objEtapaMO.sObservacao = "Apontamento OP " & objItemOP.sCodigo
        
        objEtapa.colMaoDeObra.Add objEtapaMO
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189273
    
    Loop
    
    lErro = PRJEtapas_Le_TiposCalculados_Recursivo2(objEtapa, objItemOP, colOP, colPRJCR, colPRJInfo, iSeqMaq, iSeqMO, iSeqMP, alComando)
    If lErro <> SUCESSO Then gError 189274
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PRJEtapas_Le_TiposCalculados_Recursivo = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_TiposCalculados_Recursivo:

    PRJEtapas_Le_TiposCalculados_Recursivo = gErr

    Select Case gErr
    
        Case 189270
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 189271 To 189273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)
        
        Case 189274

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189275)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function PRJEtapas_Le_TiposCalculados_Recursivo2(ByVal objEtapa As ClassPRJEtapas, ByVal objItemOP As ClassItemOP, ByVal colOP As Collection, ByVal colPRJCR As Collection, ByVal colPRJInfo As Collection, iSeqMaq As Integer, iSeqMO As Integer, iSeqMP As Integer, alComando() As Long) As Long

Dim lErro As Long
Dim iIndice As Integer, iIndice2 As Integer
Dim sSQL2 As String, sSQL3 As String, sSQL4 As String, sSQL5 As String
Dim lNumIntDoc3 As Long, lNumIntDoc4 As Long
Dim lNumIntDocItemOP As Long
Dim dQtdMP As Double, dQtdProdEnt As Double, dQtdReq As Double, dQuantidade As Double, dQtdItemOP As Double, dQtdNF As Double
Dim objEtapaMP As ClassPRJEtapaMateriais
Dim dCustoMP As Double
Dim sDescMP As String
Dim objPRJInfo As ClassProjetoInfo
Dim objPRJCR As ClassPRJCR
Dim dPercentual As Double
Dim sProduto As String
Dim bAchou As Boolean
Dim sUM1 As String, sUM2 As String
Dim iTipoMov As Integer
Dim iProdutoComprado As Integer
Dim dFatorProdEntrada As Double, dFatorNF As Double
Dim objItemOPFilha As ClassItemOP
Dim iItem1 As Integer, iItem2 As Integer, iItem3 As Integer
Dim sCodigoOP As String
Dim iFilialEmpresa As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colMovEstoque As Collection
Dim colNF As Collection
Dim objNF As ClassNFiscal
Dim colItemNF As Collection
Dim objItemNF As ClassItemNF
Dim bFinalizaLoop As Boolean
Dim dFator As Double
Dim iClasseUM As Integer
Dim lCodigoMovEst As Long
Dim lCodigoProdEnt As Long
Dim dtData As Date

On Error GoTo Erro_PRJEtapas_Le_TiposCalculados_Recursivo2
    
    sProduto = String(STRING_PRODUTO, 0)
    sDescMP = String(STRING_PRODUTO_DESCRICAO, 0)
    sUM1 = String(STRING_UM_SIGLA, 0)
    
    'Le as requisições para esse item da OP
    sSQL2 = "SELECT M.Data, M.Codigo, M.NumIntDoc,M.Item, M.Custo,M.SiglaUM, M.Quantidade, M.Produto, P.Descricao, M.TipoMov, P.Compras, P.ClasseUM FROM MovimentoEstoque AS M, Produtos AS P WHERE M.Produto = P.Codigo AND M.NumIntDocEst=0 AND M.TipoMov IN (?,?) AND M.TipoNumIntDocOrigem = 5 AND M.NumIntDocOrigem = ?"
    
    lErro = Comando_Executar(alComando(2), sSQL2, dtData, lCodigoMovEst, lNumIntDoc3, iItem1, dCustoMP, sUM1, dQtdMP, sProduto, sDescMP, iTipoMov, iProdutoComprado, iClasseUM, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 189281

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189282
    
    'Para cada requisição
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Busca a requisição para ver se ela foi marcada para fazer parte dessa etapa
        bAchou = False
        iIndice = 0
        dQuantidade = 0
        For Each objPRJCR In colPRJCR
            iIndice = iIndice + 1
            Set objPRJInfo = colPRJInfo.Item(iIndice)
            If objPRJInfo.lNumIntDocOrigem = lCodigoMovEst And objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_REQPROD Then
                If objPRJCR.sItem = CStr(iItem1) Then
                    dQuantidade = dQuantidade + objPRJCR.dQuantidade
                    bAchou = True
                End If
            End If
        Next
    
        'Se faz parte da etapa
        If bAchou Then

            dQtdReq = dQuantidade
            
            'Calcula o percentual alocada dessa requisição para esse projeto
            dPercentual = dQuantidade / dQtdMP
            
            'Se o produto for produzível tem que ver se não existiu uma produção entrada para ele
            dFatorProdEntrada = 1
            dFatorNF = 1
            If iProdutoComprado <> PRODUTO_COMPRAVEL Then
            
                sCodigoOP = String(STRING_OPCODIGO, 0)
                sUM2 = String(STRING_UM_SIGLA, 0)
            
                'Tenta Ler as entradas de produção desse produto
                sSQL3 = "SELECT ME.Codigo, ME.NumIntDoc, ME.NumIntDocOrigem, ME.Quantidade, ME.SiglaUM, ME.Item,IOP.Item,IOP.Quantidade, IOP.Codigo, IOP.FilialEmpresa FROM MovimentoEstoque AS ME, ItensOPTodos IOP  WHERE IOP.NumIntDoc = ME.NumIntDocOrigem AND ME.TipoMov IN (?,?) AND ME.NumIntDocEst=0 AND ME.TipoNumIntDocOrigem=5  AND ME.Produto = ?"
                
                lErro = Comando_Executar(alComando(3), sSQL3, lCodigoProdEnt, lNumIntDoc4, lNumIntDocItemOP, dQtdProdEnt, sUM2, iItem2, iItem3, dQtdItemOP, sCodigoOP, iFilialEmpresa, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 189283
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189284
                
                dQuantidade = 0
                bAchou = False
                bFinalizaLoop = False
                Do While lErro <> AD_SQL_SEM_DADOS
                
                    'Descobre o fator de conversao da UM2 d p/UM1
                    lErro = CF("UM_Conversao_Trans", iClasseUM, sUM2, sUM1, dFator)
                    If lErro <> SUCESSO Then gError 189290

                    'Busca para ver se essa produção entrada faz parte desse projeto\etapa
                    iIndice = 0
                    For Each objPRJCR In colPRJCR
                        iIndice = iIndice + 1
                        Set objPRJInfo = colPRJInfo.Item(iIndice)
                        If objPRJInfo.lNumIntDocOrigem = lCodigoProdEnt And objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_PRODENTRADA Then
                            'Se é esse item que faz parte da etapa e ainda falta quantidade
                            'para a produção chegar a quantidade da requisição
                            If objPRJCR.sItem = CStr(iItem2) And objPRJCR.iUsado <> MARCADO Then
                                dQuantidade = dQuantidade + (objPRJCR.dQuantidade - objPRJCR.dQtdUsado) * dFator
                                bAchou = True
                                'Se o requisitado é maior que o produzido
                                If dQtdReq > dQuantidade Then
                                    objPRJCR.dQtdUsado = objPRJCR.dQuantidade
                                    objPRJCR.iUsado = MARCADO
                                Else
                                    objPRJCR.dQtdUsado = objPRJCR.dQuantidade - (dQuantidade - (dQtdReq / dFator))
                                    dQuantidade = dQtdReq
                                    
                                    'Já achou todas as produções necessária
                                    bFinalizaLoop = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                                        
                    'Se existe um produção entrada
                    If bAchou Then
                    
                        Set objItemOPFilha = New ClassItemOP
                        
                        objItemOPFilha.sProduto = sProduto
                        objItemOPFilha.lNumIntDoc = lNumIntDocItemOP
                        
                        objItemOPFilha.iItem = iItem3
                        objItemOPFilha.sCodigo = sCodigoOP
                        objItemOPFilha.iFilialEmpresa = iFilialEmpresa
                        objItemOPFilha.dQuantidade = dQtdItemOP
                    
                        lErro = PRJEtapas_Le_TiposCalculados_Recursivo(objEtapa, objItemOPFilha, colOP, colPRJCR, colPRJInfo, iSeqMaq, iSeqMO, iSeqMP)
                        If lErro <> SUCESSO Then gError 189286
                    
                    End If
                    
                    If bFinalizaLoop Then Exit Do

                    'Busca Próximo
                    lErro = Comando_BuscarProximo(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189285

                Loop
                
                dQtdProdEnt = dQuantidade
            
                'Abate do produto intermediário o que foi produzido
                dFatorProdEntrada = 1 - dQtdProdEnt / dQtdReq
                
            Else
            'Se comprado dar preferência pelo custo das notas de entrada ligadas ao projeto
            'Se não, pegar pelo custo médio do movimento de estoque
            
                Set colNF = New Collection
                Set colItemNF = New Collection
                Set colMovEstoque = New Collection

                lErro = PRJEtapas_Le_TiposCalculados_Recursivo3(objEtapa, colPRJCR, colPRJInfo, sProduto, colMovEstoque, colNF, colItemNF)
                If lErro <> SUCESSO Then gError 189287
                
                iIndice2 = 0
                dQtdNF = 0
                dQuantidade = 0
                For Each objItemMovEstoque In colMovEstoque
                    iIndice2 = iIndice2 + 1
                    Set objNF = colNF.Item(iIndice2)
                    Set objItemNF = colItemNF.Item(iIndice2)
                    
                    'Descobre o fator de conversao da UM2 d p/UM1
                    lErro = CF("UM_Conversao_Trans", iClasseUM, objItemMovEstoque.sSiglaUM, sUM1, dFator)
                    If lErro <> SUCESSO Then gError 189291

                    iIndice = 0
                    For Each objPRJCR In colPRJCR
                        iIndice = iIndice + 1
                        Set objPRJInfo = colPRJInfo.Item(iIndice)
                        If objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_NF And objPRJInfo.lNumIntDocOrigem = objNF.lNumIntDoc Then
                            If objPRJCR.iUsado = DESMARCADO And objPRJCR.sItem = CStr(objItemNF.iItem) Then
                                
                                dFatorNF = (objPRJCR.dQuantidade - objPRJCR.dQtdUsado) / objItemNF.dQuantidade
                                
                                dQtdNF = dQtdNF + objItemMovEstoque.dQuantidade * dFatorNF * dFator
                                
                                iSeqMP = iSeqMP + 1
                                Set objEtapaMP = New ClassPRJEtapaMateriais
                                
                                objEtapaMP.iSeq = iSeqMP
                                objEtapaMP.iTipo = INDICE_CALC_REAL
                                objEtapaMP.lNumIntDocEtapa = objEtapa.lNumIntDoc
                                objEtapaMP.sDescricao = objItemNF.sDescricaoItem
                                objEtapaMP.sProduto = objItemNF.sProduto
                                objEtapaMP.sUM = objItemMovEstoque.sSiglaUM
                                objEtapaMP.dtData = objItemMovEstoque.dtData
                                objEtapaMP.sObservacao = "Requisição para produção código " & CStr(lCodigoMovEst)
                                
                                objEtapa.colMateriaPrima.Add objEtapaMP
                                
                                If dQtdReq > dQtdNF Then
                                    objEtapaMP.dCusto = objItemMovEstoque.dCusto * dFatorNF
                                    objEtapaMP.dQuantidade = objItemMovEstoque.dQuantidade * dFatorNF
                                    
                                    objPRJCR.dQtdUsado = objPRJCR.dQuantidade
                                    objPRJCR.iUsado = MARCADO
                                Else
                                    objEtapaMP.dCusto = objItemMovEstoque.dCusto * (objPRJCR.dQuantidade - dQtdNF - dQtdReq - objPRJCR.dQtdUsado) * dFatorNF
                                    objEtapaMP.dQuantidade = objItemMovEstoque.dQuantidade * (objPRJCR.dQuantidade - dQtdNF - dQtdReq - objPRJCR.dQtdUsado) * dFatorNF
                                    
                                    objPRJCR.dQtdUsado = objPRJCR.dQuantidade - (dQtdNF - dQtdReq)
                                    dQtdNF = dQtdReq
                                    
                                    'Já achou todas as nf necessária
                                    bFinalizaLoop = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                                        
                    If bFinalizaLoop Then Exit For
                Next
                
                dFatorNF = 1 - dQtdNF / dQtdReq

            End If
            
            'Se não teve entrada de produção completa para esse item,
            'Adiciona como Insumo considerando o percentual que é dessa etapa
            'e o percentual que veio de uma produção
            If dFatorProdEntrada > 0 And dFatorNF > 0 Then
            
                iSeqMP = iSeqMP + 1
                Set objEtapaMP = New ClassPRJEtapaMateriais
                
                objEtapaMP.dCusto = dCustoMP * dPercentual * dFatorProdEntrada * dFatorNF
                objEtapaMP.dQuantidade = dQtdMP * dPercentual * dFatorProdEntrada * dFatorNF
                objEtapaMP.iSeq = iSeqMP
                objEtapaMP.iTipo = INDICE_CALC_REAL
                objEtapaMP.lNumIntDocEtapa = objEtapa.lNumIntDoc
                objEtapaMP.sDescricao = sDescMP
                objEtapaMP.sProduto = sProduto
                objEtapaMP.sUM = sUM1
                objEtapaMP.dtData = dtData
                objEtapaMP.sObservacao = "Requisição para produção código " & CStr(lCodigoMovEst)
                                
                objEtapa.colMateriaPrima.Add objEtapaMP
                
            End If
        
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189288
    
    Loop

    PRJEtapas_Le_TiposCalculados_Recursivo2 = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_TiposCalculados_Recursivo2:

    PRJEtapas_Le_TiposCalculados_Recursivo2 = gErr

    Select Case gErr
    
        Case 189281 To 189285, 189288
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)
            
        Case 189286, 189287, 189290, 189291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189289)

    End Select

    Exit Function

End Function

Private Function PRJEtapas_Le_TiposCalculados_Recursivo3(ByVal objEtapa As ClassPRJEtapas, ByVal colPRJCR As Collection, ByVal colPRJInfo As Collection, ByVal sProduto As String, ByVal colMovEstoque As Collection, ByVal colNF As Collection, ByVal colItemNF As Collection) As Long
'Le as produções entradas relacionadas a um produto para a etapa

Dim lErro As Long
Dim iIndice As Integer
Dim objPRJInfo As ClassProjetoInfo
Dim objPRJCR As ClassPRJCR
Dim objNF As ClassNFiscal
Dim objItemNF As ClassItemNF
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colMovEstoqueAux As Collection
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_PRJEtapas_Le_TiposCalculados_Recursivo3
    
    'Busca as notas de entrada que ainda não foram usadas
    iIndice = 0
    For Each objPRJCR In colPRJCR
        iIndice = iIndice + 1
        Set objPRJInfo = colPRJInfo.Item(iIndice)
        If objPRJInfo.iTipoOrigem = PRJ_CR_TIPO_NF Then
            If objPRJCR.iUsado = DESMARCADO Then
            
                Set objNF = New ClassNFiscal
                Set objTipoDocInfo = New ClassTipoDocInfo
                
                objNF.lNumIntDoc = objPRJInfo.lNumIntDocOrigem
                
                lErro = CF("NFiscal_Le", objNF)
                If lErro <> SUCESSO And lErro <> 31442 Then gError 189276
                
                objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
            
                'lê o Tipo da Nota Fiscal
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 189277
                
                'Se participa da estatística de compras
                If objTipoDocInfo.iCompras = MARCADO Then
                
                    lErro = CF("NFiscalItens_Le", objNF)
                    If lErro <> SUCESSO Then gError 189278
                    
                    For Each objItemNF In objNF.ColItensNF
                    
                        If objPRJCR.sItem = CStr(objItemNF.iItem) And objItemNF.sProduto = sProduto Then
                
                            Set colMovEstoqueAux = New Collection
                            Set objItemMovEstoque = New ClassItemMovEstoque
                            
                            objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
                            objItemMovEstoque.iFilialEmpresa = objNF.iFilialEmpresa
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        
                            lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoqueAux)
                            If lErro <> SUCESSO Then gError 189279

                            For Each objItemMovEstoque In colMovEstoqueAux
                                colMovEstoque.Add objItemMovEstoque
                                colNF.Add objNF
                                colItemNF.Add objItemNF
                            Next
                            
                        End If
                        
                    Next

                End If

            End If
            
        End If
        
    Next

    PRJEtapas_Le_TiposCalculados_Recursivo3 = SUCESSO

    Exit Function

Erro_PRJEtapas_Le_TiposCalculados_Recursivo3:

    PRJEtapas_Le_TiposCalculados_Recursivo3 = gErr

    Select Case gErr
    
        Case 189276 To 189279

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189280)

    End Select

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJContratos_Le(ByVal objPRJContratos As ClassPRJContratos, Optional ByVal bLeItens As Boolean = True, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJContratos As typePRJContratos

On Error GoTo Erro_PRJContratos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185857

    'Alocação de espaço no buffer
    tPRJContratos.sCodigo = String(STRING_CONTRATOPRJ_CODIGO, 0)
    tPRJContratos.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPRJContratos.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

    'Le a tabelaPRJContratos
    lErro = Comando_Executar(lComando, "SELECT NaturezaOP, NumIntDoc, NumIntDocPRJ, NumIntDocProposta, Codigo, Data, Observacao, Cliente, " & _
                "FilialCliente, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, " & _
                "CustoInformado, CustoCalculado, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco FROM PRJContratos WHERE NumIntDocPRJ= ?  AND Codigo= ? ", _
                tPRJContratos.sNaturezaOp, tPRJContratos.lNumIntDoc, tPRJContratos.lNumIntDocPRJ, tPRJContratos.lNumIntDocProposta, tPRJContratos.sCodigo, tPRJContratos.dtData, _
                tPRJContratos.sObservacao, tPRJContratos.lCliente, tPRJContratos.iFilialCliente, tPRJContratos.dValorTotal, tPRJContratos.dValorProdutos, _
                tPRJContratos.dValorFrete, tPRJContratos.dValorDesconto, tPRJContratos.dValorSeguro, tPRJContratos.dValorOutrasDespesas, tPRJContratos.dCustoInformado, _
                tPRJContratos.dCustoCalculado, tPRJContratos.iExibirProdutos, tPRJContratos.iExibirCustoCalc, tPRJContratos.iExibirCustoInfo, tPRJContratos.iExibirPreco, _
                objPRJContratos.lNumIntDocPRJ, objPRJContratos.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185858

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185859

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJContratos.lNumIntDoc = tPRJContratos.lNumIntDoc
    objPRJContratos.lNumIntDocPRJ = tPRJContratos.lNumIntDocPRJ
    objPRJContratos.lNumIntDocProposta = tPRJContratos.lNumIntDocProposta
    objPRJContratos.sCodigo = tPRJContratos.sCodigo
    objPRJContratos.dtData = tPRJContratos.dtData
    objPRJContratos.sObservacao = tPRJContratos.sObservacao
    objPRJContratos.lCliente = tPRJContratos.lCliente
    objPRJContratos.iFilialCliente = tPRJContratos.iFilialCliente
    objPRJContratos.dValorTotal = tPRJContratos.dValorTotal
    objPRJContratos.dValorProdutos = tPRJContratos.dValorProdutos
    objPRJContratos.dValorFrete = tPRJContratos.dValorFrete
    objPRJContratos.dValorDesconto = tPRJContratos.dValorDesconto
    objPRJContratos.dValorSeguro = tPRJContratos.dValorSeguro
    objPRJContratos.dValorOutrasDespesas = tPRJContratos.dValorOutrasDespesas
    objPRJContratos.dCustoInformado = tPRJContratos.dCustoInformado
    objPRJContratos.dCustoCalculado = tPRJContratos.dCustoCalculado
    objPRJContratos.iExibirCustoCalc = tPRJContratos.iExibirCustoCalc
    objPRJContratos.iExibirCustoInfo = tPRJContratos.iExibirCustoInfo
    objPRJContratos.iExibirPreco = tPRJContratos.iExibirPreco
    objPRJContratos.iExibirProdutos = tPRJContratos.iExibirProdutos
    objPRJContratos.sNaturezaOp = tPRJContratos.sNaturezaOp
    
    If bLeItens Then
    
        lErro = CF("PRJContratoItens_Le", objPRJContratos, bLeTrib)
        If lErro <> SUCESSO Then gError 187515
    
        lErro = CF("PRJContratoEtapa_Le", objPRJContratos)
        If lErro <> SUCESSO Then gError 187525
        
    End If
    
    If bLeTrib Then
    
        lErro = CF("PRJContratos_Le_Tributacao", objPRJContratos)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187599
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJContratos_Le = SUCESSO

    Exit Function

Erro_PRJContratos_Le:

    PRJContratos_Le = gErr

    Select Case gErr

        Case 185857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 185858, 185859
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOS", gErr)
            
        Case 187515, 187525, 187599

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185860)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJContratos_Le_NumIntDoc(ByVal objPRJContratos As ClassPRJContratos, Optional ByVal bLeItens As Boolean = True, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJContratos As typePRJContratos

On Error GoTo Erro_PRJContratos_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187997

    'Alocação de espaço no buffer
    tPRJContratos.sCodigo = String(STRING_CONTRATOPRJ_CODIGO, 0)
    tPRJContratos.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPRJContratos.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

    'Le a tabelaPRJContratos
    lErro = Comando_Executar(lComando, "SELECT NaturezaOP, NumIntDoc, NumIntDocPRJ,NumIntDocProposta, Codigo, Data, Observacao, Cliente, " & _
                "FilialCliente, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, " & _
                "CustoInformado, CustoCalculado, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco FROM PRJContratos WHERE NumIntDoc= ? ", _
                tPRJContratos.sNaturezaOp, tPRJContratos.lNumIntDoc, tPRJContratos.lNumIntDocPRJ, tPRJContratos.lNumIntDocProposta, tPRJContratos.sCodigo, tPRJContratos.dtData, _
                tPRJContratos.sObservacao, tPRJContratos.lCliente, tPRJContratos.iFilialCliente, tPRJContratos.dValorTotal, tPRJContratos.dValorProdutos, _
                tPRJContratos.dValorFrete, tPRJContratos.dValorDesconto, tPRJContratos.dValorSeguro, tPRJContratos.dValorOutrasDespesas, tPRJContratos.dCustoInformado, _
                tPRJContratos.dCustoCalculado, tPRJContratos.iExibirProdutos, tPRJContratos.iExibirCustoCalc, tPRJContratos.iExibirCustoInfo, tPRJContratos.iExibirPreco, _
                objPRJContratos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187998

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187999

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJContratos.lNumIntDoc = tPRJContratos.lNumIntDoc
    objPRJContratos.lNumIntDocPRJ = tPRJContratos.lNumIntDocPRJ
    objPRJContratos.lNumIntDocProposta = tPRJContratos.lNumIntDocProposta
    objPRJContratos.sCodigo = tPRJContratos.sCodigo
    objPRJContratos.dtData = tPRJContratos.dtData
    objPRJContratos.sObservacao = tPRJContratos.sObservacao
    objPRJContratos.lCliente = tPRJContratos.lCliente
    objPRJContratos.iFilialCliente = tPRJContratos.iFilialCliente
    objPRJContratos.dValorTotal = tPRJContratos.dValorTotal
    objPRJContratos.dValorProdutos = tPRJContratos.dValorProdutos
    objPRJContratos.dValorFrete = tPRJContratos.dValorFrete
    objPRJContratos.dValorDesconto = tPRJContratos.dValorDesconto
    objPRJContratos.dValorSeguro = tPRJContratos.dValorSeguro
    objPRJContratos.dValorOutrasDespesas = tPRJContratos.dValorOutrasDespesas
    objPRJContratos.dCustoInformado = tPRJContratos.dCustoInformado
    objPRJContratos.dCustoCalculado = tPRJContratos.dCustoCalculado
    objPRJContratos.iExibirCustoCalc = tPRJContratos.iExibirCustoCalc
    objPRJContratos.iExibirCustoInfo = tPRJContratos.iExibirCustoInfo
    objPRJContratos.iExibirPreco = tPRJContratos.iExibirPreco
    objPRJContratos.iExibirProdutos = tPRJContratos.iExibirProdutos
    objPRJContratos.sNaturezaOp = tPRJContratos.sNaturezaOp
    
    If bLeItens Then
    
        lErro = CF("PRJContratoItens_Le", objPRJContratos, bLeTrib)
        If lErro <> SUCESSO Then gError 189000
    
        lErro = CF("PRJContratoEtapa_Le", objPRJContratos)
        If lErro <> SUCESSO Then gError 189001
        
    End If
    
    If bLeTrib Then
    
        lErro = CF("PRJContratos_Le_NumIntDoc_Tributacao", objPRJContratos)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189002
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJContratos_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_PRJContratos_Le_NumIntDoc:

    PRJContratos_Le_NumIntDoc = gErr

    Select Case gErr

        Case 187997
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187998, 187999
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOS", gErr)
            
        Case 189000, 189001, 189002

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189003)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJRecebPagto_Le_Contrato(ByVal objPRJRecebPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJRecebPagto As typePRJRecebPagto

On Error GoTo Erro_PRJRecebPagto_Le_Contrato

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187488

    'Le a tabelaPRJRecebPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPRJ, Tipo, Numero, Valor, CliForn, " & _
                "Filial, NumIntDocContrato, IncluiCFF, FilialEmpresa FROM PRJRecebPagto WHERE NumIntDocPRJ= ?  AND NumIntDocContrato= ? AND Tipo = ?", _
                tPRJRecebPagto.lNumIntDoc, tPRJRecebPagto.lNumIntDocPRJ, tPRJRecebPagto.iTipo, tPRJRecebPagto.lNumero, _
                tPRJRecebPagto.dValor, tPRJRecebPagto.lCliForn, tPRJRecebPagto.iFilial, tPRJRecebPagto.lNumIntDocContrato, tPRJRecebPagto.iIncluiCFF, _
                tPRJRecebPagto.iFilialEmpresa, _
                objPRJRecebPagto.lNumIntDocPRJ, objPRJRecebPagto.lNumIntDocContrato, objPRJRecebPagto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 187489

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187490

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJRecebPagto.lNumIntDoc = tPRJRecebPagto.lNumIntDoc
    objPRJRecebPagto.lNumIntDocPRJ = tPRJRecebPagto.lNumIntDocPRJ
    objPRJRecebPagto.iTipo = tPRJRecebPagto.iTipo
    objPRJRecebPagto.lNumero = tPRJRecebPagto.lNumero
    objPRJRecebPagto.dValor = tPRJRecebPagto.dValor
    objPRJRecebPagto.lCliForn = tPRJRecebPagto.lCliForn
    objPRJRecebPagto.iFilial = tPRJRecebPagto.iFilial
    objPRJRecebPagto.lNumIntDocContrato = tPRJRecebPagto.lNumIntDocContrato
    objPRJRecebPagto.iIncluiCFF = tPRJRecebPagto.iIncluiCFF
    objPRJRecebPagto.iFilialEmpresa = tPRJRecebPagto.iFilialEmpresa

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJRecebPagto_Le_Contrato = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Le_Contrato:

    PRJRecebPagto_Le_Contrato = gErr

    Select Case gErr

        Case 187488
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187489, 187490
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJRECEBPAGTO", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187491)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJContratoItens_Le(ByVal objPRJContrato As ClassPRJContratos, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJContratoItem As typePRJContratoItem
Dim objPRJContratoItem As ClassPRJContratoItem

On Error GoTo Erro_PRJContratoItens_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187510

    'Alocação de espaço no buffer
    tPRJContratoItem.sProduto = String(STRING_PRODUTO, 0)
    tPRJContratoItem.sDescProd = String(STRING_DESCRICAO_ITEM, 0)
    tPRJContratoItem.sDescEtapa = String(STRING_DESCRICAO_ITEM, 0)
    tPRJContratoItem.sUM = String(STRING_UM_SIGLA, 0)
    tPRJContratoItem.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaPRJContratoItens
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Item, NumIntDocContrato, Produto, DescProd, " & _
                "NumIntDocEtapa, DescEtapa, Quantidade, PrecoUnitario, PrecoTotal, UnidadeMed, ValorDesconto, " & _
                "DataEntrega, Observacao FROM PRJContratoItens WHERE NumIntDocContrato= ? ", _
                tPRJContratoItem.lNumIntDoc, tPRJContratoItem.iFilialEmpresa, tPRJContratoItem.iItem, tPRJContratoItem.lNumIntDocContrato, _
                tPRJContratoItem.sProduto, tPRJContratoItem.sDescProd, tPRJContratoItem.lNumIntDocEtapa, tPRJContratoItem.sDescEtapa, tPRJContratoItem.dQuantidade, _
                tPRJContratoItem.dPrecoUnitario, tPRJContratoItem.dPrecoTotal, tPRJContratoItem.sUM, tPRJContratoItem.dValorDesconto, tPRJContratoItem.dtDataEntrega, _
                tPRJContratoItem.sObservacao, _
                objPRJContrato.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187511

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187512

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPRJContratoItem = New ClassPRJContratoItem

        objPRJContratoItem.lNumIntDoc = tPRJContratoItem.lNumIntDoc
        objPRJContratoItem.iFilialEmpresa = tPRJContratoItem.iFilialEmpresa
        objPRJContratoItem.iItem = tPRJContratoItem.iItem
        objPRJContratoItem.lNumIntDocContrato = tPRJContratoItem.lNumIntDocContrato
        objPRJContratoItem.sProduto = tPRJContratoItem.sProduto
        objPRJContratoItem.sDescProd = tPRJContratoItem.sDescProd
        objPRJContratoItem.lNumIntDocEtapa = tPRJContratoItem.lNumIntDocEtapa
        objPRJContratoItem.sDescEtapa = tPRJContratoItem.sDescEtapa
        objPRJContratoItem.dQuantidade = tPRJContratoItem.dQuantidade
        objPRJContratoItem.dPrecoUnitario = tPRJContratoItem.dPrecoUnitario
        objPRJContratoItem.dPrecoTotal = tPRJContratoItem.dPrecoTotal
        objPRJContratoItem.sUM = tPRJContratoItem.sUM
        objPRJContratoItem.dValorDesconto = tPRJContratoItem.dValorDesconto
        objPRJContratoItem.dtDataEntrega = tPRJContratoItem.dtDataEntrega
        objPRJContratoItem.sObservacao = tPRJContratoItem.sObservacao
        
        objPRJContrato.colItens.Add objPRJContratoItem
        
        If bLeTrib Then
        
            lErro = PRJContratoItem_Le_Tributacao(objPRJContratoItem)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187605
        
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187513
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJContratoItens_Le = SUCESSO

    Exit Function

Erro_PRJContratoItens_Le:

    PRJContratoItens_Le = gErr

    Select Case gErr

        Case 187510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187511, 187512, 187513
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOITENS", gErr)
            
        Case 187605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187514)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJContratoEtapa_Le(ByVal objPRJContrato As ClassPRJContratos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJContratoEtapa As typePRJContratoEtapa
Dim objPRJContratoEtapa As ClassPRJContratoEtapa

On Error GoTo Erro_PRJContratoEtapa_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 187516

    tPRJContratoEtapa.sDescricao = String(STRING_CONTRATOETAPA_DESCRICAO, 0)
    tPRJContratoEtapa.sObservacao = String(STRING_CONTRATOETAPA_OBSERVACAO, 0)

    'Le a tabelaPRJContratoEtapa
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocContrato, NumIntDocEtapa, NumIntDocEtapaItemProd, CustoInformado, Preco, Selecionado,Imprimir,Descricao,Observacao FROM PRJContratoEtapa WHERE NumIntDocContrato= ? ", _
                tPRJContratoEtapa.lNumIntDoc, tPRJContratoEtapa.lNumIntDocContrato, tPRJContratoEtapa.lNumIntDocEtapa, tPRJContratoEtapa.lNumIntDocEtapaItemProd, _
                tPRJContratoEtapa.dCustoInformado, tPRJContratoEtapa.dPreco, tPRJContratoEtapa.iSelecionado, tPRJContratoEtapa.iImprimir, tPRJContratoEtapa.sDescricao, tPRJContratoEtapa.sObservacao, _
                objPRJContrato.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 187517

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187518

    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objPRJContratoEtapa = New ClassPRJContratoEtapa

        objPRJContratoEtapa.lNumIntDoc = tPRJContratoEtapa.lNumIntDoc
        objPRJContratoEtapa.lNumIntDocContrato = tPRJContratoEtapa.lNumIntDocContrato
        objPRJContratoEtapa.lNumIntDocEtapa = tPRJContratoEtapa.lNumIntDocEtapa
        objPRJContratoEtapa.lNumIntDocEtapaItemProd = tPRJContratoEtapa.lNumIntDocEtapaItemProd
        objPRJContratoEtapa.iSelecionado = tPRJContratoEtapa.iSelecionado
        objPRJContratoEtapa.dCustoInformado = tPRJContratoEtapa.dCustoInformado
        objPRJContratoEtapa.dPreco = tPRJContratoEtapa.dPreco
        objPRJContratoEtapa.iImprimir = tPRJContratoEtapa.iImprimir
        objPRJContratoEtapa.sDescricao = tPRJContratoEtapa.sDescricao
        objPRJContratoEtapa.sObservacao = tPRJContratoEtapa.sObservacao
        
        objPRJContrato.colEtapas.Add objPRJContratoEtapa

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187519

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJContratoEtapa_Le = SUCESSO

    Exit Function

Erro_PRJContratoEtapa_Le:

    PRJContratoEtapa_Le = gErr

    Select Case gErr

        Case 187516
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 187517, 187518, 187519
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOETAPA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187520)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PRJContratos_Le_Tributacao(ByVal objContrato As ClassPRJContratos) As Long
'Le a Tributacao da Contrato de Venda Passado como parametro.

Dim lErro As Long
'Dim lComando As Long
'Dim tTributacaoPRJCTR As typeTributacaoPRJCTR
'Dim objTributacaoPRJCTR As New ClassTributacaoPRJCTR

On Error GoTo Erro_PRJContratos_Le_Tributacao

'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 187585
'
'    With tTributacaoPRJCTR
'
'        .sCodContrato = String(STRING_CONTRATOPRJ_CODIGO, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, CodContrato, TaxacaoAutomatica, TipoTributacao, TipoTributacaoManual, IPIBase, IPIBaseManual, IPIValor, IPIValorManual, ICMSBase, ICMSBaseManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, ICMSSubstValor, ICMSSubstValorManual, ISSIncluso, ISSBase, ISSAliquota, ISSAliquotaManual, ISSValor, ISSValorManual, IRRFBase, IRRFAliquota, IRRFAliquotaManual, IRRFValor, IRRFValorManual, PISRetidoManual, ISSRetidoManual, COFINSRetidoManual, CSLLRetidoManual, PISRetido, ISSRetido, COFINSRetido, CSLLRetido, " & _
'        "PISCredito,PISCreditoManual,COFINSCredito,COFINSCreditoManual,IPICredito,IPICreditoManual,ICMSCredito,ICMSCreditoManual,ValorINSS,ValorINSSManual,INSSRetido,INSSRetidoManual,INSSValorBase,INSSValorBaseManual,INSSValorDeducoes ,INSSValorDeducoesManual FROM TributacaoPRJCTR WHERE FilialEmpresa = ? AND CodContrato = ?", _
'            .iFilialEmpresa, .sCodContrato, .iTaxacaoAutomatica, .iTipoTributacao, .iTipoTributacaoManual, .dIPIBase, .iIPIBaseManual, .dIPIValor, .iIPIValorManual, .dICMSBase, .iICMSBaseManual, .dICMSValor, .iICMSValorManual, .dICMSSubstBase, .iICMSSubstBaseManual, .dICMSSubstValor, .iICMSSubstValorManual, .iISSIncluso, .dISSBase, .dISSAliquota, .iISSAliquotaManual, .dISSValor, .iISSValorManual, .dIRRFBase, .dIRRFAliquota, .iIRRFAliquotaManual, .dIRRFValor, .iIRRFValorManual, .iPISRetidoManual, .iISSRetidoManual, .iCOFINSRetidoManual, .iCSLLRetidoManual, .dPISRetido, .dISSRetido, .dCOFINSRetido, .dCSLLRetido, _
'            .dPISCredito, .iPISCreditoManual, .dCOFINSCredito, .iCOFINSCreditoManual, .dIPICredito, .iIPICreditoManual, .dICMSCredito, .iICMSCreditoManual, .dValorINSS, .iINSSValorManual, .iINSSRetido, .iINSSRetidoManual, .dINSSBase, .iINSSBaseManual, .dINSSDeducoes, .iINSSDeducoesManual, _
'            objContrato.iFilialEmpresa, objContrato.sCodigo)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 187586
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187587
'
'    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    With objTributacaoPRJCTR
'        .dICMSBase = tTributacaoPRJCTR.dICMSBase
'        .dICMSSubstBase = tTributacaoPRJCTR.dICMSSubstBase
'        .dICMSSubstValor = tTributacaoPRJCTR.dICMSSubstValor
'        .dICMSValor = tTributacaoPRJCTR.dICMSValor
'        .dIPIBase = tTributacaoPRJCTR.dIPIBase
'        .dIPIValor = tTributacaoPRJCTR.dIPIValor
'        .dIRRFAliquota = tTributacaoPRJCTR.dIRRFAliquota
'        .dIRRFBase = tTributacaoPRJCTR.dIRRFBase
'        .dIRRFValor = tTributacaoPRJCTR.dIRRFValor
'        .dISSAliquota = tTributacaoPRJCTR.dISSAliquota
'        .dISSBase = tTributacaoPRJCTR.dISSBase
'        .dISSValor = tTributacaoPRJCTR.dISSValor
'        .iFilialEmpresa = tTributacaoPRJCTR.iFilialEmpresa
'        .iICMSBaseManual = tTributacaoPRJCTR.iICMSBaseManual
'        .iICMSSubstBaseManual = tTributacaoPRJCTR.iICMSSubstBaseManual
'        .iICMSSubstValorManual = tTributacaoPRJCTR.iICMSSubstValorManual
'        .iICMSValorManual = tTributacaoPRJCTR.iICMSValorManual
'        .iIPIBaseManual = tTributacaoPRJCTR.iIPIBaseManual
'        .iIPIValorManual = tTributacaoPRJCTR.iIPIValorManual
'        .iIRRFAliquotaManual = tTributacaoPRJCTR.iIRRFAliquotaManual
'        .iIRRFValorManual = tTributacaoPRJCTR.iIRRFValorManual
'        .iISSIncluso = tTributacaoPRJCTR.iISSIncluso
'        .iISSInclusoManual = tTributacaoPRJCTR.iISSValorManual
'        .iISSValorManual = tTributacaoPRJCTR.iISSValorManual
'        .iTaxacaoAutomatica = tTributacaoPRJCTR.iTaxacaoAutomatica
'        .iTipoTributacao = tTributacaoPRJCTR.iTipoTributacao
'        .iTipoTributacaoManual = tTributacaoPRJCTR.iTipoTributacaoManual
'        .sCodContrato = tTributacaoPRJCTR.sCodContrato
'        .iPISRetidoManual = tTributacaoPRJCTR.iPISRetidoManual
'        .iISSRetidoManual = tTributacaoPRJCTR.iISSRetidoManual
'        .iCOFINSRetidoManual = tTributacaoPRJCTR.iCOFINSRetidoManual
'        .iCSLLRetidoManual = tTributacaoPRJCTR.iCSLLRetidoManual
'        .dPISRetido = tTributacaoPRJCTR.dPISRetido
'        .dISSRetido = tTributacaoPRJCTR.dISSRetido
'        .dCOFINSRetido = tTributacaoPRJCTR.dCOFINSRetido
'        .dCSLLRetido = tTributacaoPRJCTR.dCSLLRetido
'
'        .dPISCredito = tTributacaoPRJCTR.dPISCredito
'        .iPISCreditoManual = tTributacaoPRJCTR.iPISCreditoManual
'        .dCOFINSCredito = tTributacaoPRJCTR.dCOFINSCredito
'        .iCOFINSCreditoManual = tTributacaoPRJCTR.iCOFINSCreditoManual
'        .dIPICredito = tTributacaoPRJCTR.dIPICredito
'        .iIPICreditoManual = tTributacaoPRJCTR.iIPICreditoManual
'        .dICMSCredito = tTributacaoPRJCTR.dICMSCredito
'        .iICMSCreditoManual = tTributacaoPRJCTR.iICMSCreditoManual
'        .dValorINSS = tTributacaoPRJCTR.dValorINSS
'        .iINSSValorManual = tTributacaoPRJCTR.iINSSValorManual
'        .iINSSRetido = tTributacaoPRJCTR.iINSSRetido
'        .iINSSRetidoManual = tTributacaoPRJCTR.iINSSRetidoManual
'        .dINSSBase = tTributacaoPRJCTR.dINSSBase
'        .iINSSBaseManual = tTributacaoPRJCTR.iINSSBaseManual
'        .dINSSDeducoes = tTributacaoPRJCTR.dINSSDeducoes
'        .iINSSDeducoesManual = tTributacaoPRJCTR.iINSSDeducoesManual
'    End With
'
'    Set objContrato.objTributacaoPRJCTR = objTributacaoPRJCTR
'
'    lErro = CF("PRJContratos_Le_Complementos", objContrato)
'    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187589
'
'    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187590
'
'    Call Comando_Fechar(lComando)
'
    objContrato.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_CTR
    objContrato.objTributacao.lNumIntDoc = objContrato.lNumIntDoc

    lErro = CF("TributacaoDoc_Le", objContrato.objTributacao)
    If lErro <> SUCESSO Then gError 187589
    
    PRJContratos_Le_Tributacao = SUCESSO

    Exit Function

Erro_PRJContratos_Le_Tributacao:

    PRJContratos_Le_Tributacao = gErr

    Select Case gErr

'        Case 187585
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 187586, 187587
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBCONTRATO", gErr)
'
'        Case ERRO_LEITURA_SEM_DADOS
        
        Case 187589
'
'        Case 187590
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_CONTRATO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187591)

    End Select

'    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Public Function PRJContratos_Le_Complementos(ByVal objContrato As ClassPRJContratos) As Long
'
'Dim lComando As Long, objTributacaoPRJCTR As ClassTributacaoPRJCTR
'Dim tTributacaoComplPRJCTR As typeTributacaoComplPV, objTributacaoItemPRJCTR As ClassTributacaoItemPV
'Dim lErro As Long
'
'On Error GoTo Erro_PRJContratos_Le_Complementos
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 187592
'
'    Set objTributacaoPRJCTR = objContrato.objTributacaoPRJCTR
'
'    tTributacaoComplPRJCTR.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'
'    'ler dados de frete, seguro, despesas, descontos da tabela TributacaoComplPRJCTR
'    lErro = Comando_Executar(lComando, "SELECT Item, NaturezaOp, NaturezaOpManual, TipoTributacao, TipoTributacaoManual, IPITipo, IPITipoManual, IPIBaseCalculo, IPIBaseManual, IPIPercRedBase, IPIPercRedBaseManual, IPIAliquota, " & _
'        "IPIAliquotaManual, IPIValor, IPIValorManual, ICMSTipo, ICMSTipoManual, ICMSBase, ICMSBaseManual, ICMSPercRedBase, ICMSPercRedBaseManual, ICMSAliquota, ICMSAliquotaManual, ICMSValor, ICMSValorManual, ICMSSubstBase, ICMSSubstBaseManual, " & _
'        "ICMSSubstAliquota, ICMSSubstAliquotaManual, ICMSSubstValor, ICMSSubstValorManual, PISCRedito, COFINSCredito, IPICredito, ICMSCredito FROM TributacaoComplPRJCTR WHERE FilialEmpresa = ? AND CodContrato = ?", _
'        tTributacaoComplPRJCTR.iItem, tTributacaoComplPRJCTR.sNaturezaOp, tTributacaoComplPRJCTR.iNaturezaOpManual, tTributacaoComplPRJCTR.iTipoTributacao, tTributacaoComplPRJCTR.iTipoTributacaoManual, tTributacaoComplPRJCTR.iIPITipo, tTributacaoComplPRJCTR.iIPITipoManual, tTributacaoComplPRJCTR.dIPIBaseCalculo, tTributacaoComplPRJCTR.iIPIBaseManual, tTributacaoComplPRJCTR.dIPIPercRedBase, _
'        tTributacaoComplPRJCTR.iIPIPercRedBaseManual, tTributacaoComplPRJCTR.dIPIAliquota, tTributacaoComplPRJCTR.iIPIAliquotaManual, tTributacaoComplPRJCTR.dIPIValor, tTributacaoComplPRJCTR.iIPIValorManual, tTributacaoComplPRJCTR.iICMSTipo, tTributacaoComplPRJCTR.iICMSTipoManual, tTributacaoComplPRJCTR.dICMSBase, tTributacaoComplPRJCTR.iICMSBaseManual, tTributacaoComplPRJCTR.dICMSPercRedBase, _
'        tTributacaoComplPRJCTR.iICMSPercRedBaseManual, tTributacaoComplPRJCTR.dICMSAliquota, tTributacaoComplPRJCTR.iICMSAliquotaManual, tTributacaoComplPRJCTR.dICMSValor, tTributacaoComplPRJCTR.iICMSValorManual, tTributacaoComplPRJCTR.dICMSSubstBase, tTributacaoComplPRJCTR.iICMSSubstBaseManual, _
'        tTributacaoComplPRJCTR.dICMSSubstAliquota, tTributacaoComplPRJCTR.iICMSSubstAliquotaManual, tTributacaoComplPRJCTR.dICMSSubstValor, tTributacaoComplPRJCTR.iICMSSubstValorManual, tTributacaoComplPRJCTR.dPISCredito, tTributacaoComplPRJCTR.dCOFINSCredito, tTributacaoComplPRJCTR.dIPICredito, tTributacaoComplPRJCTR.dICMSCredito, objContrato.iFilialEmpresa, objContrato.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 187593
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTributacaoItemPRJCTR = New ClassTributacaoItemPV
'
'        With objTributacaoItemPRJCTR
'
'            .lNumIntDoc = tTributacaoComplPRJCTR.iItem
'            .sNaturezaOp = tTributacaoComplPRJCTR.sNaturezaOp
'            .iNaturezaOpManual = tTributacaoComplPRJCTR.iNaturezaOpManual
'            .iTipoTributacao = tTributacaoComplPRJCTR.iTipoTributacao
'            .iTipoTributacaoManual = tTributacaoComplPRJCTR.iTipoTributacaoManual
'            .iIPITipo = tTributacaoComplPRJCTR.iIPITipo
'            .iIPITipoManual = tTributacaoComplPRJCTR.iIPITipoManual
'            .dIPIBaseCalculo = tTributacaoComplPRJCTR.dIPIBaseCalculo
'            .iIPIBaseManual = tTributacaoComplPRJCTR.iIPIBaseManual
'            .dIPIPercRedBase = tTributacaoComplPRJCTR.dIPIPercRedBase
'            .iIPIPercRedBaseManual = tTributacaoComplPRJCTR.iIPIPercRedBaseManual
'            .dIPIAliquota = tTributacaoComplPRJCTR.dIPIAliquota
'            .iIPIAliquotaManual = tTributacaoComplPRJCTR.iIPIAliquotaManual
'            .dIPIValor = tTributacaoComplPRJCTR.dIPIValor
'            .iIPIValorManual = tTributacaoComplPRJCTR.iIPIValorManual
'            .iICMSTipo = tTributacaoComplPRJCTR.iICMSTipo
'            .iICMSTipoManual = tTributacaoComplPRJCTR.iICMSTipoManual
'            .dICMSBase = tTributacaoComplPRJCTR.dICMSBase
'            .iICMSBaseManual = tTributacaoComplPRJCTR.iICMSBaseManual
'            .dICMSPercRedBase = tTributacaoComplPRJCTR.dICMSPercRedBase
'            .iICMSPercRedBaseManual = tTributacaoComplPRJCTR.iICMSPercRedBaseManual
'            .dICMSAliquota = tTributacaoComplPRJCTR.dICMSAliquota
'            .iICMSAliquotaManual = tTributacaoComplPRJCTR.iICMSAliquotaManual
'            .dICMSValor = tTributacaoComplPRJCTR.dICMSValor
'            .iICMSValorManual = tTributacaoComplPRJCTR.iICMSValorManual
'            .dICMSSubstBase = tTributacaoComplPRJCTR.dICMSSubstBase
'            .iICMSSubstBaseManual = tTributacaoComplPRJCTR.iICMSSubstBaseManual
'            .dICMSSubstAliquota = tTributacaoComplPRJCTR.dICMSSubstAliquota
'            .iICMSSubstAliquotaManual = tTributacaoComplPRJCTR.iICMSSubstAliquotaManual
'            .dICMSSubstValor = tTributacaoComplPRJCTR.dICMSSubstValor
'            .iICMSSubstValorManual = tTributacaoComplPRJCTR.iICMSSubstValorManual
'            .dICMSCredito = tTributacaoComplPRJCTR.dICMSCredito
'            .dIPICredito = tTributacaoComplPRJCTR.dIPICredito
'            .dCOFINSCredito = tTributacaoComplPRJCTR.dCOFINSCredito
'            .dPISCredito = tTributacaoComplPRJCTR.dPISCredito
'        End With
'
'        Select Case tTributacaoComplPRJCTR.iItem
'
'            Case ITEM_TIPO_FRETE
'                Set objTributacaoPRJCTR.objTributacaoFrete = objTributacaoItemPRJCTR
'
'            Case ITEM_TIPO_SEGURO
'                Set objTributacaoPRJCTR.objTributacaoSeguro = objTributacaoItemPRJCTR
'
'            Case ITEM_TIPO_DESCONTO
'                Set objTributacaoPRJCTR.objTributacaoDesconto = objTributacaoItemPRJCTR
'
'            Case ITEM_TIPO_OUTRAS_DESP
'                Set objTributacaoPRJCTR.objTributacaoOutras = objTributacaoItemPRJCTR
'
'            Case Else 'item invalido
'                gError 187596
'
'        End Select
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187597
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    PRJContratos_Le_Complementos = SUCESSO
'
'    Exit Function
'
'Erro_PRJContratos_Le_Complementos:
'
'    PRJContratos_Le_Complementos = gErr
'
'    Select Case gErr
'
'        Case 187592
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 187593, ERRO_LEITURA_SEM_DADOS, 187597
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJContratos", gErr)
'
'        Case 187596
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMPL_CONTRATO_TIPO_INVALIDO", gErr)
'
'        Case ERRO_LEITURA_SEM_DADOS
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187598)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Public Function PRJContratoItem_Le_Tributacao(ByVal objContratoItem As ClassPRJContratoItem) As Long

Dim lErro As Long

On Error GoTo Erro_PRJContratoItem_Le_Tributacao

    objContratoItem.objTributacao.iTipoDoc = TIPODOC_TRIB_PRJ_CTR
    objContratoItem.objTributacao.lNumIntDocItem = objContratoItem.lNumIntDoc
    
    lErro = CF("TributacaoDocItem_Le", objContratoItem.objTributacao)
    If lErro <> SUCESSO Then gError 187601
    
    PRJContratoItem_Le_Tributacao = SUCESSO

    Exit Function

Erro_PRJContratoItem_Le_Tributacao:

    PRJContratoItem_Le_Tributacao = gErr

    Select Case gErr

        Case 187601
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187604)

    End Select


    Exit Function

End Function

Public Function RelPRJCustoReceita_Le(ByVal objFlx As ClassFluxoPRJ) As Long

Dim lErro As Long
Dim tPRJCR As typePRJCR
Dim objPRJCR As ClassPRJCR
Dim colPRJCR As New Collection
Dim colPRJInfo As New Collection
Dim objProjetoInfo As ClassProjetoInfo
Dim sSQL As String, sSQLAux As String
Dim iIndiceTipo As Integer
Dim lComando As Long

On Error GoTo Erro_RelPRJCustoReceita_Le

    lErro = CF("PRJEtapas_Le_Projeto", objFlx.objProjeto, True, True, objFlx.alComando(15), objFlx.alComando(16), objFlx.alComando(17), objFlx.alComando(18), objFlx.alComando(19))
    If lErro <> SUCESSO Then gError 189690
    
    'Alterado para limitar as leituras de forma a só pegar documentos que tem algum tipo de movimentação
    'no período (quer seja prevista ou real ou planejada)
    'Estava com um desempenho horrível na Macroplan lendo tudo e calculando para só depois descobrir se tinha algo no período filtrado
      
    'Todos documentos vinculados ao projeto em questão que tenham valor
    sSQL = "SELECT F.NumIntDoc, F.NumIntDocPRJ, F.NumIntDocPRJInfo, F.NumIntDocEtapa, F.Valor, "
    sSQL = sSQL & " F.Quantidade, F.Percentual, "
    sSQL = sSQL & " F.TipoValor , F.Observacao, F.Item, F.CalcAuto "
    sSQL = sSQL & " FROM PRJCustoReceitaReal AS F, ProjetoInfo AS I "
    sSQL = sSQL & " WHERE F.NumIntDocPRJInfo = I.NumIntDoc AND F.NumIntDocPRJ = ? AND (F.NumIntDocEtapa = ? OR ?=0) AND F.Valor > ? "

    For iIndiceTipo = 1 To 6
    
        tPRJCR.sObservacao = String(STRING_PRJ_CR_OBSERVACAO, 0)
        tPRJCR.sItem = String(STRING_PRJ_CR_ITEM, 0)
       
        If objFlx.dtDataDe <> DATA_NULA Or objFlx.dtDataAte <> DATA_NULA Then
   
            Select Case iIndiceTipo
            
                Case 1
                
                    lComando = objFlx.alComando(3)
                                       
                    'Se for um documento ligado a NF
                    sSQLAux = " AND I.TipoOrigem = 1 AND ( "
                
                    'Se a NF estiver vinculado a um a pagar com data de vencimento ou data de baixa no período filtrado-> considera
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM NFiscal AS N WHERE N.ClasseDocCPR = 1 AND N.NumIntDoc = I.NumIntDocOrigem AND "
                    sSQLAux = sSQLAux & " ((EXISTS (SELECT * FROM ParcelasPagTodas AS X WHERE X.NumIntTitulo = N.NumIntDocCPR " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimentoReal >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimentoReal <= ? ", "") & ") OR "
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM BaixasPag AS X, BaixasParcPag AS Y, ParcelasPagTodas AS Z WHERE X.NumIntBaixa = Y.NumIntBaixa AND Z.NumIntDoc = Y.NumIntParcela AND Z.NumIntTitulo = N.NumIntDocCPR AND X.Status <> 5 " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.Data >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.Data <= ? ", "") & "))))) OR "
                
                    'Se a NF estiver vinculado a um a receber com data de vencimento ou data de baixa no período filtrado -> considera
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM NFiscal AS N WHERE N.ClasseDocCPR = 2 AND N.NumIntDoc = I.NumIntDocOrigem AND "
                    sSQLAux = sSQLAux & " ((EXISTS (SELECT * FROM ParcelasRecTodas AS X WHERE X.NumIntTitulo = N.NumIntDocCPR " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimento >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimento <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM BaixasRec AS X, BaixasParcRec AS Y, ParcelasRecTodas AS Z WHERE X.NumIntBaixa = Y.NumIntBaixa AND Z.NumIntDoc = Y.NumIntParcela AND Z.NumIntTitulo = N.NumIntDocCPR AND X.Status <> 5 " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.Data >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.Data <= ? ", "") & "))))) OR "
                
                    'Se a NF estiver vinculado a um nfspag com data de vencimento dele ou do título vinculado ou da baixa do título no período filtrado -> considera
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM NFiscal AS N WHERE N.ClasseDocCPR = 5 AND N.NumIntDoc = I.NumIntDocOrigem AND "
                    sSQLAux = sSQLAux & " ((EXISTS (SELECT * FROM NFsPag_Todas AS X WHERE X.NumIntDoc = N.NumIntDocCPR " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimento >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimento <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM ParcelasPagTodas AS X, NFsPag_Todas AS Y WHERE X.NumIntTitulo = Y.NumIntTitPag AND Y.NumIntDoc = N.NumIntDocCPR " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimentoReal >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimentoReal <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & " (EXISTS (SELECT * FROM BaixasPag AS X, BaixasParcPag AS Y, ParcelasPagTodas AS Z, NFsPag_Todas AS W WHERE X.NumIntBaixa = Y.NumIntBaixa AND Z.NumIntDoc = Y.NumIntParcela AND Z.NumIntTitulo = W.NumIntTitPag AND W.NumIntDoc = N.NumIntDocCPR AND X.Status <> 5 " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.Data >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.Data <= ? ", "") & "))))))) "
    
    
                    'Le a tabela PRJCustoReceitaReal
                    If objFlx.dtDataDe <> DATA_NULA And objFlx.dtDataAte <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte)
                    ElseIf objFlx.dtDataDe <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataDe, objFlx.dtDataDe, objFlx.dtDataDe, objFlx.dtDataDe, objFlx.dtDataDe, objFlx.dtDataDe)
                    Else
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataAte, objFlx.dtDataAte, objFlx.dtDataAte, objFlx.dtDataAte, objFlx.dtDataAte, objFlx.dtDataAte, objFlx.dtDataAte)
                    End If
    
                Case 2
                
                    lComando = objFlx.alComando(4)
                    
                    'Se tiver um movimento de conta corrente de depósito ou saque vinculado ao projeto no período filtrado -> considera
                    sSQLAux = " AND I.TipoOrigem IN (2,3) AND EXISTS (SELECT * FROM MovimentosContaCorrente AS X WHERE X.NumMovto = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataMovimento >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataMovimento <= ? ", "") & ") "
                
                    'Le a tabela PRJCustoReceitaReal
                    If objFlx.dtDataDe <> DATA_NULA And objFlx.dtDataAte <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataAte)
                    ElseIf objFlx.dtDataDe <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe)
                    Else
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataAte)
                    End If
    
                Case 3
                
                    lComando = objFlx.alComando(5)
                
                    'Se estiver vinculado a um a receber com data de vencimento ou data de baixa no período filtrado -> considera
                    sSQLAux = " AND  I.TipoOrigem = 4 AND ((EXISTS (SELECT * FROM ParcelasRecTodas AS X WHERE X.NumIntTitulo = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimentoReal >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimentoReal <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & "             (EXISTS (SELECT * FROM BaixasRec AS X, BaixasParcRec AS Y, ParcelasRecTodas AS Z WHERE X.NumIntBaixa = Y.NumIntBaixa AND Z.NumIntDoc = Y.NumIntParcela AND Z.NumIntTitulo = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.Data >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.Data <= ? ", "") & "))) "
                
                    'Le a tabela PRJCustoReceitaReal
                    If objFlx.dtDataDe <> DATA_NULA And objFlx.dtDataAte <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte)
                    ElseIf objFlx.dtDataDe <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataDe)
                    Else
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataAte, objFlx.dtDataAte)
                    End If
                Case 4
                    
                    lComando = objFlx.alComando(6)
                    
                    'Se estiver vinculado a um a pagar com data de vencimento ou data de baixa no período filtrado -> considera
                    sSQLAux = " AND  I.TipoOrigem = 5 AND ((EXISTS (SELECT * FROM ParcelasPagTodas AS X WHERE X.NumIntTitulo = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimentoReal >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimentoReal <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & "             (EXISTS (SELECT * FROM BaixasPag AS X, BaixasParcPag AS Y, ParcelasPagTodas AS Z WHERE X.NumIntBaixa = Y.NumIntBaixa AND Z.NumIntDoc = Y.NumIntParcela AND Z.NumIntTitulo = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.Data >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.Data <= ? ", "") & "))) "
                
                    'Le a tabela PRJCustoReceitaReal
                    If objFlx.dtDataDe <> DATA_NULA And objFlx.dtDataAte <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte)
                    ElseIf objFlx.dtDataDe <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataDe)
                    Else
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataAte, objFlx.dtDataAte)
                    End If
                
                Case 5
                
                    lComando = objFlx.alComando(7)
    
                    'Se estiver vinculado a um pv não faturado com data de vencimento no período filtrado -> considera
                    sSQLAux = " AND  I.TipoOrigem = 7 AND EXISTS (SELECT * FROM ParcelasPedidoDeVenda AS X, PedidosDeVenda AS Y WHERE Y.Codigo = X.PedidodeVenda AND Y.FilialEmpresa = X.FilialEmpresa AND Y.NumIntDoc = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimento >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimento <= ? ", "") & ") "
                
                    'Le a tabela PRJCustoReceitaReal
                    If objFlx.dtDataDe <> DATA_NULA And objFlx.dtDataAte <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataAte)
                    ElseIf objFlx.dtDataDe <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe)
                    Else
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataAte)
                    End If
                
                Case 6
                
                    lComando = objFlx.alComando(8)
    
                    'Se estiver vinculado a um nfspag com data de vencimento dele ou do título vinculado ou da baixa do título no período filtrado -> considera
                    sSQLAux = " AND  I.TipoOrigem = 8 AND ((EXISTS (SELECT * FROM NFsPag_Todas AS X WHERE X.NumIntDoc = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimento >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimento <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & "             (EXISTS (SELECT * FROM ParcelasPagTodas AS X, NFsPag_Todas AS Y WHERE X.NumIntTitulo = Y.NumIntTitPag AND Y.NumIntDoc = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.DataVencimentoReal >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.DataVencimentoReal <= ? ", "") & ")) OR "
                    sSQLAux = sSQLAux & "             (EXISTS (SELECT * FROM BaixasPag AS X, BaixasParcPag AS Y, ParcelasPagTodas AS Z, NFsPag_Todas AS W WHERE X.NumIntBaixa = Y.NumIntBaixa AND Z.NumIntDoc = Y.NumIntParcela AND Z.NumIntTitulo = W.NumIntTitPag AND W.NumIntDoc = I.NumIntDocOrigem " & IIf(objFlx.dtDataDe <> DATA_NULA, " AND X.Data >= ? ", "") & IIf(objFlx.dtDataAte <> DATA_NULA, " AND X.Data <= ? ", "") & "))) "
    
                
                    'Le a tabela PRJCustoReceitaReal
                    If objFlx.dtDataDe <> DATA_NULA And objFlx.dtDataAte <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte, objFlx.dtDataDe, objFlx.dtDataAte)
                    ElseIf objFlx.dtDataDe <> DATA_NULA Then
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataDe, objFlx.dtDataDe, objFlx.dtDataDe)
                    Else
                        lErro = Comando_Executar(lComando, sSQL & sSQLAux & " ORDER BY I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                                    objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0, objFlx.dtDataAte, objFlx.dtDataAte, objFlx.dtDataAte)
                    End If
            End Select
        Else
            If iIndiceTipo > 1 Then Exit For
            
            lComando = objFlx.alComando(3)
            
            'Le a tabela PRJCustoReceitaReal
            lErro = Comando_Executar(lComando, sSQL & " ORDER BY I.TipoOrigem, I.NumIntDoc", tPRJCR.lNumIntDoc, tPRJCR.lNumIntDocPRJ, tPRJCR.lNumIntDocPRJInfo, tPRJCR.lNumIntDocEtapa, tPRJCR.dValor, tPRJCR.dQuantidade, tPRJCR.dPercentual, tPRJCR.iTipoValor, tPRJCR.sObservacao, tPRJCR.sItem, tPRJCR.iCalcAuto, _
                        objFlx.objProjeto.lNumIntDoc, objFlx.lNumIntEtapa, objFlx.lNumIntEtapa, 0)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 189646
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189647
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objPRJCR = New ClassPRJCR
            Set objProjetoInfo = New ClassProjetoInfo
    
            objPRJCR.lNumIntDoc = tPRJCR.lNumIntDoc
            objPRJCR.lNumIntDocPRJ = tPRJCR.lNumIntDocPRJ
            objPRJCR.lNumIntDocEtapa = tPRJCR.lNumIntDocEtapa
            objPRJCR.dValor = tPRJCR.dValor
            objPRJCR.dQuantidade = tPRJCR.dQuantidade
            objPRJCR.dPercent = tPRJCR.dPercentual
            objPRJCR.iTipoValor = tPRJCR.iTipoValor
            objPRJCR.sObservacao = tPRJCR.sObservacao
            objPRJCR.sItem = tPRJCR.sItem
            objPRJCR.iCalcAuto = tPRJCR.iCalcAuto
            objPRJCR.lNumIntDocPRJInfo = tPRJCR.lNumIntDocPRJInfo
            
            objProjetoInfo.lNumIntDoc = tPRJCR.lNumIntDocPRJInfo
            
            lErro = ProjetoInfo_Le_NumIntDoc(objProjetoInfo, False, objFlx.alComando(9))
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189648
            
            objPRJCR.iTipoOrigem = objProjetoInfo.iTipoOrigem
            objPRJCR.lNumIntDocOrigem = objProjetoInfo.lNumIntDocOrigem
            
            colPRJCR.Add objPRJCR
            colPRJInfo.Add objProjetoInfo
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189649
            
        Loop
        
    Next
    
    lErro = RelPRJCustoReceita_Trata_TipoDocOrigem(objFlx, colPRJCR, colPRJInfo)
    If lErro <> SUCESSO Then gError 189650
    
    lErro = RelPRJCustoReceita_Trata_Producao(objFlx)
    If lErro <> SUCESSO Then gError 189682
    
    lErro = RelPRJCustoReceita_Trata_RecebPagto(objFlx)
    If lErro <> SUCESSO Then gError 189689

    RelPRJCustoReceita_Le = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Le:

    RelPRJCustoReceita_Le = gErr

    Select Case gErr
        
        Case 189646, 189647, 189649
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCUSTORECEITAREAL", gErr)
            
        Case 189648, 189650, 189682, 189689, 189690
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189651)

    End Select
    
    Resume Next

    Exit Function

End Function

Public Function RelPRJCustoReceita_Trata_TipoDocOrigem(ByVal objFlx As ClassFluxoPRJ, ByVal colPRJCR As Collection, ByVal colPRJInfo As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lIndiceAux As Long
Dim objPRJCR As ClassPRJCR
Dim objPRJCRAux As ClassPRJCR
Dim objProjetoInfo As ClassProjetoInfo
Dim objProjetoInfoAux As ClassProjetoInfo
Dim objOrigem As Object
Dim sDocOrigem As String
Dim sNomeTela As String
Dim objFluxo As ClassPRJFluxoFinanc
Dim objNFiscal As ClassNFiscal
Dim objTituloRec As ClassTituloReceber
Dim objTituloPag As ClassTituloPagar
Dim objNFPag As ClassNFsPag
Dim objOrcamentoVenda As ClassOrcamentoVenda
Dim objPedidoVenda As ClassPedidoDeVenda
Dim objMovConta As ClassMovContaCorrente
Dim objOP As ClassOrdemDeProducao
Dim objMovEstoque As ClassMovEstoque
Dim objEtapa As ClassPRJEtapas
Dim sDocOrigemRed As String, bAchou As Boolean, iTipoCPR As Integer
Dim lNumIntPRJInfoAnt As Long
Dim objNatureza As ClassNatMovCta, colNat As New Collection

On Error GoTo Erro_RelPRJCustoReceita_Trata_TipoDocOrigem

    lNumIntPRJInfoAnt = -1
  
    iIndice = 0
    For Each objProjetoInfo In colPRJInfo
    
        Set objFluxo = New ClassPRJFluxoFinanc
        
        If lNumIntPRJInfoAnt <> objProjetoInfo.lNumIntDoc Then
        
            lNumIntPRJInfoAnt = objProjetoInfo.lNumIntDoc
        
            lErro = CF("Projeto_Le_DocInfo", objProjetoInfo.lNumIntDoc, sDocOrigem, objOrigem, sNomeTela, sDocOrigemRed)
            If lErro <> SUCESSO Then gError 189653
            
        End If
    
        iIndice = iIndice + 1
        Set objPRJCR = colPRJCR.Item(iIndice)
        
        objFluxo.sDocOrigem = sDocOrigemRed
        objFluxo.sNomeTela = sNomeTela
        
        Set objEtapa = New ClassPRJEtapas
        Set objFluxo.objOrigem = objOrigem
        Set objFluxo.objPRJCR = objPRJCR
        Set objFluxo.objPRJInfo = objProjetoInfo
        Set objFluxo.objProjeto = objFlx.objProjeto
        
        objFluxo.sObs = objPRJCR.sObservacao
        
        If objPRJCR.lNumIntDocEtapa <> 0 Then
        
            objEtapa.lNumIntDoc = objPRJCR.lNumIntDocEtapa
            
            lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa, objFlx.alComando(10))
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189653
        
        End If
        
        Set objFluxo.objEtapa = objEtapa
        
        Select Case objProjetoInfo.iTipoOrigem
    
            Case PRJ_CR_TIPO_NF
                'Pega os valores baixados como real e os a baixar como previsto
                'Depende do tipo de nota (entrada ou saída) para considerar se
                'é receita ou despesa
                Set objNFiscal = objOrigem
                
                objFluxo.sNatureza = objNFiscal.sNatureza
                
                If objNFiscal.iStatus <> STATUS_CANCELADO Then
                
                    'Verifica se associaram diretamente o projeto ao financeiro ligado a NF
                    'Se sim ignora a associação na NF, vai prevalecer a do financeiro
                    iTipoCPR = -1
                    Select Case objNFiscal.iClasseDocCPR
                        Case CPR_TITULO_PAGAR
                            iTipoCPR = PRJ_CR_TIPO_TITPAG
                        Case CPR_TITULO_RECEBER
                            iTipoCPR = PRJ_CR_TIPO_TITREC
                        Case CPR_NF_PAGAR
                            iTipoCPR = PRJ_CR_TIPO_NFPAG
                    End Select
                
                    bAchou = False
                    lIndiceAux = 0
                    
                    For Each objProjetoInfoAux In colPRJInfo
                        lIndiceAux = lIndiceAux + 1
                        Set objPRJCRAux = colPRJCR.Item(lIndiceAux)
                        If objProjetoInfoAux.iTipoOrigem = iTipoCPR And objPRJCRAux.lNumIntDocOrigem = objNFiscal.lNumIntDocCPR Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                    If Not bAchou Then
                        
                        Select Case objNFiscal.iClasseDocCPR
                        
                            'Consulta de títulos a pagar
                            Case CPR_TITULO_PAGAR
                            
                                Set objTituloPag = New ClassTituloPagar
                                
                                objTituloPag.lNumIntDoc = objNFiscal.lNumIntDocCPR
                                
                                lErro = RelPRJCustoReceita_Trata_TituloPag(objFlx.alComando(11), objTituloPag, objFluxo, objFlx.colFluxo)
                                If lErro <> SUCESSO Then gError 189654
            
                            'Consulta de títulos a receber
                            Case CPR_TITULO_RECEBER
                            
                                Set objTituloRec = New ClassTituloReceber
                                
                                objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
                                
                                lErro = RelPRJCustoReceita_Trata_TituloRec(objFlx.alComando(12), objTituloRec, objFluxo, objFlx.colFluxo)
                                If lErro <> SUCESSO Then gError 189655
                                
                            'Consulta de notas fiscais no cp (daqui o usuário pode ir para a tela de faturas)
                            Case CPR_NF_PAGAR
                            
                                Set objNFPag = New ClassNFsPag
                                
                                objNFPag.lNumIntDoc = objNFiscal.lNumIntDocCPR
                                
                                lErro = RelPRJCustoReceita_Trata_NFPag(objFlx.alComando(13), objNFPag, objFluxo, objFlx.colFluxo)
                                If lErro <> SUCESSO Then gError 189656
                            
                            Case Else
                                'Os demais casos não estão sendo considerados nesse momento
                              
                        End Select
                        
                    End If
                    
                End If
                
            Case PRJ_CR_TIPO_SAQUE
                'Valor real de custo
                
                Set objMovConta = objOrigem
                
                objFluxo.dtData = objMovConta.dtDataMovimento
                objFluxo.dDespesaReal = objPRJCR.dValor
                objFluxo.sDescricao = "Saque"
                objFluxo.sNatureza = objMovConta.sNatureza
            
                objFlx.colFluxo.Add objFluxo
            
            Case PRJ_CR_TIPO_DEPOSITO
                'Valor real de receita
            
                Set objMovConta = objOrigem
                
                objFluxo.dtData = objMovConta.dtDataMovimento
                objFluxo.dReceitaReal = objPRJCR.dValor
                objFluxo.sDescricao = "Depósito"
                objFluxo.sNatureza = objMovConta.sNatureza
                
                objFlx.colFluxo.Add objFluxo
            
            Case PRJ_CR_TIPO_TITREC
                'Pega os valores baixados como real e os a baixar como previsto de receita
            
                Set objTituloRec = objOrigem
            
                objFluxo.sNatureza = objTituloRec.sNatureza
           
                lErro = RelPRJCustoReceita_Trata_TituloRec(objFlx.alComando(12), objTituloRec, objFluxo, objFlx.colFluxo)
                If lErro <> SUCESSO Then gError 189657
                
            Case PRJ_CR_TIPO_TITPAG
                'Pega os valores baixados como real e os a baixar como previsto de despesa
            
                Set objTituloPag = objOrigem
            
                objFluxo.sNatureza = objTituloPag.sNatureza
            
                lErro = RelPRJCustoReceita_Trata_TituloPag(objFlx.alComando(11), objTituloPag, objFluxo, objFlx.colFluxo)
                If lErro <> SUCESSO Then gError 189658
                        
            Case PRJ_CR_TIPO_OV
                'Só informativo
            
            Case PRJ_CR_TIPO_PV
                'Se já virou nota desconsidera, senão é previsto de receita
            
                Set objPedidoVenda = objOrigem
            
                'objFluxo.sNatureza = objPedidoVenda.sNatureza
            
                lErro = RelPRJCustoReceita_Trata_PV(objPedidoVenda, objFluxo, objFlx.colFluxo)
                If lErro <> SUCESSO Then gError 189659
            
            Case PRJ_CR_TIPO_NFPAG
                'Se já virou fatura busca os dados dela, senão é previsto de despesa
            
                Set objNFPag = objOrigem
                
                lErro = RelPRJCustoReceita_Trata_NFPag(objFlx.alComando(14), objNFPag, objFluxo, objFlx.colFluxo)
                If lErro <> SUCESSO Then gError 189660
                        
            Case PRJ_CR_TIPO_OP
                'Só informativo\Custo de Produção
            
            Case PRJ_CR_TIPO_PRODENTRADA
                'Só informativo\Custo de Produção
            
            Case PRJ_CR_TIPO_REQPROD
                'Só informativo\Custo de Produção
                
        End Select
        
    Next
    
    For Each objFluxo In objFlx.colFluxo
        If Len(Trim(objFluxo.sNatureza)) <> 0 And Len(Trim(objFluxo.sDescNatureza)) = 0 Then
            bAchou = False
            For Each objNatureza In colNat
                If objNatureza.sCodigo = objFluxo.sNatureza Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objNatureza = New ClassNatMovCta
                objNatureza.sCodigo = objFluxo.sNatureza
                lErro = CF("NatMovCta_Le", objNatureza)
                If lErro <> SUCESSO And lErro <> 122786 Then gError ERRO_SEM_MENSAGEM
                colNat.Add objNatureza
            End If
            
            objFluxo.sDescNatureza = objNatureza.sDescricao
    
        End If
    Next
    
    RelPRJCustoReceita_Trata_TipoDocOrigem = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_TipoDocOrigem:

    RelPRJCustoReceita_Trata_TipoDocOrigem = gErr

    Select Case gErr

        Case 189652
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 189653 To 189660
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189660)

    End Select
    
    Exit Function

End Function

Public Function RelPRJCustoReceita_Trata_BaixasPag(ByVal lComando As Long, ByVal lNumIntTitulo As Long, ByVal objFluxoBase As ClassPRJFluxoFinanc, ByVal colFluxo As Collection, ByVal dPercentual As Double) As Long

Dim lErro As Long
Dim dtData As Date
Dim dValor As Double
Dim objFluxo As ClassPRJFluxoFinanc

On Error GoTo Erro_RelPRJCustoReceita_Trata_BaixasPag
                        
    lErro = Comando_Executar(lComando, "SELECT B.Data, BP.ValorBaixado + BP.ValorMulta + ValorJuros - BP.ValorDesconto AS Valor FROM BaixasParcPag AS BP, BaixasPag AS B, ParcelasPagTodas AS P WHERE BP.NumIntParcela = P.NumIntDoc AND BP.NumIntBaixa = B.NumIntBaixa AND B.Status <> 5 AND BP.Status <> 5 AND P.NumIntTitulo = ?", dtData, dValor, lNumIntTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 189661

    'le o primeiro Sequencial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189662

    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objFluxo = New ClassPRJFluxoFinanc
    
        Call RelPRJCustoReceita_Cria_Fluxo(objFluxo, objFluxoBase)
        
        objFluxo.dDespesaReal = dValor * dPercentual
        objFluxo.sDescricao = "Tit.Pag.(Baixa)"
        objFluxo.dtData = dtData
        
        colFluxo.Add objFluxo
        
        'lê o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189663
    
    Loop

    RelPRJCustoReceita_Trata_BaixasPag = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_BaixasPag:

    RelPRJCustoReceita_Trata_BaixasPag = gErr

    Select Case gErr
    
        Case 189661 To 189663
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASPAG", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189664)

    End Select

    Exit Function

End Function

Public Function RelPRJCustoReceita_Trata_BaixasRec(ByVal lComando As Long, ByVal lNumIntTitulo As Long, ByVal objFluxoBase As ClassPRJFluxoFinanc, ByVal colFluxo As Collection, ByVal dPercentual As Double) As Long

Dim lErro As Long
Dim dtData As Date
Dim dValor As Double
Dim objFluxo As ClassPRJFluxoFinanc

On Error GoTo Erro_RelPRJCustoReceita_Trata_BaixasRec
                        
    lErro = Comando_Executar(lComando, "SELECT B.Data, BP.ValorRecebido FROM BaixasParcRec AS BP, BaixasRec AS B, ParcelasRecTodas AS P WHERE BP.NumIntParcela = P.NumIntDoc AND BP.NumIntBaixa = B.NumIntBaixa AND B.Status <> 5 AND BP.Status <> 5 AND P.NumIntTitulo = ?", dtData, dValor, lNumIntTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 189665

    'le o primeiro Sequencial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189666

    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objFluxo = New ClassPRJFluxoFinanc
    
        Call RelPRJCustoReceita_Cria_Fluxo(objFluxo, objFluxoBase)
        
        objFluxo.dReceitaReal = dValor * dPercentual
        objFluxo.sDescricao = "Tit.Rec.(Baixa)"
        objFluxo.dtData = dtData
        
        colFluxo.Add objFluxo
        
        'lê o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189667
    
    Loop

    RelPRJCustoReceita_Trata_BaixasRec = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_BaixasRec:

    RelPRJCustoReceita_Trata_BaixasRec = gErr

    Select Case gErr
        
        Case 189665 To 189667
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASPAG", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189668)

    End Select

    Exit Function

End Function

Private Function RelPRJCustoReceita_Cria_Fluxo(ByVal objFluxo As ClassPRJFluxoFinanc, ByVal objFluxoBase As ClassPRJFluxoFinanc) As Long

Dim lErro As Long

On Error GoTo Erro_RelPRJCustoReceita_Cria_Fluxo

    objFluxo.sDocOrigem = objFluxoBase.sDocOrigem
    objFluxo.sNomeTela = objFluxoBase.sNomeTela
    objFluxo.sObs = objFluxoBase.sObs
    objFluxo.sNatureza = objFluxoBase.sNatureza
    
    Set objFluxo.objOrigem = objFluxoBase.objOrigem
    Set objFluxo.objPRJCR = objFluxoBase.objPRJCR
    Set objFluxo.objPRJInfo = objFluxoBase.objPRJInfo
    Set objFluxo.objProjeto = objFluxoBase.objProjeto
    Set objFluxo.objEtapa = objFluxoBase.objEtapa
    
    RelPRJCustoReceita_Cria_Fluxo = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Cria_Fluxo:

    RelPRJCustoReceita_Cria_Fluxo = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189669)

    End Select
    
    Exit Function
    
End Function

Private Function RelPRJCustoReceita_Trata_TituloRec(ByVal lComando As Long, ByVal objTituloRec As ClassTituloReceber, ByVal objFluxoBase As ClassPRJFluxoFinanc, ByVal colFluxo As Collection) As Long

Dim lErro As Long
Dim colParcelasRec As New ColParcelaReceber
Dim objParcelaRec As ClassParcelaReceber
Dim dPercentual As Double
Dim objFluxo As ClassPRJFluxoFinanc

On Error GoTo Erro_RelPRJCustoReceita_Trata_TituloRec

    lErro = CF("TituloReceber_Le", objTituloRec)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 189670
    If lErro <> SUCESSO Then
        lErro = CF("TituloReceberBaixado_Le", objTituloRec)
        If lErro <> SUCESSO And lErro <> 56570 Then gError 189671
    End If

    'Lê as Parcelas a Receber vinculadas ao Título
    lErro = CF("ParcelasReceber_Le_Todas", objTituloRec, colParcelasRec)
    If lErro <> SUCESSO And lErro <> 58990 Then gError 189672
        
    dPercentual = objFluxoBase.objPRJCR.dValor / objTituloRec.dValor
    
    For Each objParcelaRec In colParcelasRec

        If objParcelaRec.dSaldo * dPercentual > 0 Then
        
            Set objFluxo = New ClassPRJFluxoFinanc

            Call RelPRJCustoReceita_Cria_Fluxo(objFluxo, objFluxoBase)

            objFluxo.dtData = objParcelaRec.dtDataVencimentoReal
            objFluxo.dReceitaPrev = objParcelaRec.dSaldo * dPercentual
            objFluxo.sDescricao = "Tít.Rec.(Aberto)"
            
            colFluxo.Add objFluxo
        
        End If
        
    Next
    
    'Título a receber saldo baixado -> Receita Real
    lErro = RelPRJCustoReceita_Trata_BaixasRec(lComando, objTituloRec.lNumIntDoc, objFluxoBase, colFluxo, dPercentual)
    If lErro <> SUCESSO Then gError 189673
        
    RelPRJCustoReceita_Trata_TituloRec = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_TituloRec:

    RelPRJCustoReceita_Trata_TituloRec = gErr

    Select Case gErr
    
        Case 189670 To 189673
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189674)

    End Select
    
    Exit Function
    
End Function

Private Function RelPRJCustoReceita_Trata_TituloPag(ByVal lComando As Long, ByVal objTituloPag As ClassTituloPagar, ByVal objFluxoBase As ClassPRJFluxoFinanc, ByVal colFluxo As Collection) As Long

Dim lErro As Long
Dim colParcelasPag As New ColParcelaPagar
Dim objParcelaPag As ClassParcelaPagar
Dim dPercentual As Double
Dim objFluxo As ClassPRJFluxoFinanc

On Error GoTo Erro_RelPRJCustoReceita_Trata_TituloPag

    lErro = CF("TituloPagar_Le", objTituloPag)
    If lErro <> SUCESSO And lErro <> 18372 Then gError 189675
    If lErro <> SUCESSO Then
        lErro = CF("TituloPagarBaixado_Le", objTituloPag)
        If lErro <> SUCESSO And lErro <> 56661 Then gError 189676
    End If
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    lErro = CF("ParcelasPagar_Le_Todas", objTituloPag, colParcelasPag)
    If lErro <> SUCESSO Then gError 189677
        
    dPercentual = objFluxoBase.objPRJCR.dValor / objTituloPag.dValorTotal
                            
    For Each objParcelaPag In colParcelasPag
    
        If objParcelaPag.dSaldo * dPercentual > 0 Then
        
            Set objFluxo = New ClassPRJFluxoFinanc

            Call RelPRJCustoReceita_Cria_Fluxo(objFluxo, objFluxoBase)

            objFluxo.dtData = objParcelaPag.dtDataVencimentoReal
            objFluxo.dDespesaPrev = objParcelaPag.dSaldo * dPercentual
            objFluxo.sDescricao = "Tít.Pag.(Aberto)"
            
            colFluxo.Add objFluxo
        
        End If
        
    Next
                   
    'Pega as baixas do título como despesa real
    lErro = RelPRJCustoReceita_Trata_BaixasPag(lComando, objTituloPag.lNumIntDoc, objFluxoBase, colFluxo, dPercentual)
    If lErro <> SUCESSO Then gError 189678
        
    RelPRJCustoReceita_Trata_TituloPag = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_TituloPag:

    RelPRJCustoReceita_Trata_TituloPag = gErr

    Select Case gErr
    
        Case 189675 To 189678
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189679)

    End Select
    
    Exit Function
    
End Function

Private Function RelPRJCustoReceita_Trata_PV(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objFluxoBase As ClassPRJFluxoFinanc, ByVal colFluxo As Collection) As Long

Dim lErro As Long
Dim dPercentual As Double
Dim objFluxo As ClassPRJFluxoFinanc
Dim objItemPV As ClassItemPedido
Dim bFaturado As Boolean
Dim objParcelaPV As ClassParcelaPedidoVenda
Dim bAchou As Boolean
Dim dValor As Double
Dim dFator As Double
Dim dTotalParcelas As Double
Dim iItem As Integer

On Error GoTo Erro_RelPRJCustoReceita_Trata_PV

    dPercentual = 1

    bFaturado = False
    For Each objItemPV In objPedidoVenda.colItensPedido
        If objItemPV.dQuantFaturada > 0 Then
            bFaturado = True
        End If
    Next
           
    If IsNumeric(objFluxoBase.objPRJCR.sItem) Then
    
        bAchou = False
        iItem = 0
        For Each objItemPV In objPedidoVenda.colItensPedido
            iItem = iItem + 1
            If iItem = StrParaInt(objFluxoBase.objPRJCR.sItem) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then gError 189680

        'Percentual para abater o que já foi faturado
        dPercentual = 1 - objItemPV.dQuantFaturada / objItemPV.dQuantidade
        
        dValor = objFluxoBase.objPRJCR.dValor * dPercentual
    
    Else
    'Trata-se de seguro, frete, desconto ou outras despesas
        'Se ainda não foi faturado considera, senão ignora, já tem
        'que estar na nota fiscal
        If Not bFaturado Then
            dValor = objFluxoBase.objPRJCR.dValor
        Else
            dValor = 0
        End If
        
    End If
    
    dTotalParcelas = 0
    For Each objParcelaPV In objPedidoVenda.colParcelas
        dTotalParcelas = dTotalParcelas + objParcelaPV.dValor
    Next
    
    For Each objParcelaPV In objPedidoVenda.colParcelas
    
        dFator = objParcelaPV.dValor / dTotalParcelas
    
        If dValor * dFator > 0 Then
        
            Set objFluxo = New ClassPRJFluxoFinanc

            Call RelPRJCustoReceita_Cria_Fluxo(objFluxo, objFluxoBase)

            objFluxo.dtData = objParcelaPV.dtDataVencimento
            objFluxo.dReceitaPrev = dValor * dFator
            objFluxo.sDescricao = "PV não faturado."
            
            colFluxo.Add objFluxo
            
        End If
    Next
        
    RelPRJCustoReceita_Trata_PV = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_PV:

    RelPRJCustoReceita_Trata_PV = gErr

    Select Case gErr
    
        Case 189680
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJCR_ITEM_INEXISTENTE_PV", gErr, objFluxoBase.objPRJCR.sItem, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189681)

    End Select
    
    Exit Function
    
End Function

Private Function RelPRJCustoReceita_Trata_NFPag(ByVal lComando As Long, ByVal objNFPag As ClassNFsPag, ByVal objFluxoBase As ClassPRJFluxoFinanc, ByVal colFluxo As Collection) As Long

Dim lErro As Long
Dim dPercentual As Double
Dim objFluxo As ClassPRJFluxoFinanc
Dim objTituloPag As New ClassTituloPagar

On Error GoTo Erro_RelPRJCustoReceita_Trata_NFPag

    If objNFPag.lNumNotaFiscal = 0 Then
        lErro = CF("NFPag_Le_Todos2", objNFPag)
        If lErro <> SUCESSO And lErro <> 60139 Then gError 182058
    End If

    'Se tem data de vencimento
    If objNFPag.dtDataVencimento <> DATA_NULA Then
    
        'Se não está associado a um título
        If objNFPag.lNumIntTitPag = 0 Then
        
            Set objFluxo = New ClassPRJFluxoFinanc
    
            Call RelPRJCustoReceita_Cria_Fluxo(objFluxo, objFluxoBase)
    
            objFluxo.dtData = objNFPag.dtDataVencimento
            objFluxo.dDespesaPrev = objFluxoBase.objPRJCR.dValor
            objFluxo.sDescricao = "NFP não faturada"
            
            colFluxo.Add objFluxo
            
        Else
        
            objTituloPag.lNumIntDoc = objNFPag.lNumIntTitPag
            
            lErro = RelPRJCustoReceita_Trata_TituloPag(lComando, objTituloPag, objFluxoBase, colFluxo)
            If lErro <> SUCESSO Then gError 99999
            
        End If
        
    End If
 
    RelPRJCustoReceita_Trata_NFPag = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_NFPag:

    RelPRJCustoReceita_Trata_NFPag = gErr

    Select Case gErr
    
        Case 182058
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select
    
    Exit Function
    
End Function

Private Function RelPRJCustoReceita_Trata_Producao(ByVal objFlx As ClassFluxoPRJ) As Long

Dim lErro As Long
Dim objFluxoItem As ClassPRJFluxoFinanc
Dim objEtapa As ClassPRJEtapas
Dim objEtapaMO As ClassPRJEtapaMO
Dim objEtapaMP As ClassPRJEtapaMateriais
Dim objEtapaMaq As ClassPRJEtapaMaquinas

On Error GoTo Erro_RelPRJCustoReceita_Trata_Producao

    For Each objEtapa In objFlx.objProjeto.colEtapas
    
        For Each objEtapaMP In objEtapa.colMateriaPrima
        
            If (objEtapaMP.dtData >= objFlx.dtDataDe Or objFlx.dtDataDe = DATA_NULA) And (objEtapaMP.dtData <= objFlx.dtDataAte Or objFlx.dtDataAte = DATA_NULA) Then
        
                Set objFluxoItem = New ClassPRJFluxoFinanc
            
                objFluxoItem.sDocOrigem = "Apt (Insumo " & objEtapaMP.sProduto & SEPARADOR & objEtapaMP.sDescricao & ")"
            
                Set objFluxoItem.objProjeto = objFlx.objProjeto
                Set objFluxoItem.objEtapa = objEtapa
                Set objFluxoItem.objOrigem = objEtapa
                objFluxoItem.sNomeTela = "EtapaPRJ"
                objFluxoItem.sObs = objEtapaMP.sObservacao
                    
                Select Case objEtapaMP.iTipo
                
                    Case INDICE_INF_PREV
                        objFluxoItem.sDescricao = "Insumo previsto - Informado"
                        objFluxoItem.dtData = objEtapaMP.dtData
                        objFluxoItem.dCustoProdPrevInfo = objEtapaMP.dCusto
                    
                    Case INDICE_INF_REAL
                        objFluxoItem.sDescricao = "Insumo real - Informado"
                        objFluxoItem.dtData = objEtapaMP.dtData
                        objFluxoItem.dCustoProdRealInfo = objEtapaMP.dCusto
                    
                    Case INDICE_CALC_PREV
                        objFluxoItem.sDescricao = "Insumo previsto - Calculado"
                        objFluxoItem.dtData = objEtapaMP.dtData
                        objFluxoItem.dCustoProdPrevCalc = objEtapaMP.dCusto
                    
                    Case INDICE_CALC_REAL
                        objFluxoItem.sDescricao = "Insumo real - Calculado"
                        objFluxoItem.dtData = objEtapaMP.dtData
                        objFluxoItem.dCustoProdRealCalc = objEtapaMP.dCusto
                    
                End Select
                
                objFlx.colFluxo.Add objFluxoItem
                
            End If
            
        Next
        
        For Each objEtapaMO In objEtapa.colMaoDeObra
        
            If (objEtapaMO.dtData >= objFlx.dtDataDe Or objFlx.dtDataDe = DATA_NULA) And (objEtapaMO.dtData <= objFlx.dtDataAte Or objFlx.dtDataAte = DATA_NULA) Then
            
                Set objFluxoItem = New ClassPRJFluxoFinanc
            
                objFluxoItem.sDocOrigem = "Apt (Mão de Obra " & objEtapaMO.iMaoDeObra & SEPARADOR & objEtapaMO.sDescricao & ")"
                objFluxoItem.sObs = objEtapaMO.sObservacao
            
                Set objFluxoItem.objProjeto = objFlx.objProjeto
                Set objFluxoItem.objEtapa = objEtapa
                Set objFluxoItem.objOrigem = objEtapa
                objFluxoItem.sNomeTela = "EtapaPRJ"
                    
                Select Case objEtapaMO.iTipo
                
                    Case INDICE_INF_PREV
                        objFluxoItem.sDescricao = "Mão de Obra prevista - Informada"
                        objFluxoItem.dtData = objEtapaMO.dtData
                        objFluxoItem.dCustoProdPrevInfo = objEtapaMO.dCusto
                    
                    Case INDICE_INF_REAL
                        objFluxoItem.sDescricao = "Mão de Obra real - Informada"
                        objFluxoItem.dtData = objEtapaMO.dtData
                        objFluxoItem.dCustoProdRealInfo = objEtapaMO.dCusto
                    
                    Case INDICE_CALC_PREV
                        objFluxoItem.sDescricao = "Mão de Obra prevista - Calculada"
                        objFluxoItem.dtData = objEtapaMO.dtData
                        objFluxoItem.dCustoProdPrevCalc = objEtapaMO.dCusto
                    
                    Case INDICE_CALC_REAL
                        objFluxoItem.sDescricao = "Mão de Obra real - Calculada"
                        objFluxoItem.dtData = objEtapaMO.dtData
                        objFluxoItem.dCustoProdRealCalc = objEtapaMO.dCusto
                    
                End Select
                
                objFlx.colFluxo.Add objFluxoItem
                
            End If
            
        Next
        
        For Each objEtapaMaq In objEtapa.colMaquinas
        
            If (objEtapaMaq.dtData >= objFlx.dtDataDe Or objFlx.dtDataDe = DATA_NULA) And (objEtapaMaq.dtData <= objFlx.dtDataAte Or objFlx.dtDataAte = DATA_NULA) Then
            
                Set objFluxoItem = New ClassPRJFluxoFinanc
            
                objFluxoItem.sDocOrigem = "Apt (Máquinas " & objEtapaMaq.iMaquina & SEPARADOR & objEtapaMaq.sDescricao & ")"
                objFluxoItem.sObs = objEtapaMaq.sObservacao
                
                Set objFluxoItem.objProjeto = objFlx.objProjeto
                Set objFluxoItem.objEtapa = objEtapa
                Set objFluxoItem.objOrigem = objEtapa
                objFluxoItem.sNomeTela = "EtapaPRJ"
                    
                Select Case objEtapaMaq.iTipo
                
                    Case INDICE_INF_PREV
                        objFluxoItem.sDescricao = "Máquinas prevista - Informada"
                        objFluxoItem.dtData = objEtapaMaq.dtData
                        objFluxoItem.dCustoProdPrevInfo = objEtapaMaq.dCusto
                    
                    Case INDICE_INF_REAL
                        objFluxoItem.sDescricao = "Máquinas real - Informada"
                        objFluxoItem.dtData = objEtapaMaq.dtData
                        objFluxoItem.dCustoProdRealInfo = objEtapaMaq.dCusto
                    
                    Case INDICE_CALC_PREV
                        objFluxoItem.sDescricao = "Máquinas prevista - Calculada"
                        objFluxoItem.dtData = objEtapaMaq.dtData
                        objFluxoItem.dCustoProdPrevCalc = objEtapaMaq.dCusto
                    
                    Case INDICE_CALC_REAL
                        objFluxoItem.sDescricao = "Máquinas real - Calculada"
                        objFluxoItem.dtData = objEtapaMaq.dtData
                        objFluxoItem.dCustoProdRealCalc = objEtapaMaq.dCusto
                    
                End Select
                
                objFlx.colFluxo.Add objFluxoItem
                
            End If
            
        Next
        
    Next
 
    RelPRJCustoReceita_Trata_Producao = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_Producao:

    RelPRJCustoReceita_Trata_Producao = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select
    
    Exit Function
    
End Function

Public Function PRJRecebPagto_Le_Fluxo(ByVal objProjeto As ClassProjetos, ByVal colRecebPagto As Collection) As Long

Dim lErro As Long
Dim objRecebPagto As ClassPRJRecebPagto
Dim objRecebPagtoRegras As ClassPRJRecebPagtoRegras
Dim iIndice As Integer
Dim objCondicaoPagto As ClassCondicaoPagto
Dim objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_PRJRecebPagto_Le_Fluxo

    lErro = CF("PRJRecebPagto_Le_Projeto", objProjeto, colRecebPagto)
    If lErro <> SUCESSO Then gError 189683
    
    For iIndice = objProjeto.colRecebPagto.Count To 1 Step -1
        objProjeto.colRecebPagto.Remove iIndice
    Next
    
    For Each objRecebPagto In colRecebPagto
    
        objProjeto.colRecebPagto.Add objRecebPagto
    
        lErro = CF("RecebPagto_Calcula_Regras", objProjeto, objRecebPagto)
        If lErro <> SUCESSO Then gError 189684
        
        For Each objRecebPagtoRegras In objRecebPagto.colRegras
        
            Set objCondicaoPagto = New ClassCondicaoPagto
    
            objCondicaoPagto.iCodigo = objRecebPagtoRegras.iCondPagto
            
            'Lê a condição de pagamento
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then gError 189685
            
            'Calcula os valores das Parcelas
            objCondicaoPagto.dValorTotal = objRecebPagto.dValor * objRecebPagtoRegras.dPercentual
            objCondicaoPagto.dtDataRef = objRecebPagtoRegras.dtRegraValor
            
            'Calcula os valores das Parcelas
            lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
            If lErro <> SUCESSO Then gError 189686
            
            Set objRecebPagtoRegras.colParcelas = objCondicaoPagto.colParcelas
            
        Next
        
    Next
 
    PRJRecebPagto_Le_Fluxo = SUCESSO

    Exit Function

Erro_PRJRecebPagto_Le_Fluxo:

    PRJRecebPagto_Le_Fluxo = gErr

    Select Case gErr
    
        Case 189683 To 189686
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189687)

    End Select
    
    Exit Function
    
End Function

Private Function RelPRJCustoReceita_Trata_RecebPagto(ByVal objFlx As ClassFluxoPRJ) As Long

Dim lErro As Long
Dim objFluxo As ClassPRJFluxoFinanc
Dim colRecebPagto As New Collection
Dim objRecebPagto As ClassPRJRecebPagto
Dim objRecebPagtoRegras As ClassPRJRecebPagtoRegras
Dim iIndice As Integer
Dim objCondicaoPagto As ClassCondicaoPagto
Dim objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_RelPRJCustoReceita_Trata_RecebPagto

    lErro = CF("PRJRecebPagto_Le_Fluxo", objFlx.objProjeto, colRecebPagto)
    If lErro <> SUCESSO Then gError 189688
 
    For Each objRecebPagto In colRecebPagto
        For Each objRecebPagtoRegras In objRecebPagto.colRegras
            For Each objParc In objRecebPagtoRegras.colParcelas
                 If (objParc.dtVencimento >= objFlx.dtDataDe Or objFlx.dtDataDe = DATA_NULA) And (objParc.dtVencimento <= objFlx.dtDataAte Or objFlx.dtDataAte = DATA_NULA) Then
                   
                    Set objFluxo = New ClassPRJFluxoFinanc
                    Set objFluxo.objProjeto = objFlx.objProjeto
                    Set objFluxo.objOrigem = objRecebPagto
                    objFluxo.sObs = objRecebPagtoRegras.sObservacao
                    objFluxo.dtData = objParc.dtVencimento
                    
                    If objRecebPagto.iTipo = PRJ_TIPO_PAGTO Then
                        
                        objFluxo.sDocOrigem = "Pagamento " & CStr(objRecebPagto.lNumero)
                        
                        objFluxo.sNomeTela = "PagamentoPRJ"
    
                        objFluxo.sDescricao = "Pagamento"
                        objFluxo.dDespesaPlan = objParc.dValor
                        
                    Else
    
                        objFluxo.sDocOrigem = "Recebimento " & CStr(objRecebPagto.lNumero)
                        
                        objFluxo.sNomeTela = "RecebimentoPRJ"
    
                        objFluxo.sDescricao = "Recebimento"
                        objFluxo.dReceitaPlan = objParc.dValor
    
                    End If
                    objFlx.colFluxo.Add objFluxo
                    
                End If
            Next
        Next
    Next
 
    RelPRJCustoReceita_Trata_RecebPagto = SUCESSO

    Exit Function

Erro_RelPRJCustoReceita_Trata_RecebPagto:

    RelPRJCustoReceita_Trata_RecebPagto = gErr

    Select Case gErr
    
        Case 189688
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143931)

    End Select
    
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJContratos_Le_Projeto(ByVal objPRJContratos As ClassPRJContratos, Optional ByVal bLeItens As Boolean = True, Optional ByVal bLeTrib As Boolean = False) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJContratos As typePRJContratos

On Error GoTo Erro_PRJContratos_Le_Projeto

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 189690

    'Alocação de espaço no buffer
    tPRJContratos.sCodigo = String(STRING_CONTRATOPRJ_CODIGO, 0)
    tPRJContratos.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tPRJContratos.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)

    'Le a tabelaPRJContratos
    lErro = Comando_Executar(lComando, "SELECT NaturezaOP, NumIntDoc, NumIntDocPRJ,NumIntDocProposta, Codigo, Data, Observacao, Cliente, " & _
                "FilialCliente, ValorTotal, ValorProdutos, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, " & _
                "CustoInformado, CustoCalculado, ExibirProdutos, ExibirCustoCalc, ExibirCustoInfo, ExibirPreco FROM PRJContratos WHERE NumIntDocPRJ= ? ", _
                tPRJContratos.sNaturezaOp, tPRJContratos.lNumIntDoc, tPRJContratos.lNumIntDocPRJ, tPRJContratos.lNumIntDocProposta, tPRJContratos.sCodigo, tPRJContratos.dtData, _
                tPRJContratos.sObservacao, tPRJContratos.lCliente, tPRJContratos.iFilialCliente, tPRJContratos.dValorTotal, tPRJContratos.dValorProdutos, _
                tPRJContratos.dValorFrete, tPRJContratos.dValorDesconto, tPRJContratos.dValorSeguro, tPRJContratos.dValorOutrasDespesas, tPRJContratos.dCustoInformado, _
                tPRJContratos.dCustoCalculado, tPRJContratos.iExibirProdutos, tPRJContratos.iExibirCustoCalc, tPRJContratos.iExibirCustoInfo, tPRJContratos.iExibirPreco, _
                objPRJContratos.lNumIntDocPRJ)
    If lErro <> AD_SQL_SUCESSO Then gError 189691

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189692

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJContratos.lNumIntDoc = tPRJContratos.lNumIntDoc
    objPRJContratos.lNumIntDocPRJ = tPRJContratos.lNumIntDocPRJ
    objPRJContratos.lNumIntDocProposta = tPRJContratos.lNumIntDocProposta
    objPRJContratos.sCodigo = tPRJContratos.sCodigo
    objPRJContratos.dtData = tPRJContratos.dtData
    objPRJContratos.sObservacao = tPRJContratos.sObservacao
    objPRJContratos.lCliente = tPRJContratos.lCliente
    objPRJContratos.iFilialCliente = tPRJContratos.iFilialCliente
    objPRJContratos.dValorTotal = tPRJContratos.dValorTotal
    objPRJContratos.dValorProdutos = tPRJContratos.dValorProdutos
    objPRJContratos.dValorFrete = tPRJContratos.dValorFrete
    objPRJContratos.dValorDesconto = tPRJContratos.dValorDesconto
    objPRJContratos.dValorSeguro = tPRJContratos.dValorSeguro
    objPRJContratos.dValorOutrasDespesas = tPRJContratos.dValorOutrasDespesas
    objPRJContratos.dCustoInformado = tPRJContratos.dCustoInformado
    objPRJContratos.dCustoCalculado = tPRJContratos.dCustoCalculado
    objPRJContratos.iExibirCustoCalc = tPRJContratos.iExibirCustoCalc
    objPRJContratos.iExibirCustoInfo = tPRJContratos.iExibirCustoInfo
    objPRJContratos.iExibirPreco = tPRJContratos.iExibirPreco
    objPRJContratos.iExibirProdutos = tPRJContratos.iExibirProdutos
    objPRJContratos.sNaturezaOp = tPRJContratos.sNaturezaOp
    
    If bLeItens Then
    
        lErro = CF("PRJContratoItens_Le", objPRJContratos, bLeTrib)
        If lErro <> SUCESSO Then gError 189693
    
        lErro = CF("PRJContratoEtapa_Le", objPRJContratos)
        If lErro <> SUCESSO Then gError 189694
        
    End If
    
    If bLeTrib Then
    
        lErro = CF("PRJContratos_Le_Projeto_Tributacao", objPRJContratos)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189695
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJContratos_Le_Projeto = SUCESSO

    Exit Function

Erro_PRJContratos_Le_Projeto:

    PRJContratos_Le_Projeto = gErr

    Select Case gErr

        Case 189690
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189691, 189692
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJCONTRATOS", gErr)
            
        Case 189693, 189694, 189695

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189696)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ApontPRJ_Le(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lComando As Long
Dim tApontPRJ As typeApontPRJ

On Error GoTo Erro_ApontPRJ_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194437

    'Alocação de espaço no buffer
    tApontPRJ.sDescricao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tApontPRJ.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tApontPRJ.sUsuario = String(STRING_USUARIO, 0)

    'Le a tabelaApontPRJ
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NumIntDocPRJ, NumIntDocEtapa, Data, Descricao, " & _
                "Observacao, Usuario, DataRegistro, HoraRegistro FROM ApontPRJ WHERE Codigo= ? ", _
                tApontPRJ.lNumIntDoc, tApontPRJ.lCodigo, tApontPRJ.lNumIntDocPRJ, tApontPRJ.lNumIntDocEtapa, _
                tApontPRJ.dtData, tApontPRJ.sDescricao, tApontPRJ.sObservacao, tApontPRJ.sUsuario, tApontPRJ.dtDataRegistro, _
                tApontPRJ.dHoraRegistro, _
                objApontPRJ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 194438

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194439

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objApontPRJ.lNumIntDoc = tApontPRJ.lNumIntDoc
    objApontPRJ.lCodigo = tApontPRJ.lCodigo
    objApontPRJ.lNumIntDocPRJ = tApontPRJ.lNumIntDocPRJ
    objApontPRJ.lNumIntDocEtapa = tApontPRJ.lNumIntDocEtapa
    objApontPRJ.dtData = tApontPRJ.dtData
    objApontPRJ.sDescricao = tApontPRJ.sDescricao
    objApontPRJ.sObservacao = tApontPRJ.sObservacao
    objApontPRJ.sUsuario = tApontPRJ.sUsuario
    objApontPRJ.dtDataRegistro = tApontPRJ.dtDataRegistro
    objApontPRJ.dHoraRegistro = tApontPRJ.dHoraRegistro
    
    lErro = ApontMaqPRJ_Le(objApontPRJ)
    If lErro <> SUCESSO Then gError 194440

    lErro = ApontMOPRJ_Le(objApontPRJ)
    If lErro <> SUCESSO Then gError 194441

    lErro = ApontProdPRJ_Le(objApontPRJ)
    If lErro <> SUCESSO Then gError 194442

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ApontPRJ_Le = SUCESSO

    Exit Function

Erro_ApontPRJ_Le:

    ApontPRJ_Le = gErr

    Select Case gErr

        Case 194437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194438, 194439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTPRJ", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora
        
        Case 194440 To 194442

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194443)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function ApontProdPRJ_Le(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lComando As Long
Dim tApontProdPRJ As typeApontProdPRJ
Dim objApontProdPRJ As ClassApontProdPRJ

On Error GoTo Erro_ApontProdPRJ_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194444

    'Alocação de espaço no buffer
    tApontProdPRJ.sProduto = String(STRING_PRODUTO, 0)
    tApontProdPRJ.sUM = String(STRING_UM_SIGLA, 0)
    tApontProdPRJ.sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaApontProdPRJ
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocApont, Seq, Produto, UM, Qtd, " & _
                "Custo, Obs FROM ApontProdPRJ WHERE NumIntDocApont= ? ORDER BY Seq", _
                tApontProdPRJ.lNumIntDoc, tApontProdPRJ.lNumIntDocApont, tApontProdPRJ.iSeq, tApontProdPRJ.sProduto, _
                tApontProdPRJ.sUM, tApontProdPRJ.dQtd, tApontProdPRJ.dCusto, tApontProdPRJ.sObs, _
                objApontPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194445

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194446

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objApontProdPRJ = New ClassApontProdPRJ

        objApontProdPRJ.lNumIntDoc = tApontProdPRJ.lNumIntDoc
        objApontProdPRJ.lNumIntDocApont = tApontProdPRJ.lNumIntDocApont
        objApontProdPRJ.iSeq = tApontProdPRJ.iSeq
        objApontProdPRJ.sProduto = tApontProdPRJ.sProduto
        objApontProdPRJ.sUM = tApontProdPRJ.sUM
        objApontProdPRJ.dQtd = tApontProdPRJ.dQtd
        objApontProdPRJ.dCusto = tApontProdPRJ.dCusto
        objApontProdPRJ.sObs = tApontProdPRJ.sObs
        
        objApontPRJ.colMateriaPrima.Add objApontProdPRJ
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194447
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ApontProdPRJ_Le = SUCESSO

    Exit Function

Erro_ApontProdPRJ_Le:

    ApontProdPRJ_Le = gErr

    Select Case gErr

        Case 194444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194445, 194446, 194447
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTPRODPRJ", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194448)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function ApontMOPRJ_Le(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lComando As Long
Dim tApontMOPRJ As typeApontMOPRJ
Dim objApontMOPRJ As ClassApontMOPRJ

On Error GoTo Erro_ApontMOPRJ_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194449

    'Alocação de espaço no buffer
    tApontMOPRJ.sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaApontMOPRJ
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocApont, Seq, CodMO, Horas, Qtd, " & _
                "Custo, Obs FROM ApontMOPRJ WHERE NumIntDocApont= ? ORDER BY Seq", _
                tApontMOPRJ.lNumIntDoc, tApontMOPRJ.lNumIntDocApont, tApontMOPRJ.iSeq, tApontMOPRJ.iCodMO, _
                tApontMOPRJ.dHoras, tApontMOPRJ.iQtd, tApontMOPRJ.dCusto, tApontMOPRJ.sObs, _
                objApontPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194450

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194451

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objApontMOPRJ = New ClassApontMOPRJ
        
        objApontMOPRJ.lNumIntDoc = tApontMOPRJ.lNumIntDoc
        objApontMOPRJ.lNumIntDocApont = tApontMOPRJ.lNumIntDocApont
        objApontMOPRJ.iSeq = tApontMOPRJ.iSeq
        objApontMOPRJ.iCodMO = tApontMOPRJ.iCodMO
        objApontMOPRJ.dHoras = tApontMOPRJ.dHoras
        objApontMOPRJ.iQtd = tApontMOPRJ.iQtd
        objApontMOPRJ.dCusto = tApontMOPRJ.dCusto
        objApontMOPRJ.sObs = tApontMOPRJ.sObs
        
        objApontPRJ.colMaoDeObra.Add objApontMOPRJ
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194452
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ApontMOPRJ_Le = SUCESSO

    Exit Function

Erro_ApontMOPRJ_Le:

    ApontMOPRJ_Le = gErr

    Select Case gErr

        Case 194449
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194450, 194451, 194452
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTMOPRJ", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194453)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function ApontMaqPRJ_Le(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim lComando As Long
Dim tApontMaqPRJ As typeApontMaqPRJ
Dim objApontMaqPRJ As ClassApontMaqPRJ

On Error GoTo Erro_ApontMaqPRJ_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 194454

    'Alocação de espaço no buffer
    tApontMaqPRJ.sObs = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaApontMaqPRJ
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocApont, Seq, CodMaq, Horas, Qtd, " & _
                "Custo, Obs FROM ApontMaqPRJ WHERE NumIntDocApont= ? ORDER BY Seq", _
                tApontMaqPRJ.lNumIntDoc, tApontMaqPRJ.lNumIntDocApont, tApontMaqPRJ.iSeq, tApontMaqPRJ.iCodMaq, _
                tApontMaqPRJ.dHoras, tApontMaqPRJ.iQtd, tApontMaqPRJ.dCusto, tApontMaqPRJ.sObs, _
                objApontPRJ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 194455

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194456

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objApontMaqPRJ = New ClassApontMaqPRJ

        objApontMaqPRJ.lNumIntDoc = tApontMaqPRJ.lNumIntDoc
        objApontMaqPRJ.lNumIntDocApont = tApontMaqPRJ.lNumIntDocApont
        objApontMaqPRJ.iSeq = tApontMaqPRJ.iSeq
        objApontMaqPRJ.iCodMaq = tApontMaqPRJ.iCodMaq
        objApontMaqPRJ.dHoras = tApontMaqPRJ.dHoras
        objApontMaqPRJ.iQtd = tApontMaqPRJ.iQtd
        objApontMaqPRJ.dCusto = tApontMaqPRJ.dCusto
        objApontMaqPRJ.sObs = tApontMaqPRJ.sObs
        
        objApontPRJ.colMaquinas.Add objApontMaqPRJ
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194457
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ApontMaqPRJ_Le = SUCESSO

    Exit Function

Erro_ApontMaqPRJ_Le:

    ApontMaqPRJ_Le = gErr

    Select Case gErr

        Case 194454
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 194455, 194456, 194457
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTMAQPRJ", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194458)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Projetos_Le_Customizado(ByVal objProjetos As ClassProjetos) As Long
    Projetos_Le_Customizado = SUCESSO
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function PRJEtapaVistorias_Le(ByVal objPRJEtapaVistorias As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJEtapaVistorias As typePRJEtapaVistorias

On Error GoTo Erro_PRJEtapaVistorias_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213609

    'Alocação de espaço no buffer
    tPRJEtapaVistorias.sResponsavel = String(STRING_MAXIMO, 0)
    tPRJEtapaVistorias.sLaudo = String(STRING_MAXIMO, 0)

    'Le a tabelaPRJEtapaVistorias
    If objPRJEtapaVistorias.lNumIntDoc <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntPRJEtapa, Codigo, Data, DataValidade, Responsavel, " & _
                    "Laudo FROM PRJEtapaVistorias WHERE NumIntDoc= ? ", _
                    tPRJEtapaVistorias.lNumIntDoc, tPRJEtapaVistorias.lNumIntPRJEtapa, tPRJEtapaVistorias.lCodigo, tPRJEtapaVistorias.dtData, _
                    tPRJEtapaVistorias.dtDataValidade, tPRJEtapaVistorias.sResponsavel, tPRJEtapaVistorias.sLaudo, _
                    objPRJEtapaVistorias.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntPRJEtapa, Codigo, Data, DataValidade, Responsavel, " & _
                    "Laudo FROM PRJEtapaVistorias WHERE NumIntPRJEtapa= ? AND Codigo = ? ", _
                    tPRJEtapaVistorias.lNumIntDoc, tPRJEtapaVistorias.lNumIntPRJEtapa, tPRJEtapaVistorias.lCodigo, tPRJEtapaVistorias.dtData, _
                    tPRJEtapaVistorias.dtDataValidade, tPRJEtapaVistorias.sResponsavel, tPRJEtapaVistorias.sLaudo, _
                    objPRJEtapaVistorias.lNumIntPRJEtapa, objPRJEtapaVistorias.lCodigo)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 213610

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213611

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objPRJEtapaVistorias.lNumIntDoc = tPRJEtapaVistorias.lNumIntDoc
    objPRJEtapaVistorias.lNumIntPRJEtapa = tPRJEtapaVistorias.lNumIntPRJEtapa
    objPRJEtapaVistorias.lCodigo = tPRJEtapaVistorias.lCodigo
    objPRJEtapaVistorias.dtData = tPRJEtapaVistorias.dtData
    objPRJEtapaVistorias.dtDataValidade = tPRJEtapaVistorias.dtDataValidade
    objPRJEtapaVistorias.sResponsavel = tPRJEtapaVistorias.sResponsavel
    objPRJEtapaVistorias.sLaudo = tPRJEtapaVistorias.sLaudo

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PRJEtapaVistorias_Le = SUCESSO

    Exit Function

Erro_PRJEtapaVistorias_Le:

    PRJEtapaVistorias_Le = gErr

    Select Case gErr

        Case 213609
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213610, 213611
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAVISTORIAS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213612)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function VistoriaPRJ_Lista_Validade_Texto(ByVal iDiasParaExpirar As Integer, iNumVistAExpirar As Integer, sTexto As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim dtDataValidAte As Date
Dim tPRJ As typeProjetos
Dim tEtapa As typePRJEtapas
Dim sPRJMask As String
Dim sTextoAux As String

On Error GoTo Erro_VistoriaPRJ_Lista_Validade_Texto

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213669
    
    dtDataValidAte = DateAdd("d", iDiasParaExpirar, Date)

    'Alocação de espaço no buffer
    tPRJ.sCodigo = String(STRING_MAXIMO, 0)
    tPRJ.sDescricao = String(STRING_MAXIMO, 0)
    tPRJ.sResponsavel = String(STRING_MAXIMO, 0)

    tEtapa.sCodigo = String(STRING_MAXIMO, 0)
    tEtapa.sDescricao = String(STRING_MAXIMO, 0)
    tEtapa.sResponsavel = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT P.Codigo, P.Descricao, P.Responsavel, E.Codigo, E.Descricao, E.Responsavel, E.DataVistoria, E.ValidadeVistoria FROM Projetos AS P, PRJEtapas As E WHERE P.NumIntDoc = E.NumIntDocPRJ AND E.PercentualComplet < 1 AND E.DataVistoria <> ? AND E.ValidadeVistoria < ? ORDER BY ValidadeVistoria DESC ", _
                tPRJ.sCodigo, tPRJ.sDescricao, tPRJ.sResponsavel, tEtapa.sCodigo, tEtapa.sDescricao, tEtapa.sResponsavel, tEtapa.dtDataVistoria, tEtapa.dtValidadeVistoria, DATA_NULA, dtDataValidAte)
    If lErro <> AD_SQL_SUCESSO Then gError 213670

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213671

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Call Retorno_Projeto_Tela2(tPRJ.sCodigo, sPRJMask)
    
        iNumVistAExpirar = iNumVistAExpirar + 1
        sTextoAux = sTextoAux & "Projeto    : " & sPRJMask & SEPARADOR & tPRJ.sDescricao & vbNewLine
        sTextoAux = sTextoAux & "Etapa      : " & tEtapa.sCodigo & SEPARADOR & tEtapa.sDescricao & vbNewLine
        sTextoAux = sTextoAux & "Responsável: " & IIf(tEtapa.sResponsavel <> "", tEtapa.sResponsavel, tPRJ.sResponsavel) & vbNewLine
        sTextoAux = sTextoAux & "Vistoria   : De " & Format(tEtapa.dtDataVistoria, "dd/mm/yyyy") & " válida até " & Format(tEtapa.dtValidadeVistoria, "dd/mm/yyyy") & " (" & DateDiff("d", Date, tEtapa.dtValidadeVistoria) & " dias)" & vbNewLine & vbNewLine

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213672
    
    Loop
    
    sTexto = "Existe(m) " & CStr(iNumVistAExpirar) & " vistoria(s) a expirar nos próximos " & CStr(iDiasParaExpirar) & " dias. Segue relação abaixo: " & vbNewLine & vbNewLine
    sTexto = sTexto & sTextoAux

    'Fecha Comando
    Call Comando_Fechar(lComando)

    VistoriaPRJ_Lista_Validade_Texto = SUCESSO

    Exit Function

Erro_VistoriaPRJ_Lista_Validade_Texto:

    VistoriaPRJ_Lista_Validade_Texto = gErr

    Select Case gErr

        Case 213669
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213670 To 213672
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAVISTORIAS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213673)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function EtapasPRJ_Sem_Vistorias_Texto(iNumEtapasSemVist As Integer, sTexto As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPRJ As typeProjetos
Dim tEtapa As typePRJEtapas
Dim sPRJMask As String
Dim sTextoAux As String

On Error GoTo Erro_EtapasPRJ_Sem_Vistorias_Texto

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213664

    'Alocação de espaço no buffer
    tPRJ.sCodigo = String(STRING_MAXIMO, 0)
    tPRJ.sDescricao = String(STRING_MAXIMO, 0)
    tPRJ.sResponsavel = String(STRING_MAXIMO, 0)

    tEtapa.sCodigo = String(STRING_MAXIMO, 0)
    tEtapa.sDescricao = String(STRING_MAXIMO, 0)
    tEtapa.sResponsavel = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT P.Codigo, P.Descricao, P.Responsavel, E.Codigo, E.Descricao, E.Responsavel FROM Projetos AS P, PRJEtapas As E WHERE P.NumIntDoc = E.NumIntDocPRJ AND E.PercentualComplet < 1 AND E.ValidadeVistoria = ? ORDER BY P.Codigo, E.Codigo ", _
                tPRJ.sCodigo, tPRJ.sDescricao, tPRJ.sResponsavel, tEtapa.sCodigo, tEtapa.sDescricao, tEtapa.sResponsavel, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 213665

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213666

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Call Retorno_Projeto_Tela2(tPRJ.sCodigo, sPRJMask)
    
        iNumEtapasSemVist = iNumEtapasSemVist + 1
        sTextoAux = sTextoAux & "Projeto    : " & sPRJMask & SEPARADOR & tPRJ.sDescricao & vbNewLine
        sTextoAux = sTextoAux & "Etapa      : " & tEtapa.sCodigo & SEPARADOR & tEtapa.sDescricao & vbNewLine
        sTextoAux = sTextoAux & "Responsável: " & IIf(tEtapa.sResponsavel <> "", tEtapa.sResponsavel, tPRJ.sResponsavel) & vbNewLine

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213667
    
    Loop
    
    sTexto = "Existe(m) " & CStr(iNumEtapasSemVist) & " etapa(s) sem vistorias cadastradas. Segue relação abaixo: " & vbNewLine & vbNewLine
    sTexto = sTexto & sTextoAux

    'Fecha Comando
    Call Comando_Fechar(lComando)

    EtapasPRJ_Sem_Vistorias_Texto = SUCESSO

    Exit Function

Erro_EtapasPRJ_Sem_Vistorias_Texto:

    EtapasPRJ_Sem_Vistorias_Texto = gErr

    Select Case gErr

        Case 213664
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213665 To 213667
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRJETAPAVISTORIAS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213668)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

