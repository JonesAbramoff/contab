VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContabInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public mvargiExercicioAtual As Integer
Public mvargobjColOrigem As ClassColOrigemContab
Public mvargobjColNaturezaConta As ClassColNaturezaConta
Public mvargobjColTipoConta As ClassColTipoConta
Public mvargobjColCodigoSegmento As ClassColCodigoSegmento
Public mvargobjColPreenchimento As ClassColPreenchimento
Public mvargobjColTipoSegmento As ClassColTipoSegmento
Public mvargsOrigemAtual As String
Public mvargiPeriodoAtual As Integer
Public mvargiSetupLotePorPeriodo As Integer
Public mvargiSetupDocPorPeriodo As Integer
Public mvargiSetupUsoCcl As Integer
Public mvargsConta As String
Public mvargsCcl As String
Public mvargiHistPadrao As Integer
Public mvargiTipoConta As Integer
Public mvargiNaturezaConta As Integer
Public mvargiRateioOn As Integer
Public mvargobjRateioOff As ClassRateioOff
Public mvargiContabCentralizada As Integer
Public mvargiHistoricoObrigatorio As Integer
Public mvargiContabGerencial As Integer
Public mvargiFilialAuxiliar As Integer

Public mvariCPAglutinaLancamPorDia As Integer
Public mvariCPGeraLoteAutomatico As Integer
Public mvariCRAglutinaLancamPorDia As Integer
Public mvariCRGeraLoteAutomatico As Integer
Public mvariTESAglutinaLancamPorDia As Integer
Public mvariTESGeraLoteAutomatico As Integer
Public mvariFATAglutinaLancamPorDia As Integer
Public mvariFATGeraLoteAutomatico As Integer
Public mvariESTAglutinaLancamPorDia As Integer
Public mvariESTGeraLoteAutomatico As Integer

Public mvariESTTamanhoMinimoProdutoGrid As Integer
Public mvariESTTamanhoMinimoContaGrid As Integer
Public mvargiValidaCtaCcl As Integer
Public mvargiCtbLctoExcluiEstorno As Integer

Private Sub Class_Initialize()

Dim lErro As Long

On Error GoTo Erro_ClassContab_Class_Initialize
    
    Call Init_Contab
    Set mvargobjRateioOff = New ClassRateioOff
    Set mvargobjColOrigem = New ClassColOrigemContab
    Set mvargobjColNaturezaConta = New ClassColNaturezaConta
    Set mvargobjColTipoConta = New ClassColTipoConta
    Set mvargobjColCodigoSegmento = New ClassColCodigoSegmento
    Set mvargobjColPreenchimento = New ClassColPreenchimento
    Set mvargobjColTipoSegmento = New ClassColTipoSegmento

    mvargsOrigemAtual = "CTB"

    lErro = Inicializa_Variaveis_Setup()
    If lErro <> SUCESSO Then Error 5406
        
    lErro = Inicializa()
    If lErro <> SUCESSO Then Error 36456
    
    Exit Sub
    
Erro_ClassContab_Class_Initialize:

    Select Case Err

        Case 5406, 36456
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147875)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Class_Terminate()

    Set mvargobjRateioOff = Nothing
    Set mvargobjColOrigem = Nothing
    Set mvargobjColNaturezaConta = Nothing
    Set mvargobjColTipoConta = Nothing
    Set mvargobjColCodigoSegmento = Nothing
    Set mvargobjColPreenchimento = Nothing
    Set mvargobjColTipoSegmento = Nothing

End Sub

Private Function Inicializa_Variaveis_Setup() As Long

Dim lErro As Long
Dim objConfiguracao As New ClassConfiguracao
Dim objCTBConfig As New ClassCTBConfig
    
On Error GoTo Erro_Inicializa_Variaveis_Setup
    
    lErro = CF("Configuracao_Le", objConfiguracao)
    If lErro <> SUCESSO Then Error 9276
    
    mvargiSetupUsoCcl = objConfiguracao.iUsoCcl
    mvargiSetupLotePorPeriodo = objConfiguracao.iLotePorPeriodo
    mvargiSetupDocPorPeriodo = objConfiguracao.iDocPorPeriodo
    mvargiTipoConta = objConfiguracao.iTipoContaDefault
    mvargiNaturezaConta = objConfiguracao.iNaturezaDefault
    mvargiContabCentralizada = objConfiguracao.iContabCentralizada
    giContabGerencial = objConfiguracao.iContabGerencial
    giFilialAuxiliar = objConfiguracao.iFilialAuxiliar
    
    '################################################
    'Inserido or Wagner 10/03/06
    'obter conta de transferencia de CTBConfig
    objCTBConfig.sCodigo = HISTORICO_CTB_OBRIGATORIO
    
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 9276
    
    mvargiHistoricoObrigatorio = StrParaInt(objCTBConfig.sConteudo)
    
    objCTBConfig.sCodigo = VALIDA_CONTA_CCL
    
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 9276
    
    mvargiValidaCtaCcl = StrParaInt(objCTBConfig.sConteudo)
    '################################################
    
    objCTBConfig.sCodigo = CTB_LANCAMENTO_EXCLUI_EFETIVO_ESTORNO
    
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then Error 9276
    
    mvargiCtbLctoExcluiEstorno = StrParaInt(objCTBConfig.sConteudo)
    
   
    lErro = Configura_DesconsideraFechtoPeriodo()
    If lErro <> SUCESSO Then Error 9276
    
    Inicializa_Variaveis_Setup = SUCESSO

    Exit Function
    
Erro_Inicializa_Variaveis_Setup:

    Inicializa_Variaveis_Setup = Err
    
    Select Case Err

        Case 9276
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147876)
        
    End Select
    
    Exit Function
    
End Function

Private Function Inicializa() As Long

Dim lErro As Long
Dim colConfig As Object

On Error GoTo Erro_Inicializa

    lErro = Inicializa_CP()
    If lErro <> SUCESSO Then Error 36457

    lErro = Inicializa_CR()
    If lErro <> SUCESSO Then Error 36458

    lErro = Inicializa_TES()
    If lErro <> SUCESSO Then Error 36459

    lErro = Inicializa_FAT()
    If lErro <> SUCESSO Then Error 36460

    lErro = Inicializa_EST()
    If lErro <> SUCESSO Then Error 36461

    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 36457, 36458, 36459, 36460, 36461
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147877)
        
    End Select

    Exit Function

End Function

Private Function Inicializa_CP() As Long

Dim lErro As Long
Dim colConfig As Object

On Error GoTo Erro_Inicializa_CP

    'If gcolModulo.Ativo(MODULO_CONTASAPAGAR) = MODULO_ATIVO Then

        Set colConfig = CreateObject("GlobaisCPR.ColCPConfig")
    
        colConfig.Add CPCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", CPCFG_AGLUTINA_LANCAM_POR_DIA
        colConfig.Add CPCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", CPCFG_GERA_LOTE_AUTOMATICO
    
        'Lê as configurações em CPConfig
        lErro = CF("CPConfig_Le_Configs", colConfig)
        If lErro <> SUCESSO Then Error 36462
    
        mvariCPAglutinaLancamPorDia = CInt(colConfig.Item(CPCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
        mvariCPGeraLoteAutomatico = CInt(colConfig.Item(CPCFG_GERA_LOTE_AUTOMATICO).sConteudo)
        
    'End If

    Inicializa_CP = SUCESSO
    
    Exit Function
    
Erro_Inicializa_CP:

    Inicializa_CP = Err
    
    Select Case Err
    
        Case 36462
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147878)
        
    End Select

    Exit Function

End Function

Private Function Inicializa_CR() As Long

Dim lErro As Long
Dim colConfig As Object

On Error GoTo Erro_Inicializa_CR

    'If gcolModulo.Ativo(MODULO_CONTASARECEBER) = MODULO_ATIVO Then

        Set colConfig = CreateObject("GlobaisCPR.ColCRConfig")

        colConfig.Add CRCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", CRCFG_AGLUTINA_LANCAM_POR_DIA
        colConfig.Add CRCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", CRCFG_GERA_LOTE_AUTOMATICO
    
        'Lê as configurações em CRConfig
        lErro = CF("CRConfig_Le_Configs", colConfig)
        If lErro <> SUCESSO Then Error 36463
    
        mvariCRAglutinaLancamPorDia = CInt(colConfig.Item(CRCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
        mvariCRGeraLoteAutomatico = CInt(colConfig.Item(CRCFG_GERA_LOTE_AUTOMATICO).sConteudo)
        
    'End If

    Inicializa_CR = SUCESSO
    
    Exit Function
    
Erro_Inicializa_CR:

    Inicializa_CR = Err
    
    Select Case Err
    
        Case 36463
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147879)
        
    End Select

    Exit Function

End Function

Private Function Inicializa_TES() As Long

Dim lErro As Long
Dim colConfig As Object

On Error GoTo Erro_Inicializa_TES

    'If gcolModulo.Ativo(MODULO_TESOURARIA) = MODULO_ATIVO Then

        Set colConfig = CreateObject("GlobaisCPR.ColTESConfig")

        colConfig.Add TESCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", TESCFG_AGLUTINA_LANCAM_POR_DIA
        colConfig.Add TESCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", TESCFG_GERA_LOTE_AUTOMATICO
        
        'Lê as configurações em TESConfig
        lErro = CF("TESConfig_Le_Configs", colConfig)
        If lErro <> SUCESSO Then Error 36464
        
        mvariTESAglutinaLancamPorDia = CInt(colConfig.Item(TESCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
        mvariTESGeraLoteAutomatico = CInt(colConfig.Item(TESCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    
    'End If

    Inicializa_TES = SUCESSO
    
    Exit Function
    
Erro_Inicializa_TES:

    Inicializa_TES = Err
    
    Select Case Err
    
        Case 36464
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147880)
        
    End Select

    Exit Function

End Function

Private Function Inicializa_FAT() As Long

Dim lErro As Long
Dim colConfig As Object

On Error GoTo Erro_Inicializa_FAT

    'If gcolModulo.Ativo(MODULO_FATURAMENTO) = MODULO_ATIVO Then
    
        Set colConfig = CreateObject("GlobaisFAT.ColFATConfig")
        
        colConfig.Add FATCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", FATCFG_AGLUTINA_LANCAM_POR_DIA
        colConfig.Add FATCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", FATCFG_GERA_LOTE_AUTOMATICO
        
        'Lê as configurações em FATConfig
        lErro = CF("FATConfig_Le_Configs", colConfig)
        If lErro <> SUCESSO Then Error 36465
        
        mvariFATAglutinaLancamPorDia = CInt(colConfig.Item(FATCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
        mvariFATGeraLoteAutomatico = CInt(colConfig.Item(FATCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    
    'End If

    Inicializa_FAT = SUCESSO
    
    Exit Function
    
Erro_Inicializa_FAT:

    Inicializa_FAT = Err
    
    Select Case Err
    
        Case 36465
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147881)
        
    End Select

    Exit Function

End Function

Private Function Inicializa_EST() As Long

Dim lErro As Long
Dim colConfig As Object

On Error GoTo Erro_Inicializa_EST

    'If gcolModulo.Ativo(MODULO_ESTOQUE) = MODULO_ATIVO Then
    
        Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
        
        colConfig.Add ESTCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", ESTCFG_AGLUTINA_LANCAM_POR_DIA
        colConfig.Add ESTCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", ESTCFG_GERA_LOTE_AUTOMATICO
        colConfig.Add ESTCFG_TAMANHO_MINIMO_PRODUTO_GRID, EMPRESA_TODA, "", 0, "", ESTCFG_TAMANHO_MINIMO_PRODUTO_GRID
        colConfig.Add ESTCFG_TAMANHO_MINIMO_CONTA_GRID, EMPRESA_TODA, "", 0, "", ESTCFG_TAMANHO_MINIMO_CONTA_GRID
        
        'Lê as configurações em ESTConfig
        lErro = CF("ESTConfig_Le_Configs", colConfig)
        If lErro <> SUCESSO Then Error 36466
        
        mvariESTAglutinaLancamPorDia = CInt(colConfig.Item(ESTCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
        mvariESTGeraLoteAutomatico = CInt(colConfig.Item(ESTCFG_GERA_LOTE_AUTOMATICO).sConteudo)
        mvariESTTamanhoMinimoContaGrid = CInt(colConfig.Item(ESTCFG_TAMANHO_MINIMO_CONTA_GRID).sConteudo)
        mvariESTTamanhoMinimoProdutoGrid = CInt(colConfig.Item(ESTCFG_TAMANHO_MINIMO_PRODUTO_GRID).sConteudo)
    'End If

    Inicializa_EST = SUCESSO
    
    Exit Function
    
Erro_Inicializa_EST:

    Inicializa_EST = Err
    
    Select Case Err
    
        Case 36466
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147882)
        
    End Select

    Exit Function

End Function

Private Function Configura_DesconsideraFechtoPeriodo() As Long
'marca se deve ou nao desconsiderar o fechamento de periodo contabil (e exercicio) para bloquear gravacao de lctos contabeis

Dim lErro As Long, iTipoDeAcesso As Integer
Dim lComando As Long

On Error GoTo Erro_Configura_DesconsideraFechtoPeriodo

    giDesconsideraFechamentoPeriodo = DESMARCADO

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 201294
    
    lErro = Comando_Executar(lComando, "SELECT TipoDeAcesso FROM Rotinas R, GrupoRotinas G, Usuarios U where U.CodUsuario = ? AND R.Sigla = 'DesconsideraFechtoPeriodoCtb' AND G.SiglaRotina = R.Sigla AND G.CodGrupo = U.CodGrupo AND G.TipodeAcesso=1", iTipoDeAcesso, gsUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 201295
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201296
    
    If lErro = AD_SQL_SUCESSO Then giDesconsideraFechamentoPeriodo = MARCADO
    
    Call Comando_Fechar(lComando)
    
    Configura_DesconsideraFechtoPeriodo = SUCESSO
    
    Exit Function
    
Erro_Configura_DesconsideraFechtoPeriodo:

    Configura_DesconsideraFechtoPeriodo = gErr

    Select Case gErr

        Case 201294, 201295, 201296
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFIG_FECHTO_PERIODO_CTB", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201297)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
