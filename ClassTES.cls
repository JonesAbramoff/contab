VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAglutinaLancamPorDia As Integer
Private mvariGeraLoteAutomatico As Integer
Private mvariFluxoIncluiPedidos As Integer
Private mvariMovCCVerificaBloqData As Integer
Private mvardtMovCCBloqDataLimite As Date

Public Function Inicializa() As Long

Dim lErro As Long, colConfig As New ColTESConfig
Dim sConteudo As String

On Error GoTo Erro_Inicializa

    colConfig.Add TESCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", TESCFG_AGLUTINA_LANCAM_POR_DIA
    colConfig.Add TESCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", TESCFG_GERA_LOTE_AUTOMATICO
    colConfig.Add TESCFG_FLUXO_INCLUI_PEDIDOS, EMPRESA_TODA, "", 0, "", TESCFG_FLUXO_INCLUI_PEDIDOS
    colConfig.Add TESCFG_MOVCC_VERIFICA_BLOQUEIO_DATA, EMPRESA_TODA, "", 0, "", TESCFG_MOVCC_VERIFICA_BLOQUEIO_DATA
    colConfig.Add TESCFG_MOVCC_BLOQUEIO_DATA_LIMITE, EMPRESA_TODA, "", 0, "", TESCFG_MOVCC_BLOQUEIO_DATA_LIMITE
    
    'Lê as configurações em TESConfig
    lErro = CF("TESConfig_Le_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33709
    
    mvariAglutinaLancamPorDia = CInt(colConfig.Item(TESCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
    mvariGeraLoteAutomatico = CInt(colConfig.Item(TESCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    mvariFluxoIncluiPedidos = CInt(colConfig.Item(TESCFG_FLUXO_INCLUI_PEDIDOS).sConteudo)
    
    mvariMovCCVerificaBloqData = StrParaInt(colConfig.Item(TESCFG_MOVCC_VERIFICA_BLOQUEIO_DATA).sConteudo)
    sConteudo = colConfig.Item(TESCFG_MOVCC_BLOQUEIO_DATA_LIMITE).sConteudo
    If Len(Trim(sConteudo)) > 0 Then
        mvardtMovCCBloqDataLimite = StrParaDate(sConteudo)
    Else
        mvardtMovCCBloqDataLimite = DATA_NULA
    End If
    
    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 33709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESCONFIG", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153605)
        
    End Select

    Exit Function

End Function

Public Property Get iAglutinaLancamPorDia() As Integer
    iAglutinaLancamPorDia = mvariAglutinaLancamPorDia
End Property

Public Property Get iGeraLoteAutomatico() As Integer
    iGeraLoteAutomatico = mvariGeraLoteAutomatico
End Property

Public Property Get iFluxoIncluiPedidos() As Integer
    iFluxoIncluiPedidos = mvariFluxoIncluiPedidos
End Property

Public Property Let iAglutinaLancamPorDia(ByVal vData As Integer)
    mvariAglutinaLancamPorDia = vData
End Property

Public Property Let iGeraLoteAutomatico(ByVal vData As Integer)
    mvariGeraLoteAutomatico = vData
End Property

Public Property Let iFluxoIncluiPedidos(ByVal vData As Integer)
    mvariFluxoIncluiPedidos = vData
End Property

Public Property Let iMovCCVerificaBloqData(ByVal vData As Integer)
    mvariMovCCVerificaBloqData = vData
End Property

Public Property Get iMovCCVerificaBloqData() As Integer
    iMovCCVerificaBloqData = mvariMovCCVerificaBloqData
End Property

Public Property Let dtMovCCBloqDataLimite(ByVal vData As Date)
    mvardtMovCCBloqDataLimite = vData
End Property

Public Property Get dtMovCCBloqDataLimite() As Date
    dtMovCCBloqDataLimite = mvardtMovCCBloqDataLimite
End Property

Function Gravar() As Long
'Grava no bd a configuração atual

Dim lErro As Long, colConfig As New ColTESConfig

On Error GoTo Erro_Gravar
    
    colConfig.Add TESCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add TESCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add TESCFG_FLUXO_INCLUI_PEDIDOS, EMPRESA_TODA, "", 0, CStr(mvariFluxoIncluiPedidos)
    colConfig.Add TESCFG_MOVCC_VERIFICA_BLOQUEIO_DATA, EMPRESA_TODA, "", 0, CStr(mvariMovCCVerificaBloqData)
    colConfig.Add TESCFG_MOVCC_BLOQUEIO_DATA_LIMITE, EMPRESA_TODA, "", 0, CStr(mvardtMovCCBloqDataLimite)
    
    'Chama TESConfig_Grava_Configs
    lErro = CF("TESConfig_Grava_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33710
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = Err
    
    Select Case Err
    
        Case 33710
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153606)
    
    End Select
        
    Exit Function
    
End Function

Function Gravar_Trans() As Long
'Grava no bd a configuração atual

Dim lErro As Long, colConfig As New ColTESConfig

On Error GoTo Erro_Gravar_Trans
    
    colConfig.Add TESCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add TESCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add TESCFG_FLUXO_INCLUI_PEDIDOS, EMPRESA_TODA, "", 0, CStr(mvariFluxoIncluiPedidos)
    colConfig.Add TESCFG_MOVCC_VERIFICA_BLOQUEIO_DATA, EMPRESA_TODA, "", 0, CStr(mvariMovCCVerificaBloqData)
    colConfig.Add TESCFG_MOVCC_BLOQUEIO_DATA_LIMITE, EMPRESA_TODA, "", 0, CStr(mvardtMovCCBloqDataLimite)
    
    'Chama TESConfig_Grava_Configs
    lErro = CF("TESConfig_Grava_Configs_Trans", colConfig)
    If lErro <> SUCESSO Then Error 44694
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = Err
    
    Select Case Err
    
        Case 44694
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153607)
    
    End Select
        
    Exit Function
    
End Function



