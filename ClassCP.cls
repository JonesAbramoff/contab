VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAglutinaLancamPorDia As Integer
Private mvariGeraLoteAutomatico As Integer
Private mvariContabSemDet As Integer 'gera contabilizacao sem detalhar a parcela, na baixa cp ou cr
Private mvariFilialCentralizadora As Integer

Public Function Inicializa() As Long

Dim lErro As Long, colConfig As New ColCPConfig

On Error GoTo Erro_Inicializa

    colConfig.Add CPCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", CPCFG_AGLUTINA_LANCAM_POR_DIA
    colConfig.Add CPCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", CPCFG_GERA_LOTE_AUTOMATICO
    colConfig.Add CPCFG_CPR_CONTAB_SEM_DET, EMPRESA_TODA, "", 0, "", CPCFG_CPR_CONTAB_SEM_DET
    colConfig.Add CPCFG_FILIAL_CENTRALIZADORA, EMPRESA_TODA, "", 0, "", CPCFG_FILIAL_CENTRALIZADORA
    
    'Lê as configurações em CPConfig
    lErro = CF("CPConfig_Le_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33669
    
    mvariAglutinaLancamPorDia = CInt(colConfig.Item(CPCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
    mvariGeraLoteAutomatico = CInt(colConfig.Item(CPCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    mvariContabSemDet = CInt(colConfig.Item(CPCFG_CPR_CONTAB_SEM_DET).sConteudo)
    mvariFilialCentralizadora = CInt(colConfig.Item(CPCFG_FILIAL_CENTRALIZADORA).sConteudo)
    
    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 33669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPCONFIG", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147883)
        
    End Select

    Exit Function

End Function

Public Property Get iAglutinaLancamPorDia() As Integer
    iAglutinaLancamPorDia = mvariAglutinaLancamPorDia
End Property

Public Property Let iAglutinaLancamPorDia(ByVal vData As Integer)
    mvariAglutinaLancamPorDia = vData
End Property

Public Property Get iContabSemDet() As Integer
    iContabSemDet = mvariContabSemDet
End Property

Public Property Let iContabSemDet(ByVal vData As Integer)
    mvariContabSemDet = vData
End Property

Public Property Get iGeraLoteAutomatico() As Integer
    iGeraLoteAutomatico = mvariGeraLoteAutomatico
End Property

Public Property Let iGeraLoteAutomatico(ByVal vData As Integer)
    mvariGeraLoteAutomatico = vData
End Property

Public Property Get iFilialCentralizadora() As Integer
    iFilialCentralizadora = mvariFilialCentralizadora
End Property

Public Property Let iFilialCentralizadora(ByVal vData As Integer)
    mvariFilialCentralizadora = vData
End Property

Function Gravar() As Long
'Grava no bd a configuração atual

Dim lErro As Long, colConfig As New ColCPConfig

On Error GoTo Erro_Gravar
    
    colConfig.Add CPCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add CPCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add CPCFG_CPR_CONTAB_SEM_DET, EMPRESA_TODA, "", 0, CStr(mvariContabSemDet)
    colConfig.Add CPCFG_FILIAL_CENTRALIZADORA, EMPRESA_TODA, "", 0, CStr(mvariFilialCentralizadora)
    
    'Chama CPConfig_Grava_Configs
    lErro = CF("CPConfig_Grava_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33670
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = Err
    
    Select Case Err
    
        Case 33670
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147884)
    
    End Select
        
    Exit Function
    
End Function

Function Gravar_Trans() As Long
'Grava no bd a configuração atual

Dim lErro As Long, colConfig As New ColCPConfig

On Error GoTo Erro_Gravar_Trans
    
    colConfig.Add CPCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add CPCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add CPCFG_CPR_CONTAB_SEM_DET, EMPRESA_TODA, "", 0, CStr(mvariContabSemDet)
    colConfig.Add CPCFG_FILIAL_CENTRALIZADORA, EMPRESA_TODA, "", 0, CStr(mvariFilialCentralizadora)
    
    'Chama CPConfig_Grava_Configs
    lErro = CF("CPConfig_Grava_Configs_Trans", colConfig)
    If lErro <> SUCESSO Then Error 41813
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = Err
    
    Select Case Err
    
        Case 41813
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147885)
    
    End Select
        
    Exit Function
    
End Function


