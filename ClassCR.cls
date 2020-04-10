VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAglutinaLancamPorDia As Integer
Private mvariGeraLoteAutomatico As Integer
Private mvariContabSemDet As Integer 'gera contabilizacao sem detalhar a parcela, na baixa cp ou cr
Private mvariFilialCentralizadora As Integer
Private mvariIRRF As Integer
Private mvariVerificaChqMesmaConta As Integer

Public Function Inicializa() As Long

Dim lErro As Long, colConfigCR As New ColCRConfig

On Error GoTo Erro_Inicializa

    colConfigCR.Add CRCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", CRCFG_AGLUTINA_LANCAM_POR_DIA
    colConfigCR.Add CRCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", CRCFG_GERA_LOTE_AUTOMATICO
    colConfigCR.Add CRCFG_CR_CONTAB_SEM_DET, EMPRESA_TODA, "", 0, "", CRCFG_CR_CONTAB_SEM_DET
    colConfigCR.Add CRCFG_FILIAL_CENTRALIZADORA, EMPRESA_TODA, "", 0, "", CRCFG_FILIAL_CENTRALIZADORA
    colConfigCR.Add CRCFG_CR_IRRF_LANCAM_BAIXA, EMPRESA_TODA, "", 0, "", CRCFG_CR_IRRF_LANCAM_BAIXA
    colConfigCR.Add CRCFG_VERIFICA_CHEQUES_MESMA_CONTA, EMPRESA_TODA, "", 0, "", CRCFG_VERIFICA_CHEQUES_MESMA_CONTA
    
    'Lê as configurações em CRConfig
    lErro = CF("CRConfig_Le_Configs", colConfigCR)
    If lErro <> SUCESSO Then Error 33689
    
    mvariAglutinaLancamPorDia = CInt(colConfigCR.Item(CRCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
    mvariGeraLoteAutomatico = CInt(colConfigCR.Item(CRCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    mvariContabSemDet = CInt(colConfigCR.Item(CRCFG_CR_CONTAB_SEM_DET).sConteudo)
    mvariFilialCentralizadora = CInt(colConfigCR.Item(CRCFG_FILIAL_CENTRALIZADORA).sConteudo)
    mvariIRRF = CInt(colConfigCR.Item(CRCFG_CR_IRRF_LANCAM_BAIXA).sConteudo)
    mvariVerificaChqMesmaConta = CInt(colConfigCR.Item(CRCFG_VERIFICA_CHEQUES_MESMA_CONTA).sConteudo)
    
    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 33689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CRCONFIG", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148810)
        
    End Select

    Exit Function

End Function

Public Property Get iVerificaChqMesmaConta() As Integer
    iVerificaChqMesmaConta = mvariVerificaChqMesmaConta
End Property

Public Property Let iVerificaChqMesmaConta(ByVal vData As Integer)
    mvariVerificaChqMesmaConta = vData
End Property

Public Property Get iIRRF() As Integer
    iIRRF = mvariIRRF
End Property

Public Property Let iIRRF(ByVal vData As Integer)
    mvariIRRF = vData
End Property

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

Dim lErro As Long, colConfigCR As New ColCRConfig

On Error GoTo Erro_Gravar
    
    colConfigCR.Add CRCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfigCR.Add CRCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfigCR.Add CRCFG_CR_CONTAB_SEM_DET, EMPRESA_TODA, "", 0, CStr(mvariContabSemDet)
    colConfigCR.Add CRCFG_FILIAL_CENTRALIZADORA, EMPRESA_TODA, "", 0, CStr(mvariFilialCentralizadora)
    
    'Chama CRConfig_Grava_Configs
    lErro = CF("CRConfig_Grava_Configs", colConfigCR)
    If lErro <> SUCESSO Then Error 33690
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = Err
    
    Select Case Err
    
        Case 33690
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148811)
    
    End Select
        
    Exit Function
    
End Function

Function Gravar_Trans() As Long
'Grava no bd a configuração atual

Dim lErro As Long, colConfig As New ColCRConfig

On Error GoTo Erro_Gravar_Trans
    
    colConfig.Add CRCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add CRCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add CRCFG_CR_CONTAB_SEM_DET, EMPRESA_TODA, "", 0, CStr(mvariContabSemDet)
    colConfig.Add CRCFG_FILIAL_CENTRALIZADORA, EMPRESA_TODA, "", 0, CStr(mvariFilialCentralizadora)
    
    'Chama CRConfig_Grava_Configs
    lErro = CF("CRConfig_Grava_Configs_Trans", colConfig)
    If lErro <> SUCESSO Then Error 41814
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = Err
    
    Select Case Err
    
        Case 41814
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148812)
    
    End Select
        
    Exit Function
    
End Function
