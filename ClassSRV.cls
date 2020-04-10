VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAglutinaLancamDia As Integer
Private mvariGeraLoteAutomatico As Integer
Private mvariValidaGarantia As Integer
Private mvariValidaManutencao As Integer
Private mvariGarantiaAutoSolic As Integer
Private mvariContratoAutoSolic As Integer
Private mvariVerificaLote As Integer

Public Property Get iVerificaLote() As Integer
    iVerificaLote = mvariVerificaLote
End Property

Public Property Let iVerificaLote(ByVal vData As Integer)
    mvariVerificaLote = vData
End Property

Public Property Get iContratoAutoSolic() As Integer
    iContratoAutoSolic = mvariContratoAutoSolic
End Property

Public Property Let iContratoAutoSolic(ByVal vData As Integer)
    mvariContratoAutoSolic = vData
End Property

Public Property Get iGarantiaAutoSolic() As Integer
    iGarantiaAutoSolic = mvariGarantiaAutoSolic
End Property

Public Property Let iGarantiaAutoSolic(ByVal vData As Integer)
    mvariGarantiaAutoSolic = vData
End Property

Public Property Get iAglutinaLancamDia() As Integer
    iAglutinaLancamDia = mvariAglutinaLancamDia
End Property

Public Property Let iAglutinaLancamDia(ByVal vData As Integer)
    mvariAglutinaLancamDia = vData
End Property

Public Property Get iGeraLoteAutomatico() As Integer
    iGeraLoteAutomatico = mvariGeraLoteAutomatico
End Property

Public Property Let iGeraLoteAutomatico(ByVal vData As Integer)
    mvariGeraLoteAutomatico = vData
End Property

Public Property Get iValidaGarantia() As Integer
    iValidaGarantia = mvariValidaGarantia
End Property

Public Property Let iValidaGarantia(ByVal vData As Integer)
    mvariValidaGarantia = vData
End Property

Public Property Get iValidaManutencao() As Integer
    iValidaManutencao = mvariValidaManutencao
End Property

Public Property Let iValidaManutencao(ByVal vData As Integer)
    mvariValidaManutencao = vData
End Property

Public Function Inicializa() As Long

Dim lErro As Long
Dim colConfigSRV As New Collection
Dim objSRVConfig As ClassSRVConfig

On Error GoTo Erro_Inicializa

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_AGLUTINA_LANCAM_POR_DIA
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_GERA_LOTE_AUTOMATICO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_VALIDA_GARANTIA
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_VALIDA_MANUTENCAO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_GARANTIA_AUTOMATICA_SOLICITACAO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_CONTRATO_AUTOMATICO_SOLICITACAO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_VERIFICA_LOTE
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    
    colConfigSRV.Add objSRVConfig

    'Lê as configurações em CRFATConfig
    lErro = CF("SRVConfig_Le_Configs", colConfigSRV)
    If lErro <> SUCESSO Then gError 183384
    
    For Each objSRVConfig In colConfigSRV
    
        Select Case objSRVConfig.sCodigo
        
            Case SRVCFG_AGLUTINA_LANCAM_POR_DIA
                mvariAglutinaLancamDia = StrParaInt(objSRVConfig.sConteudo)
            
            Case SRVCFG_GERA_LOTE_AUTOMATICO
                mvariGeraLoteAutomatico = StrParaInt(objSRVConfig.sConteudo)
    
            Case SRVCFG_VALIDA_GARANTIA
                mvariValidaGarantia = StrParaInt(objSRVConfig.sConteudo)
    
            Case SRVCFG_VALIDA_MANUTENCAO
                mvariValidaManutencao = StrParaInt(objSRVConfig.sConteudo)
    
            Case SRVCFG_GARANTIA_AUTOMATICA_SOLICITACAO
                mvariGarantiaAutoSolic = StrParaInt(objSRVConfig.sConteudo)
    
            Case SRVCFG_CONTRATO_AUTOMATICO_SOLICITACAO
                mvariContratoAutoSolic = StrParaInt(objSRVConfig.sConteudo)
    
            Case SRVCFG_VERIFICA_LOTE
                mvariVerificaLote = StrParaInt(objSRVConfig.sConteudo)
    
        End Select
        
    Next
    

    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = gErr
    
    Select Case gErr
    
        Case 183384
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 183385)
        
    End Select

    Exit Function

End Function

Function Gravar() As Long
'Grava no bd a configuração atual

Dim lErro As Long
Dim colConfigSRV As New Collection
Dim objSRVConfig As ClassSRVConfig

On Error GoTo Erro_Gravar
    
    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_AGLUTINA_LANCAM_POR_DIA
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariAglutinaLancamDia)
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_GERA_LOTE_AUTOMATICO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariGeraLoteAutomatico)
    
    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_VALIDA_GARANTIA
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariValidaGarantia)

    colConfigSRV.Add objSRVConfig

    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_VALIDA_MANUTENCAO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariValidaManutencao)
    
    colConfigSRV.Add objSRVConfig
        
    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_GARANTIA_AUTOMATICA_SOLICITACAO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariGarantiaAutoSolic)
    
    colConfigSRV.Add objSRVConfig
        
    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_CONTRATO_AUTOMATICO_SOLICITACAO
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariContratoAutoSolic)
    
    colConfigSRV.Add objSRVConfig
        
    Set objSRVConfig = New ClassSRVConfig
    
    objSRVConfig.sCodigo = SRVCFG_VERIFICA_LOTE
    objSRVConfig.iFilialEmpresa = EMPRESA_TODA
    objSRVConfig.sConteudo = CStr(mvariVerificaLote)

    colConfigSRV.Add objSRVConfig
        
    lErro = CF("SRVConfig_Grava_Configs", colConfigSRV)
    If lErro <> SUCESSO Then gError 183386
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = gErr
    
    Select Case gErr
    
        Case 183386
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183387)
    
    End Select
        
    Exit Function
    
End Function

