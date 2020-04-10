VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodFiscalServico As Integer
Private mvariLivFiscSemCtaCtb As Integer
Private mvariRegravaNF As Integer
Private mvardtFisBloqDataLimite As Date
Private mvariIPIOutrasIncluiIPI As Integer 'se o próprio ipi nao creditado deve ser somado na coluna "outras"...ex...compra de 1000 com 100 de ipi nao creditado...a coluna outras deve ser 1100...
Private mvariRegInvPegaCtaProd As Integer

Public Property Let iRegInvPegaCtaProd(ByVal vData As Integer)
    mvariRegInvPegaCtaProd = vData
End Property

Public Property Get iRegInvPegaCtaProd() As Integer
    iRegInvPegaCtaProd = mvariRegInvPegaCtaProd
End Property

Public Property Let iIPIOutrasIncluiIPI(ByVal vData As Integer)
    mvariIPIOutrasIncluiIPI = vData
End Property

Public Property Get iIPIOutrasIncluiIPI() As Integer
    iIPIOutrasIncluiIPI = mvariIPIOutrasIncluiIPI
End Property

Public Property Let dtFisBloqDataLimite(ByVal vData As Date)
    mvardtFisBloqDataLimite = vData
End Property

Public Property Get dtFisBloqDataLimite() As Date
    dtFisBloqDataLimite = mvardtFisBloqDataLimite
End Property

Public Property Let iRegravaNF(ByVal vData As Integer)
    mvariRegravaNF = vData
End Property

Public Property Get iRegravaNF() As Integer
    iRegravaNF = mvariRegravaNF
End Property

Public Property Let iLivFiscSemCtaCtb(ByVal vData As Integer)
    mvariLivFiscSemCtaCtb = vData
End Property


Public Property Get iLivFiscSemCtaCtb() As Integer
    iLivFiscSemCtaCtb = mvariLivFiscSemCtaCtb
End Property

Public Property Let iCodFiscalServico(ByVal vData As Integer)
    mvariCodFiscalServico = vData
End Property

Public Property Get iCodFiscalServico() As Integer
    iCodFiscalServico = mvariCodFiscalServico
End Property

Public Function Inicializa() As Long

Dim objFis As New ClassFIS
Dim lErro As Long
Dim objFISConfig As New ClassFISConfig

On Error GoTo Erro_Inicializa

    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_CODIGO_FISCAL_SERVICOS"

    'Lê a configurações em FisConfig
    lErro = CF("FISConfig_Le2", objFISConfig)
    If lErro <> SUCESSO Then gError 80153

    mvariCodFiscalServico = StrParaInt(objFISConfig.sConteudo)
        
    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_LIVFISC_SEM_CTA_CTB"

    'Lê a configurações em FisConfig
    lErro = CF("FISConfig_Le2", objFISConfig)
    If lErro <> SUCESSO Then gError 80153

    mvariLivFiscSemCtaCtb = StrParaInt(objFISConfig.sConteudo)
    
    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_REGRAVA_NF"

    'Lê a configurações em FisConfig
    lErro = CF("FISConfig_Le2", objFISConfig)
    If lErro <> SUCESSO Then gError 80153

    mvariRegravaNF = StrParaInt(objFISConfig.sConteudo)
    
    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_BLOQUEIO_DATA_LIMITE"

    'Lê a configurações em FisConfig
    lErro = CF("FISConfig_Le2", objFISConfig)
    If lErro <> SUCESSO Then gError 80153

    mvardtFisBloqDataLimite = StrParaDate(objFISConfig.sConteudo)

    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_IPIOUTRAS_INCLUI_IPI"

    'Lê a configurações em FisConfig
    lErro = CF("FISConfig_Le2", objFISConfig)
    If lErro <> SUCESSO Then gError 80153

    mvariIPIOutrasIncluiIPI = StrParaInt(objFISConfig.sConteudo)

    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "REG_INV_PEGA_CTA_PROD"

    'Lê a configurações em FisConfig
    lErro = CF("FISConfig_Le2", objFISConfig)
    If lErro <> SUCESSO Then gError 80153

    mvariRegInvPegaCtaProd = StrParaInt(objFISConfig.sConteudo)
    
    Inicializa = SUCESSO

    Exit Function

Erro_Inicializa:

    Inicializa = gErr

    Select Case gErr

        Case 80153
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150291)

    End Select

    Exit Function

End Function

