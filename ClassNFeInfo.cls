VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sToken As String
Public sNFCeSerie As String
Public lNFCeProximoNum As Long
Public lNFCeProximoLote As Long
Public sDirArqXml As String
Public sCertificadoA1A3 As String
Public iNFeAmbiente As Integer
Public sNFeSerie As String
Public lNFeProximoNum As Long
Public lNFeProximoLote As Long
Public sDirXsd As String
Public iNFCeAmbiente As Integer
Public sNFCECSC As String
Public sidNFCECSC As String
Public iEmContingencia As Integer
Public dtContingenciaDataEntrada As Date
Public dContingenciaHoraEntrada As Double
Public sContigenciaxJust As String
Public iModeloImpressora As Integer
Public sPortaImpressora As String

Public iNFDescricaoProd As Integer '0:usar o nome reduzido do produto, 1: usar a descricao

Public sSMTP As String
Public sSMTPUsu As String
Public sSMTPSenha As String
Public lSMTPPorta As Long
Public iNFCeImprimir As Integer
Public iNFCeEnviarEmail As Integer

'dados do emitente
Public sEmitCNPJ As String
Public sEmitIE As String
Public sEmitIM As String
Public sEmitRazaoSocial As String
Public sEmitNomeReduzido As String
Public sEmitEndereco As String
Public iRegTribISSQN As Integer '1 - Microempresa Municipal; 2 - Estimativa; 3 - Sociedade de Profissionais; 4 - Cooperativa; 5 - Microempresário Individual (MEI);
Public sindRatISSQN As String 'Informa se o Desconto sobre subtotal deve ser rateado entre os itens sujeitos à tributação pelo ISSQN. 'S' - Desconto sobre subtotal será rateado entre os itens sujeitos ao ISSQN. 'N' - Desconto sobre subtotal não será rateado entre os itens sujeitos ao ISSQN.
Public sEmitUF As String
Public iUFCodIBGE As Integer
Public cMun As Long 'codigo do municipio pelo ibge
Public sEmitCNAE As String
Public sEmitCEP As String
Public iVersaoNFe As Integer

Public objEndereco As New ClassEndereco

Public sMD5PAFECF As String
Public iFocaTipoVenda As Integer

Public Function PGM_SGENFE() As String
    Select Case iVersaoNFe
        Case 310
            PGM_SGENFE = "sgenfe"
        Case 400
            PGM_SGENFE = "sgenfe4"
    End Select
End Function

Public Sub CopiaConfig(ByVal objNFeConfig As ClassConfiguracaoNFe)

    With objNFeConfig
        sToken = .sToken
        sNFCeSerie = .sNFCeSerie
        lNFCeProximoNum = .lNFCeProximoNum
        lNFCeProximoLote = .lNFCeProximoLote
        sDirArqXml = .sDirArqXml
        sCertificadoA1A3 = .sCertificadoA1A3
        iNFeAmbiente = .iNFeAmbiente
        sNFeSerie = .sNFeSerie
        lNFeProximoNum = .lNFeProximoNum
        lNFeProximoLote = .lNFeProximoLote
        sDirXsd = .sDirXsd
        iNFCeAmbiente = .iNFCeAmbiente
        sNFCECSC = .sNFCECSC
        sidNFCECSC = .sidNFCECSC
        iEmContingencia = .iEmContingencia
        dtContingenciaDataEntrada = .dtContingenciaDataEntrada
        dContingenciaHoraEntrada = .dContingenciaHoraEntrada
        sContigenciaxJust = .sContigenciaxJust
        iModeloImpressora = .iModeloImpressora
        sPortaImpressora = .sPortaImpressora
        sSMTP = .sSMTP
        sSMTPUsu = .sSMTPUsu
        sSMTPSenha = .sSMTPSenha
         lSMTPPorta = .lSMTPPorta
        iNFCeImprimir = .iNFCeImprimir
        iNFCeEnviarEmail = .iNFCeEnviarEmail
        iNFDescricaoProd = .iNFDescricaoProd
        iVersaoNFe = .iVersaoNFe
        iFocaTipoVenda = .iFocaTipoVenda
        
    End With

End Sub

Public Function Inicializa() As Long
Dim X As New ClassECFConfig, objNFeConfig As New ClassConfiguracaoNFe, lErro As Long

On Error GoTo Erro_Inicializa

    'copia informacoes de ECFConfig
    sEmitCNPJ = X.gsCNPJ
    sEmitIE = X.gsInscricaoEstadual
    sEmitIM = X.gsInscricaoMunicipal
    sEmitRazaoSocial = gsNomeEmpresa
    sEmitEndereco = X.gsEndereco
    sEmitNomeReduzido = X.gsNomeReduzido
    sEmitUF = X.gsUF
    sEmitCEP = X.gsCEP
    sEmitCNAE = X.gsCNAE
    
    With objEndereco
        .sLogradouro = X.gsEndLogradouro
        .lNumero = StrParaLong(X.gsEndNumero)
        .sComplemento = X.gsEndComplemento
        .sBairro = X.gsBairro
        .sCidade = X.gsCidade
        .sSiglaEstado = X.gsUF
        .sCEP = sEmitCEP
    End With
    
    '???
    'iRegTribISSQN
    'sindRatISSQN
    
    lErro = CF_ECF("ConfiguracaoNFe_Le", objNFeConfig)
    If lErro <> SUCESSO Then gError 214014
    
    Call CopiaConfig(objNFeConfig)
    
    sMD5PAFECF = X.gsMD5PAFECF
    
    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = gErr

    Select Case gErr
    
        Case 214014
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 214015)

    End Select
    
    Exit Function
    
End Function



