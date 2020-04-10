VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSATInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sCodigoDeAtivacao As String
Public iEmuladorSefaz As Integer
Public sDirArqXml As String
Public sNomeArqDLL As String
Public sNomeArqEmulador As String
Public iModeloImpressora As Integer
Public iLayoutImpressao As Integer
Public iEmTeste As Integer
Public sNomeArqLogo As String
Public sPortaImpressora As String
Public ssignAC As String

'dados do emitente
Public sEmitCNPJ As String
Public sEmitIE As String
Public sEmitIM As String
Public sEmitRazaoSocial As String
Public sEmitNomeReduzido As String
Public sEmitEndereco As String
Public iRegTribISSQN As Integer '1 - Microempresa Municipal; 2 - Estimativa; 3 - Sociedade de Profissionais; 4 - Cooperativa; 5 - Microempresário Individual (MEI);
Public sindRatISSQN As String 'Informa se o Desconto sobre subtotal deve ser rateado entre os itens sujeitos à tributação pelo ISSQN. 'S' - Desconto sobre subtotal será rateado entre os itens sujeitos ao ISSQN. 'N' - Desconto sobre subtotal não será rateado entre os itens sujeitos ao ISSQN.

Public Sub Copia(ByVal objSATConfig As ClassConfiguracaoSAT)

    With objSATConfig
        iEmuladorSefaz = .iEmuladorSefaz
        iEmTeste = .iEmTeste
        sCodigoDeAtivacao = .sCodigoDeAtivacao
        sDirArqXml = .sDirArqXml
        sNomeArqDLL = .sNomeArqDLL
        sNomeArqEmulador = .sNomeArqEmulador
        iModeloImpressora = .iModeloImpressora
        iLayoutImpressao = .iLayoutImpressao
        sNomeArqLogo = .sNomeArqLogo
        sPortaImpressora = .sPortaImpressora
        ssignAC = .ssignAC
    End With

End Sub

Public Function Inicializa() As Long
Dim X As New ClassECFConfig, objSATConfig As New ClassConfiguracaoSAT, lErro As Long

On Error GoTo Erro_Inicializa

    'copia informacoes de ECFConfig
    sEmitCNPJ = X.gsCNPJ
    sEmitIE = X.gsInscricaoEstadual
    sEmitIM = X.gsInscricaoMunicipal
    sEmitRazaoSocial = gsNomeEmpresa
    sEmitEndereco = X.gsEndereco
    
    '???
    'iRegTribISSQN
    'sindRatISSQN
    
    lErro = CF_ECF("ConfiguracaoSAT_Le", objSATConfig)
    If lErro <> SUCESSO Then gError 214014
    
    Call Copia(objSATConfig)
    
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


