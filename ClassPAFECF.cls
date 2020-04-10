VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPAFECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCNPJDesenv As String
Private mvarsInscricaoMunicipalDesenv As String
Private mvarsInscricaoEstadualDesenv As String
Private mvarsNomeDesenv As String
Private mvarsNomePAFECF As String
Private mvarsVersaoPAFECF As String
Private mvarsMD5PAFECF As String
Private mvarsEnderecoDesenv As String
Private mvarsContatoDesenv As String
Private mvarsTelefoneDesenv As String
Private mvarsNomePrincArqExec As String
Private mvarsMD5PrincArqExec As String
Private mvariCodModeloECF As Integer
Private mvardtDataLog As Date
Private mvariAtivo As Integer
Private mvariImpressoraECF As Integer
Private mvarsNome_Fabricante As String
Private mvarsNome_Caixa As String
Private mvariAtualizacaoECF As Integer
Private mvariHorarioVerao As Integer
Private mvariTipoTEF As Integer
Private mvarsPOS As String
Private mvariLeitoraCodigoBarras As Integer



Public Property Get sCNPJDesenv() As String
    sCNPJDesenv = mvarsCNPJDesenv
End Property

Public Property Let sCNPJDesenv(ByVal vData As String)
    mvarsCNPJDesenv = vData
End Property

Public Property Get sInscricaoMunicipalDesenv() As String
    sInscricaoMunicipalDesenv = mvarsInscricaoMunicipalDesenv
End Property

Public Property Let sInscricaoMunicipalDesenv(ByVal vData As String)
    mvarsInscricaoMunicipalDesenv = vData
End Property

Public Property Get sInscricaoEstadualDesenv() As String
    sInscricaoEstadualDesenv = mvarsInscricaoEstadualDesenv
End Property

Public Property Let sInscricaoEstadualDesenv(ByVal vData As String)
    mvarsInscricaoEstadualDesenv = vData
End Property

Public Property Get sNomeDesenv() As String
    sNomeDesenv = mvarsNomeDesenv
End Property

Public Property Let sNomeDesenv(ByVal vData As String)
    mvarsNomeDesenv = vData
End Property

Public Property Get sNomePAFECF() As String
    sNomePAFECF = mvarsNomePAFECF
End Property

Public Property Let sNomePAFECF(ByVal vData As String)
    mvarsNomePAFECF = vData
End Property

Public Property Get sVersaoPAFECF() As String
    sVersaoPAFECF = mvarsVersaoPAFECF
End Property

Public Property Let sVersaoPAFECF(ByVal vData As String)
    mvarsVersaoPAFECF = vData
End Property

Public Property Get sMD5PAFECF() As String
    sMD5PAFECF = mvarsMD5PAFECF
End Property

Public Property Let sMD5PAFECF(ByVal vData As String)
    mvarsMD5PAFECF = vData
End Property

Public Property Get sEnderecoDesenv() As String
    sEnderecoDesenv = mvarsEnderecoDesenv
End Property

Public Property Let sEnderecoDesenv(ByVal vData As String)
    mvarsEnderecoDesenv = vData
End Property

Public Property Get sContatoDesenv() As String
    sContatoDesenv = mvarsContatoDesenv
End Property

Public Property Let sContatoDesenv(ByVal vData As String)
    mvarsContatoDesenv = vData
End Property

Public Property Get sTelefoneDesenv() As String
    sTelefoneDesenv = mvarsTelefoneDesenv
End Property

Public Property Let sTelefoneDesenv(ByVal vData As String)
    mvarsTelefoneDesenv = vData
End Property

Public Property Get sNomePrincArqExec() As String
    sNomePrincArqExec = mvarsNomePrincArqExec
End Property

Public Property Let sNomePrincArqExec(ByVal vData As String)
    mvarsNomePrincArqExec = vData
End Property

Public Property Get sMD5PrincArqExec() As String
    sMD5PrincArqExec = mvarsMD5PrincArqExec
End Property

Public Property Let sMD5PrincArqExec(ByVal vData As String)
    mvarsMD5PrincArqExec = vData
End Property

Public Property Get gsVersaoEspecPAFECF() As String
    gsVersaoEspecPAFECF = SistGlobECF_sVersaoEspecPAFECF
End Property

Public Property Let gsVersaoEspecPAFECF(ByVal vData As String)
    SistGlobECF_sVersaoEspecPAFECF = vData
End Property

Public Property Get gsLaudo() As String
    gsLaudo = SistGlobECF_sLaudo
End Property

Public Property Let gsLaudo(ByVal vData As String)
    SistGlobECF_sLaudo = vData
End Property

Public Property Let iCodModeloECF(ByVal vData As Integer)
    mvariCodModeloECF = vData
End Property

Public Property Get iCodModeloECF() As Integer
    iCodModeloECF = mvariCodModeloECF
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iImpressoraECF(ByVal vData As Integer)
    mvariImpressoraECF = vData
End Property

Public Property Get iImpressoraECF() As Integer
    iImpressoraECF = mvariImpressoraECF
End Property

Public Property Let sNome_Fabricante(ByVal vData As String)
    mvarsNome_Fabricante = vData
End Property

Public Property Get sNome_Fabricante() As String
    sNome_Fabricante = mvarsNome_Fabricante
End Property

Public Property Let sNome_Caixa(ByVal vData As String)
    mvarsNome_Caixa = vData
End Property

Public Property Get sNome_Caixa() As String
    sNome_Caixa = mvarsNome_Caixa
End Property

Public Property Let iAtualizacaoECF(ByVal vData As Integer)
    mvariAtualizacaoECF = vData
End Property

Public Property Get iAtualizacaoECF() As Integer
    iAtualizacaoECF = mvariAtualizacaoECF
End Property

Public Property Let iHorarioVerao(ByVal vData As Integer)
    mvariHorarioVerao = vData
End Property

Public Property Get iHorarioVerao() As Integer
    iHorarioVerao = mvariHorarioVerao
End Property

Public Property Let iTipoTEF(ByVal vData As Integer)
    mvariTipoTEF = vData
End Property

Public Property Get iTipoTEF() As Integer
    iTipoTEF = mvariTipoTEF
End Property

Public Property Let sPOS(ByVal vData As String)
    mvarsPOS = vData
End Property

Public Property Get sPOS() As String
    sPOS = mvarsPOS
End Property

Public Property Let iLeitoraCodigoBarras(ByVal vData As Integer)
    mvariLeitoraCodigoBarras = vData
End Property

Public Property Get iLeitoraCodigoBarras() As Integer
    iLeitoraCodigoBarras = mvariLeitoraCodigoBarras
End Property

