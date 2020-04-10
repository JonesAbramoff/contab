VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodFamilia As Long
Private mvarsSobrenome As String
Private mvarsTitularNome As String
Private mvarsTitularNomeHebr As String
Private mvarlTitularEnderecoRes As Long
Private mvarsTitularNomeFirma As String
Private mvarlTitularEnderecoCom As Long
Private mvariLocalCobranca As Integer
Private mvariEstadoCivil As Integer
Private mvarsTitularProfissao As String
Private mvardtTitularDtNasc As Date
Private mvariTitularDtNascNoite As Integer
Private mvardtDataCasamento As Date
Private mvariDataCasamentoNoite As Integer
Private mvarsCohenLeviIsrael As String
Private mvarsTitularPai As String
Private mvarsTitularPaiHebr As String
Private mvarsTitularMae As String
Private mvarsTitularMaeHebr As String
Private mvardtTitularDtNascPai As Date
Private mvariTitularDtNascPaiNoite As Integer
Private mvardtTitularDtFalecPai As Date
Private mvariTitularDtFalecPaiNoite As Integer
Private mvardtTitularDtNascMae As Date
Private mvariTitularDtNascMaeNoite As Integer
Private mvardtTitularDtFalecMae As Date
Private mvariTitularDtFalecMaeNoite As Integer
Private mvarsConjugeNome As String
Private mvarsConjugeNomeHebr As String
Private mvardtConjugeDtNasc As Date
Private mvariConjugeDtNascNoite As Integer
Private mvarsConjugeProfissao As String
Private mvarsConjugeNomeFirma As String
Private mvarlConjugeEnderecoCom As Long
Private mvarsConjugePai As String
Private mvarsConjugePaiHebr As String
Private mvarsConjugeMae As String
Private mvarsConjugeMaeHebr As String
Private mvardtConjugeDtNascPai As Date
Private mvariConjugeDtNascPaiNoite As Integer
Private mvardtConjugeDtFalecPai As Date
Private mvariConjugeDtFalecPaiNoite As Integer
Private mvardtConjugeDtNascMae As Date
Private mvariConjugeDtNascMaeNoite As Integer
Private mvardtConjugeDtFalecMae As Date
Private mvariConjugeDtFalecMaeNoite As Integer
Private mvardtConjugeDtFalec As Date
Private mvariConjugeDtFalecNoite As Integer
Private mvardtAtualizadoEm As Date
Private mvarlCodCliente As Long
Private mvardValorContribuicao As Double
Private mvarcolFilhos As New Collection
Private mvarcolFamiliaInfo As New Collection
Private mvarobjEnderecoRes As New ClassEndereco
Private mvarobjEnderecoCom As New ClassEndereco
Private mvarobjEnderecoComConj As New ClassEndereco
Private mvarsTitularSaudacao As String
Private mvarsConjugeSaudacao As String

Private mvarsTitularDtNascHebr As String
Private mvarsDataCasamentoHebr As String
Private mvarsTitularDtNascPaiHebr As String
Private mvarsTitularDtFalecPaiHebr As String
Private mvarsTitularDtNascMaeHebr As String
Private mvarsTitularDtFalecMaeHebr As String
Private mvarsConjugeDtNascHebr As String

Private mvarsConjugeDtNascPaiHebr As String
Private mvarsConjugeDtFalecPaiHebr As String
Private mvarsConjugeDtNascMaeHebr As String
Private mvarsConjugeDtFalecMaeHebr As String
Private mvarsConjugeDtFalecHebr As String

Public iAN As Integer
Public iCD As Integer
Public iCH As Integer
Public iCJ As Integer
Public iH As Integer
Public iH1 As Integer
Public iH2 As Integer
Public iLE As Integer
Public iLR As Integer
Public iPA As Integer
Public iRE As Integer
Public iSH As Integer
Public iSI As Integer
Public iTH As Integer
Public iVF As Integer

Public iANConj As Integer
Public iCDConj As Integer
Public iCHConj As Integer
Public iCJConj As Integer
Public iHConj As Integer
Public iH1Conj As Integer
Public iH2Conj As Integer
Public iLEConj As Integer
Public iLRConj As Integer
Public iPAConj As Integer
Public iREConj As Integer
Public iSHConj As Integer
Public iSIConj As Integer
Public iTHConj As Integer
Public iVFConj As Integer

Public Property Get sConjugeDtFalecHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtConjugeDtFalec = DATA_NULA Then
        sConjugeDtFalecHebr = ""
    Else
        sConjugeDtFalecHebr = GregToHeb(mvardtConjugeDtFalec + mvariConjugeDtFalecNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sConjugeDtFalecMaeHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtConjugeDtFalecMae = DATA_NULA Then
        sConjugeDtFalecMaeHebr = ""
    Else
        sConjugeDtFalecMaeHebr = GregToHeb(mvardtConjugeDtFalecMae + mvariConjugeDtFalecMaeNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sConjugeDtNascMaeHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtConjugeDtNascMae = DATA_NULA Then
        sConjugeDtNascMaeHebr = ""
    Else
        sConjugeDtNascMaeHebr = GregToHeb(mvardtConjugeDtNascMae + mvariConjugeDtNascMaeNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sConjugeDtFalecPaiHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtConjugeDtFalecPai = DATA_NULA Then
        sConjugeDtFalecPaiHebr = ""
    Else
        sConjugeDtFalecPaiHebr = GregToHeb(mvardtConjugeDtFalecPai + mvariConjugeDtFalecPaiNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sConjugeDtNascPaiHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtConjugeDtNascPai = DATA_NULA Then
        sConjugeDtNascPaiHebr = ""
    Else
        sConjugeDtNascPaiHebr = GregToHeb(mvardtConjugeDtNascPai + mvariConjugeDtNascPaiNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sConjugeDtNascHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtConjugeDtNasc = DATA_NULA Then
        sConjugeDtNascHebr = ""
    Else
        sConjugeDtNascHebr = GregToHeb(mvardtConjugeDtNasc + mvariConjugeDtNascNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sTitularDtFalecMaeHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtTitularDtFalecMae = DATA_NULA Then
        sTitularDtFalecMaeHebr = ""
    Else
        sTitularDtFalecMaeHebr = GregToHeb(mvardtTitularDtFalecMae + mvariTitularDtFalecMaeNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sTitularDtNascMaeHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtTitularDtNascMae = DATA_NULA Then
        sTitularDtNascMaeHebr = ""
    Else
        sTitularDtNascMaeHebr = GregToHeb(mvardtTitularDtNascMae + mvariTitularDtNascMaeNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sTitularDtFalecPaiHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtTitularDtFalecPai = DATA_NULA Then
        sTitularDtFalecPaiHebr = ""
    Else
        sTitularDtFalecPaiHebr = GregToHeb(mvardtTitularDtFalecPai + mvariTitularDtFalecPaiNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sTitularDtNascPaiHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtTitularDtNascPai = DATA_NULA Then
        sTitularDtNascPaiHebr = ""
    Else
        sTitularDtNascPaiHebr = GregToHeb(mvardtTitularDtNascPai + mvariTitularDtNascPaiNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sDataCasamentoHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtDataCasamento = DATA_NULA Then
        sDataCasamentoHebr = ""
    Else
        sDataCasamentoHebr = GregToHeb(mvardtDataCasamento + mvariDataCasamentoNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sTitularDtNascHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtTitularDtNasc = DATA_NULA Then
        sTitularDtNascHebr = ""
    Else
        sTitularDtNascHebr = GregToHeb(mvardtTitularDtNasc + mvariTitularDtNascNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Let sConjugeSaudacao(ByVal vData As String)
    mvarsConjugeSaudacao = vData
End Property

Public Property Get sConjugeSaudacao() As String
    sConjugeSaudacao = mvarsConjugeSaudacao
End Property

Public Property Let sTitularSaudacao(ByVal vData As String)
    mvarsTitularSaudacao = vData
End Property

Public Property Get sTitularSaudacao() As String
    sTitularSaudacao = mvarsTitularSaudacao
End Property

Public Property Set objEnderecoComConj(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoComConj = vData
End Property

Public Property Get objEnderecoComConj() As ClassEndereco
    Set objEnderecoComConj = mvarobjEnderecoComConj
End Property

Public Property Set objEnderecoCom(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoCom = vData
End Property

Public Property Get objEnderecoCom() As ClassEndereco
    Set objEnderecoCom = mvarobjEnderecoCom
End Property

Public Property Set objEnderecoRes(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoRes = vData
End Property

Public Property Get objEnderecoRes() As ClassEndereco
    Set objEnderecoRes = mvarobjEnderecoRes
End Property

Public Property Set colFamiliaInfo(ByVal vData As Object)
    Set mvarcolFamiliaInfo = vData
End Property

Public Property Get colFamiliaInfo() As Collection
    Set colFamiliaInfo = mvarcolFamiliaInfo
End Property

Public Property Set colFilhos(ByVal vData As Object)
    Set mvarcolFilhos = vData
End Property

Public Property Get colFilhos() As Collection
    Set colFilhos = mvarcolFilhos
End Property

Public Property Let lCodFamilia(ByVal vData As Long)
    mvarlCodFamilia = vData
End Property

Public Property Get lCodFamilia() As Long
    lCodFamilia = mvarlCodFamilia
End Property

Public Property Let sSobrenome(ByVal vData As String)
    mvarsSobrenome = vData
End Property

Public Property Get sSobrenome() As String
    sSobrenome = mvarsSobrenome
End Property

Public Property Let sTitularNome(ByVal vData As String)
    mvarsTitularNome = vData
End Property

Public Property Get sTitularNome() As String
    sTitularNome = mvarsTitularNome
End Property

Public Property Let sTitularNomeHebr(ByVal vData As String)
    mvarsTitularNomeHebr = vData
End Property

Public Property Get sTitularNomeHebr() As String
    sTitularNomeHebr = mvarsTitularNomeHebr
End Property

Public Property Let lTitularEnderecoRes(ByVal vData As Long)
    mvarlTitularEnderecoRes = vData
End Property

Public Property Get lTitularEnderecoRes() As Long
    lTitularEnderecoRes = mvarlTitularEnderecoRes
End Property

Public Property Let sTitularNomeFirma(ByVal vData As String)
    mvarsTitularNomeFirma = vData
End Property

Public Property Get sTitularNomeFirma() As String
    sTitularNomeFirma = mvarsTitularNomeFirma
End Property

Public Property Let lTitularEnderecoCom(ByVal vData As Long)
    mvarlTitularEnderecoCom = vData
End Property

Public Property Get lTitularEnderecoCom() As Long
    lTitularEnderecoCom = mvarlTitularEnderecoCom
End Property

Public Property Let iLocalCobranca(ByVal vData As Integer)
    mvariLocalCobranca = vData
End Property

Public Property Get iLocalCobranca() As Integer
    iLocalCobranca = mvariLocalCobranca
End Property

Public Property Let iEstadoCivil(ByVal vData As Integer)
    mvariEstadoCivil = vData
End Property

Public Property Get iEstadoCivil() As Integer
    iEstadoCivil = mvariEstadoCivil
End Property

Public Property Let sTitularProfissao(ByVal vData As String)
    mvarsTitularProfissao = vData
End Property

Public Property Get sTitularProfissao() As String
    sTitularProfissao = mvarsTitularProfissao
End Property

Public Property Let dtTitularDtNasc(ByVal vData As Date)
    mvardtTitularDtNasc = vData
End Property

Public Property Get dtTitularDtNasc() As Date
    dtTitularDtNasc = mvardtTitularDtNasc
End Property

Public Property Let iTitularDtNascNoite(ByVal vData As Integer)
    mvariTitularDtNascNoite = vData
End Property

Public Property Get iTitularDtNascNoite() As Integer
    iTitularDtNascNoite = mvariTitularDtNascNoite
End Property

Public Property Let dtDataCasamento(ByVal vData As Date)
    mvardtDataCasamento = vData
End Property

Public Property Get dtDataCasamento() As Date
    dtDataCasamento = mvardtDataCasamento
End Property

Public Property Let iDataCasamentoNoite(ByVal vData As Integer)
    mvariDataCasamentoNoite = vData
End Property

Public Property Get iDataCasamentoNoite() As Integer
    iDataCasamentoNoite = mvariDataCasamentoNoite
End Property

Public Property Let sCohenLeviIsrael(ByVal vData As String)
    mvarsCohenLeviIsrael = vData
End Property

Public Property Get sCohenLeviIsrael() As String
    sCohenLeviIsrael = mvarsCohenLeviIsrael
End Property

Public Property Let sTitularPai(ByVal vData As String)
    mvarsTitularPai = vData
End Property

Public Property Get sTitularPai() As String
    sTitularPai = mvarsTitularPai
End Property

Public Property Let sTitularPaiHebr(ByVal vData As String)
    mvarsTitularPaiHebr = vData
End Property

Public Property Get sTitularPaiHebr() As String
    sTitularPaiHebr = mvarsTitularPaiHebr
End Property

Public Property Let sTitularMae(ByVal vData As String)
    mvarsTitularMae = vData
End Property

Public Property Get sTitularMae() As String
    sTitularMae = mvarsTitularMae
End Property

Public Property Let sTitularMaeHebr(ByVal vData As String)
    mvarsTitularMaeHebr = vData
End Property

Public Property Get sTitularMaeHebr() As String
    sTitularMaeHebr = mvarsTitularMaeHebr
End Property

Public Property Let dtTitularDtNascPai(ByVal vData As Date)
    mvardtTitularDtNascPai = vData
End Property

Public Property Get dtTitularDtNascPai() As Date
    dtTitularDtNascPai = mvardtTitularDtNascPai
End Property

Public Property Let iTitularDtNascPaiNoite(ByVal vData As Integer)
    mvariTitularDtNascPaiNoite = vData
End Property

Public Property Get iTitularDtNascPaiNoite() As Integer
    iTitularDtNascPaiNoite = mvariTitularDtNascPaiNoite
End Property

Public Property Let dtTitularDtFalecPai(ByVal vData As Date)
    mvardtTitularDtFalecPai = vData
End Property

Public Property Get dtTitularDtFalecPai() As Date
    dtTitularDtFalecPai = mvardtTitularDtFalecPai
End Property

Public Property Let iTitularDtFalecPaiNoite(ByVal vData As Integer)
    mvariTitularDtFalecPaiNoite = vData
End Property

Public Property Get iTitularDtFalecPaiNoite() As Integer
    iTitularDtFalecPaiNoite = mvariTitularDtFalecPaiNoite
End Property

Public Property Let dtTitularDtNascMae(ByVal vData As Date)
    mvardtTitularDtNascMae = vData
End Property

Public Property Get dtTitularDtNascMae() As Date
    dtTitularDtNascMae = mvardtTitularDtNascMae
End Property

Public Property Let iTitularDtNascMaeNoite(ByVal vData As Integer)
    mvariTitularDtNascMaeNoite = vData
End Property

Public Property Get iTitularDtNascMaeNoite() As Integer
    iTitularDtNascMaeNoite = mvariTitularDtNascMaeNoite
End Property

Public Property Let dtTitularDtFalecMae(ByVal vData As Date)
    mvardtTitularDtFalecMae = vData
End Property

Public Property Get dtTitularDtFalecMae() As Date
    dtTitularDtFalecMae = mvardtTitularDtFalecMae
End Property

Public Property Let iTitularDtFalecMaeNoite(ByVal vData As Integer)
    mvariTitularDtFalecMaeNoite = vData
End Property

Public Property Get iTitularDtFalecMaeNoite() As Integer
    iTitularDtFalecMaeNoite = mvariTitularDtFalecMaeNoite
End Property

Public Property Let sConjugeNome(ByVal vData As String)
    mvarsConjugeNome = vData
End Property

Public Property Get sConjugeNome() As String
    sConjugeNome = mvarsConjugeNome
End Property

Public Property Let sConjugeNomeHebr(ByVal vData As String)
    mvarsConjugeNomeHebr = vData
End Property

Public Property Get sConjugeNomeHebr() As String
    sConjugeNomeHebr = mvarsConjugeNomeHebr
End Property

Public Property Let dtConjugeDtNasc(ByVal vData As Date)
    mvardtConjugeDtNasc = vData
End Property

Public Property Get dtConjugeDtNasc() As Date
    dtConjugeDtNasc = mvardtConjugeDtNasc
End Property

Public Property Let iConjugeDtNascNoite(ByVal vData As Integer)
    mvariConjugeDtNascNoite = vData
End Property

Public Property Get iConjugeDtNascNoite() As Integer
    iConjugeDtNascNoite = mvariConjugeDtNascNoite
End Property

Public Property Let sConjugeProfissao(ByVal vData As String)
    mvarsConjugeProfissao = vData
End Property

Public Property Get sConjugeProfissao() As String
    sConjugeProfissao = mvarsConjugeProfissao
End Property

Public Property Let sConjugeNomeFirma(ByVal vData As String)
    mvarsConjugeNomeFirma = vData
End Property

Public Property Get sConjugeNomeFirma() As String
    sConjugeNomeFirma = mvarsConjugeNomeFirma
End Property

Public Property Let lConjugeEnderecoCom(ByVal vData As Long)
    mvarlConjugeEnderecoCom = vData
End Property

Public Property Get lConjugeEnderecoCom() As Long
    lConjugeEnderecoCom = mvarlConjugeEnderecoCom
End Property

Public Property Let sConjugePai(ByVal vData As String)
    mvarsConjugePai = vData
End Property

Public Property Get sConjugePai() As String
    sConjugePai = mvarsConjugePai
End Property

Public Property Let sConjugePaiHebr(ByVal vData As String)
    mvarsConjugePaiHebr = vData
End Property

Public Property Get sConjugePaiHebr() As String
    sConjugePaiHebr = mvarsConjugePaiHebr
End Property

Public Property Let sConjugeMae(ByVal vData As String)
    mvarsConjugeMae = vData
End Property

Public Property Get sConjugeMae() As String
    sConjugeMae = mvarsConjugeMae
End Property

Public Property Let sConjugeMaeHebr(ByVal vData As String)
    mvarsConjugeMaeHebr = vData
End Property

Public Property Get sConjugeMaeHebr() As String
    sConjugeMaeHebr = mvarsConjugeMaeHebr
End Property

Public Property Let dtConjugeDtNascPai(ByVal vData As Date)
    mvardtConjugeDtNascPai = vData
End Property

Public Property Get dtConjugeDtNascPai() As Date
    dtConjugeDtNascPai = mvardtConjugeDtNascPai
End Property

Public Property Let iConjugeDtNascPaiNoite(ByVal vData As Integer)
    mvariConjugeDtNascPaiNoite = vData
End Property

Public Property Get iConjugeDtNascPaiNoite() As Integer
    iConjugeDtNascPaiNoite = mvariConjugeDtNascPaiNoite
End Property

Public Property Let dtConjugeDtFalecPai(ByVal vData As Date)
    mvardtConjugeDtFalecPai = vData
End Property

Public Property Get dtConjugeDtFalecPai() As Date
    dtConjugeDtFalecPai = mvardtConjugeDtFalecPai
End Property

Public Property Let iConjugeDtFalecPaiNoite(ByVal vData As Integer)
    mvariConjugeDtFalecPaiNoite = vData
End Property

Public Property Get iConjugeDtFalecPaiNoite() As Integer
    iConjugeDtFalecPaiNoite = mvariConjugeDtFalecPaiNoite
End Property

Public Property Let dtConjugeDtNascMae(ByVal vData As Date)
    mvardtConjugeDtNascMae = vData
End Property

Public Property Get dtConjugeDtNascMae() As Date
    dtConjugeDtNascMae = mvardtConjugeDtNascMae
End Property

Public Property Let iConjugeDtNascMaeNoite(ByVal vData As Integer)
    mvariConjugeDtNascMaeNoite = vData
End Property

Public Property Get iConjugeDtNascMaeNoite() As Integer
    iConjugeDtNascMaeNoite = mvariConjugeDtNascMaeNoite
End Property

Public Property Let dtConjugeDtFalecMae(ByVal vData As Date)
    mvardtConjugeDtFalecMae = vData
End Property

Public Property Get dtConjugeDtFalecMae() As Date
    dtConjugeDtFalecMae = mvardtConjugeDtFalecMae
End Property

Public Property Let iConjugeDtFalecMaeNoite(ByVal vData As Integer)
    mvariConjugeDtFalecMaeNoite = vData
End Property

Public Property Get iConjugeDtFalecMaeNoite() As Integer
    iConjugeDtFalecMaeNoite = mvariConjugeDtFalecMaeNoite
End Property

Public Property Let dtConjugeDtFalec(ByVal vData As Date)
    mvardtConjugeDtFalec = vData
End Property

Public Property Get dtConjugeDtFalec() As Date
    dtConjugeDtFalec = mvardtConjugeDtFalec
End Property

Public Property Let iConjugeDtFalecNoite(ByVal vData As Integer)
    mvariConjugeDtFalecNoite = vData
End Property

Public Property Get iConjugeDtFalecNoite() As Integer
    iConjugeDtFalecNoite = mvariConjugeDtFalecNoite
End Property

Public Property Let dtAtualizadoEm(ByVal vData As Date)
    mvardtAtualizadoEm = vData
End Property

Public Property Get dtAtualizadoEm() As Date
    dtAtualizadoEm = mvardtAtualizadoEm
End Property

Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Let dValorContribuicao(ByVal vData As Double)
    mvardValorContribuicao = vData
End Property

Public Property Get dValorContribuicao() As Double
    dValorContribuicao = mvardValorContribuicao
End Property

