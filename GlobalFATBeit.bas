Attribute VB_Name = "GlobalFATBeit"
Option Explicit

Public Const STRING_NOME_PESSOA = 40
Public Const STRING_PROFISSAO = 40

Public Const TRIPO_COHEN = 1
Public Const TRIPO_LEVI = 2
Public Const TRIPO_ISRAEL = 3

Public Const STRING_TRIPO_COHEN = "C"
Public Const STRING_TRIPO_LEVI = "L"
Public Const STRING_TRIPO_ISRAEL = "I"

Public Const STRING_FAMILIATIPOINFO_DESCRICAO = 50
Public Const STRING_FAMILIATIPOINFO_SIGLA = 10

Public Const STRING_FAMILIA_SAUDACAO = 50

Public Const FAMILIAINFO_TITULAR = -2
Public Const FAMILIAINFO_CONJUGE = -1

Public Const STRING_PRONOME_TRATAMENTO_SR = "Sr."
Public Const STRING_PRONOME_TRATAMENTO_SRA = "Sra."
Public Const STRING_PRONOME_TRATAMENTO_SRTA = "Srta."
Public Const STRING_PRONOME_TRATAMENTO_RABINO = "Rabino"
Public Const STRING_PRONOME_TRATAMENTO_VOSSA_EXCELENCIA = "V. Ex.ª"
Public Const STRING_PRONOME_TRATAMENTO_VOSSA_MAGINIFICENCIA = "V. M."
Public Const STRING_PRONOME_TRATAMENTO_VOSSA_SENHORIA = "V. S.ª"
Public Const STRING_PRONOME_TRATAMENTO_MERITISSIMO_JUIZ = "M. Juiz"
Public Const STRING_PRONOME_TRATAMENTO_DOUTOR = "Dr."
Public Const STRING_PRONOME_TRATAMENTO_COMENDADOR = "Com."
Public Const STRING_PRONOME_TRATAMENTO_PROFESSOR = "Prof."

Type typeFamilias
    lCodFamilia As Long
    sSobrenome As String
    sTitularNome As String
    sTitularNomeHebr As String
    lTitularEnderecoRes As Long
    sTitularNomeFirma As String
    lTitularEnderecoCom As Long
    iLocalCobranca As Integer
    iEstadoCivil As Integer
    sTitularProfissao As String
    dtTitularDtNasc As Date
    iTitularDtNascNoite As Integer
    dtDataCasamento As Date
    iDataCasamentoNoite As Integer
    sCohenLeviIsrael As String
    sTitularPai As String
    sTitularPaiHebr As String
    sTitularMae As String
    sTitularMaeHebr As String
    dtTitularDtNascPai As Date
    iTitularDtNascPaiNoite As Integer
    dtTitularDtFalecPai As Date
    iTitularDtFalecPaiNoite As Integer
    dtTitularDtNascMae As Date
    iTitularDtNascMaeNoite As Integer
    dtTitularDtFalecMae As Date
    iTitularDtFalecMaeNoite As Integer
    sConjugeNome As String
    sConjugeNomeHebr As String
    dtConjugeDtNasc As Date
    iConjugeDtNascNoite As Integer
    sConjugeProfissao As String
    sConjugeNomeFirma As String
    lConjugeEnderecoCom As Long
    sConjugePai As String
    sConjugePaiHebr As String
    sConjugeMae As String
    sConjugeMaeHebr As String
    dtConjugeDtNascPai As Date
    iConjugeDtNascPaiNoite As Integer
    dtConjugeDtFalecPai As Date
    iConjugeDtFalecPaiNoite As Integer
    dtConjugeDtNascMae As Date
    iConjugeDtNascMaeNoite As Integer
    dtConjugeDtFalecMae As Date
    iConjugeDtFalecMaeNoite As Integer
    dtConjugeDtFalec As Date
    iConjugeDtFalecNoite As Integer
    dtAtualizadoEm As Date
    lCodCliente As Long
    dValorContribuicao As Double
    sTitularSaudacao As String
    sConjugeSaudacao As String
End Type

Type typeFilhosFamilias
    lCodFamilia As Long
    iSeqFilho As Integer
    sNome As String
    sNomeHebr As String
    dtDtNasc As Date
    iDtNascNoite As Integer
    dtDtFal As Date
    iDtFalNoite As Integer
    sTelefone As String
    sEmail As String
End Type

Type typeFamiliasTipoInfo
    iCodInfo As Integer
    sDescricao As String
    sSigla As String
    iValidoPara As Integer
    iPosicao As Integer
End Type

Type typeFamiliasInfo
    lCodFamilia As Long
    iSeq As Integer
    iCodInfo As Integer
    iValor As Integer
End Type

