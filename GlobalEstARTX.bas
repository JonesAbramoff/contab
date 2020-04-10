Attribute VB_Name = "GlobalEstARTX"

Public Const ETAPA_CORTE = 1
Public Const ETAPA_FORRO = 2
Public Const ETAPA_MONTAGEM = 3

Public Const ARTX_CATEGORIA_TIPOCOURO = "TIPO DE COURO"


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeUsuProdArtlux
    iFilialEmpresa As Integer
    sCodUsuario As String
    iAcessoCorte As Integer
    iAcessoForro As Integer
    iAcessoMontagem As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeOCArtlux
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    sProduto As String
    dQuantidade As Double
    sUsuCorte As String
    dtDataIniCorte As Date
    dHoraIniCorte As Double
    dtDataFimCorte As Date
    dHoraFimCorte As Double
    sUsuForro As String
    dtDataIniForro As Date
    dHoraIniForro As Double
    dtDataFimForro As Date
    dHoraFimForro As Double
    dQuantidadeProd As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeOCProdArtlux
    lNumIntDoc As Long
    lNumIntDocOC As Long
    iSeq As Integer
    lNumIntDocMovEst As Long
    sUsuMontagem As String
    dtDataIniMontagem As Date
    dtDataFimMontagem As Date
    dQuantidadePreProd As Double
    dQuantidadeProd As Double
End Type


