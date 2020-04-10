Attribute VB_Name = "GlobalFATAF"
Option Explicit

Public Const STRING_CLIENTESAF_SEXO = 1
Public Const STRING_CLIENTESAF_RGORGAOEMISSOR = 20
Public Const STRING_CLIENTESAF_LOCALTRABALHO = 80
Public Const STRING_CLIENTESAF_CARGO = 50
Public Const STRING_CLIENTESAF_NOBENEF = 20
Public Const STRING_CLIENTESAF_CONTRSOC = 50
Public Const STRING_CLIENTESAF_OBS1 = 255
Public Const STRING_CLIENTESAF_OBS2 = 255

Public Const EMPRESA_ELETRONUCLEAR = 2
Public Const EMPRESA_FURNAS = 3
Public Const EMPRESA_FRG = 4

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeClientesAF
    lCliente As Long
    iTipoAssociado As Integer
    lMatriculaPrinc As Long
    iStatusAssociado As Integer
    lMatriculaSec As Long
    lMatriculaAF As Long
    lMatriculaFRG As Long
    iEmpresa1 As Integer
    iEmpresa2 As Integer
    sSexo As String
    sRGOrgaoEmissor As String
    dtDataExpedicaoRG As Date
    dtDataNascimento As Date
    dtDataInscricao As Date
    sLocalTrabalho As String
    dtDataAdmissaoFurnas As Date
    sCargo As String
    dtDataAposINSS As Date
    dtDataAposFRG As Date
    iTipoApos As Integer
    dtDataConBenf As Date
    sNoBenef As String
    dtDataFalecimento As Date
    sContrSoc As String
    sObservacao1 As String
    sObservacao2 As String
    iBenemerito As Integer
    iFundador As Integer
    dtDataUltAtualizacao As Date
    iMatriculaPrincDV As Integer
End Type

