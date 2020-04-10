Attribute VB_Name = "GlobalCRFATMgz"
Option Explicit

Public Const STRING_REFERENCIARATEIO As String = 30
Public Const STRING_DOSSIERATEIO As String = 30
Public Const COBRAR_RATEIO As Integer = 1

Public Const STRING_CORRESPONDENCIA_ID = 20

Public Const STRING_CONTRATO_ID = 30

Public Const STRING_PROCESSO_CONTRATO_ID = 30
Public Const STRING_PROCESSO_CONTRATO_DESCRICAO = 250
Public Const STRING_PROCESSO_CONTRATO_OBSERVACAO = 250

Type typeProcessoContrato
    sContrato As String
    lCliente As Long
    iSeq As Integer
    sProcesso As String
    iTipo As Integer
    sDescricao As String
    dValor As Double
    dtDataCobranca As Date
    sObservacao As String
End Type

Type typeTituloPagRateio
    
    lNumIntDocPag As Long
    iSeq As Integer
    lCliente As Long
    dValor As Double
    lHistorico As Long
    sReferencia As String
    sDossie As String
    iCobrar As Integer
    lNumIntDocRec As Long
    lND As Long
    dtDataGerND As Date

End Type
