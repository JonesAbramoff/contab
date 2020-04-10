VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOcorrRemParcRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntParc As Long
Private mvariNumSeqOcorr As Integer
Private mvariFilialEmpresa As Integer
Private mvariCobrador As Integer
Private mvariCodOcorrencia As Integer
Private mvardtDataRegistro As Date
Private mvardtData As Date
Private mvariTituloVoltaCarteira As Integer
Private mvardtNovaDataVcto As Date
Private mvardJuros As Double
Private mvariInstrucao1 As Integer
Private mvariDiasDeProtesto1 As Integer
Private mvariInstrucao2 As Integer
Private mvariDiasDeProtesto2 As Integer
Private mvarlNumBordero As Long
Private mvardValorCobrado As Double
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
Private mvardSaldo As Double 'saldo da parcela que está sendo transferido
Private mvarsNumTitCobrador As String

Public Property Let sNumTitCobrador(ByVal vData As String)
    mvarsNumTitCobrador = vData
End Property

Public Property Get sNumTitCobrador() As String
    sNumTitCobrador = mvarsNumTitCobrador
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dValorCobrado(ByVal vData As Double)
    mvardValorCobrado = vData
End Property

Public Property Get dValorCobrado() As Double
    dValorCobrado = mvardValorCobrado
End Property

Public Property Let lNumBordero(ByVal vData As Long)
Attribute lNumBordero.VB_Description = "# do bordero de cobranca, se estiver zerado esta ocorrencia nao foi enviada"
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let iDiasDeProtesto2(ByVal vData As Integer)
    mvariDiasDeProtesto2 = vData
End Property

Public Property Get iDiasDeProtesto2() As Integer
    iDiasDeProtesto2 = mvariDiasDeProtesto2
End Property

Public Property Let iInstrucao2(ByVal vData As Integer)
    mvariInstrucao2 = vData
End Property

Public Property Get iInstrucao2() As Integer
    iInstrucao2 = mvariInstrucao2
End Property

Public Property Let iDiasDeProtesto1(ByVal vData As Integer)
    mvariDiasDeProtesto1 = vData
End Property

Public Property Get iDiasDeProtesto1() As Integer
    iDiasDeProtesto1 = mvariDiasDeProtesto1
End Property

Public Property Let iInstrucao1(ByVal vData As Integer)
    mvariInstrucao1 = vData
End Property

Public Property Get iInstrucao1() As Integer
    iInstrucao1 = mvariInstrucao1
End Property

Public Property Let dJuros(ByVal vData As Double)
Attribute dJuros.VB_Description = "diarios"
    mvardJuros = vData
End Property

Public Property Get dJuros() As Double
    dJuros = mvardJuros
End Property

Public Property Let dtNovaDataVcto(ByVal vData As Date)
Attribute dtNovaDataVcto.VB_Description = "se foi alterada a data de vcto"
    mvardtNovaDataVcto = vData
End Property

Public Property Get dtNovaDataVcto() As Date
    dtNovaDataVcto = mvardtNovaDataVcto
End Property

Public Property Let iTituloVoltaCarteira(ByVal vData As Integer)
Attribute iTituloVoltaCarteira.VB_Description = "0 ou 1: se o titulo foi trazido de volta p/carteira"
    mvariTituloVoltaCarteira = vData
End Property

Public Property Get iTituloVoltaCarteira() As Integer
    iTituloVoltaCarteira = mvariTituloVoltaCarteira
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
Attribute dtDataRegistro.VB_Description = "data da inclusao desta ocorrencia no sistema"
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let iCodOcorrencia(ByVal vData As Integer)
Attribute iCodOcorrencia.VB_Description = "01remessa, 02 baixa,..."
    mvariCodOcorrencia = vData
End Property

Public Property Get iCodOcorrencia() As Integer
    iCodOcorrencia = mvariCodOcorrencia
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let iNumSeqOcorr(ByVal vData As Integer)
Attribute iNumSeqOcorr.VB_Description = "sequencial de ocorrencia na parcela"
    mvariNumSeqOcorr = vData
End Property

Public Property Get iNumSeqOcorr() As Integer
    iNumSeqOcorr = mvariNumSeqOcorr
End Property

Public Property Let lNumIntParc(ByVal vData As Long)
Attribute lNumIntParc.VB_Description = "identifica uma parcela"
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function


