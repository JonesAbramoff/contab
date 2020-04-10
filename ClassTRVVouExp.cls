VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVVouExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariTransacao As Integer
Private mvarsUsuario As String
Private mvardtData As Date
Private mvardHora As Double
Private mvarlNumVou As Long
Private mvarsTipVou As String
Private mvarsSerie As String
Private mvarlCliente As Long
Private mvarlNumFat As Long
Private mvardtDataFat As Date
Private mvardtDataCanc As Date
Private mvardtDataPag As Date
Private mvariExportado As Integer
Private mvarsUsuarioExportacao As String
Private mvardtDataExportacao As Date
Private mvardHoraExportacao As Double
Private mvarsNomeArq As String
Private mvarlAgenciaComissao As Long
Private mvarlEmissorComissao As Long
Private mvarlCorrentistaComissao As Long
Private mvarlRepresentanteComissao As Long
Private mvardPercComiAg As Double
Private mvardPercComiCor As Double
Private mvardPercComiRep As Double
Private mvardPercComiEmi As Double
Private mvardValorFat As Double
'ValorBrutoComOCR, ValorCMAComOCR, ValorCMCC, ValorCMR, ValorCMC, ValorCME, ValorAporte, ValorAporteCred, ValorNeto
Private mvardValorBrutoComOCR As Double
Private mvardValorCMAComOCR As Double
Private mvardValorCMCC As Double
Private mvardValorCMR As Double
Private mvardValorCMC As Double
Private mvardValorCME As Double
Private mvardValorAporte As Double
Private mvardValorAporteCred As Double
Private mvardValorNeto As Double

Public Property Let dValorNeto(ByVal vData As Double)
    mvardValorNeto = vData
End Property

Public Property Get dValorNeto() As Double
    dValorNeto = mvardValorNeto
End Property

Public Property Let dValorAporteCred(ByVal vData As Double)
    mvardValorAporteCred = vData
End Property

Public Property Get dValorAporteCred() As Double
    dValorAporteCred = mvardValorAporteCred
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property

Public Property Let dValorCME(ByVal vData As Double)
    mvardValorCME = vData
End Property

Public Property Get dValorCME() As Double
    dValorCME = mvardValorCME
End Property

Public Property Let dValorCMC(ByVal vData As Double)
    mvardValorCMC = vData
End Property

Public Property Get dValorCMC() As Double
    dValorCMC = mvardValorCMC
End Property

Public Property Let dValorCMR(ByVal vData As Double)
    mvardValorCMR = vData
End Property

Public Property Get dValorCMR() As Double
    dValorCMR = mvardValorCMR
End Property

Public Property Let dValorCMCC(ByVal vData As Double)
    mvardValorCMCC = vData
End Property

Public Property Get dValorCMCC() As Double
    dValorCMCC = mvardValorCMCC
End Property

Public Property Let dValorCMAComOCR(ByVal vData As Double)
    mvardValorCMAComOCR = vData
End Property

Public Property Get dValorCMAComOCR() As Double
    dValorCMAComOCR = mvardValorCMAComOCR
End Property

Public Property Let dValorBrutoComOCR(ByVal vData As Double)
    mvardValorBrutoComOCR = vData
End Property

Public Property Get dValorBrutoComOCR() As Double
    dValorBrutoComOCR = mvardValorBrutoComOCR
End Property

Public Property Let dValorFat(ByVal vData As Double)
    mvardValorFat = vData
End Property

Public Property Get dValorFat() As Double
    dValorFat = mvardValorFat
End Property

Public Property Let dPercComiEmi(ByVal vData As Double)
    mvardPercComiEmi = vData
End Property

Public Property Get dPercComiEmi() As Double
    dPercComiEmi = mvardPercComiEmi
End Property

Public Property Let dPercComiRep(ByVal vData As Double)
    mvardPercComiRep = vData
End Property

Public Property Get dPercComiRep() As Double
    dPercComiRep = mvardPercComiRep
End Property

Public Property Let dPercComiCor(ByVal vData As Double)
    mvardPercComiCor = vData
End Property

Public Property Get dPercComiCor() As Double
    dPercComiCor = mvardPercComiCor
End Property

Public Property Let dPercComiAg(ByVal vData As Double)
    mvardPercComiAg = vData
End Property

Public Property Get dPercComiAg() As Double
    dPercComiAg = mvardPercComiAg
End Property

Public Property Let lRepresentanteComissao(ByVal vData As Long)
    mvarlRepresentanteComissao = vData
End Property

Public Property Get lRepresentanteComissao() As Long
    lRepresentanteComissao = mvarlRepresentanteComissao
End Property

Public Property Let lCorrentistaComissao(ByVal vData As Long)
    mvarlCorrentistaComissao = vData
End Property

Public Property Get lCorrentistaComissao() As Long
    lCorrentistaComissao = mvarlCorrentistaComissao
End Property

Public Property Let lEmissorComissao(ByVal vData As Long)
    mvarlEmissorComissao = vData
End Property

Public Property Get lEmissorComissao() As Long
    lEmissorComissao = mvarlEmissorComissao
End Property

Public Property Let lAgenciaComissao(ByVal vData As Long)
    mvarlAgenciaComissao = vData
End Property

Public Property Get lAgenciaComissao() As Long
    lAgenciaComissao = mvarlAgenciaComissao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTransacao(ByVal vData As Integer)
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lNumFat(ByVal vData As Long)
    mvarlNumFat = vData
End Property

Public Property Get lNumFat() As Long
    lNumFat = mvarlNumFat
End Property

Public Property Let dtDataFat(ByVal vData As Date)
    mvardtDataFat = vData
End Property

Public Property Get dtDataFat() As Date
    dtDataFat = mvardtDataFat
End Property

Public Property Let dtDataCanc(ByVal vData As Date)
    mvardtDataCanc = vData
End Property

Public Property Get dtDataCanc() As Date
    dtDataCanc = mvardtDataCanc
End Property

Public Property Let dtDataPag(ByVal vData As Date)
    mvardtDataPag = vData
End Property

Public Property Get dtDataPag() As Date
    dtDataPag = mvardtDataPag
End Property

Public Property Let iExportado(ByVal vData As Integer)
    mvariExportado = vData
End Property

Public Property Get iExportado() As Integer
    iExportado = mvariExportado
End Property

Public Property Let sUsuarioExportacao(ByVal vData As String)
    mvarsUsuarioExportacao = vData
End Property

Public Property Get sUsuarioExportacao() As String
    sUsuarioExportacao = mvarsUsuarioExportacao
End Property

Public Property Let dtDataExportacao(ByVal vData As Date)
    mvardtDataExportacao = vData
End Property

Public Property Get dtDataExportacao() As Date
    dtDataExportacao = mvardtDataExportacao
End Property

Public Property Let dHoraExportacao(ByVal vData As Double)
    mvardHoraExportacao = vData
End Property

Public Property Get dHoraExportacao() As Double
    dHoraExportacao = mvardHoraExportacao
End Property

Public Property Let sNomeArq(ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Function Obtem_Dados_Geracao() As Long

On Error GoTo Erro_Obtem_Dados_Geracao

    mvarsUsuario = gsUsuario
    mvardtData = Date
    mvardHora = CDbl(Time)
    mvariExportado = DESMARCADO
    
    Obtem_Dados_Geracao = SUCESSO
    
    Exit Function

Erro_Obtem_Dados_Geracao:

    Obtem_Dados_Geracao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200183)

    End Select

    Exit Function
    
End Function

Public Function Obtem_Dados_Exportacao() As Long

On Error GoTo Erro_Obtem_Dados_Exportacao

    mvarsUsuarioExportacao = gsUsuario
    mvardtDataExportacao = Date
    mvardHoraExportacao = CDbl(Time)
    mvariExportado = MARCADO
    
    Obtem_Dados_Exportacao = SUCESSO

    Exit Function

Erro_Obtem_Dados_Exportacao:

    Obtem_Dados_Exportacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200183)

    End Select

    Exit Function
    
End Function

Private Sub Class_Initialize()
    mvardtDataExportacao = DATA_NULA
    mvardtData = DATA_NULA
    mvardtDataCanc = DATA_NULA
    mvardtDataFat = DATA_NULA
    mvardtDataPag = DATA_NULA
End Sub
