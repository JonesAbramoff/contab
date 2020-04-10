VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEnvioDeEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCC As String
Private mvarsAssunto As String
Private mvarsAnexo As String
Private mvarsTextoRelac As String
Private mvarsModelo As String
Private mvarsDe As String
Private mvarsNomeExibicao As String
Private mvarobjTela As Object
Private mvarlClienteRelac As Long
Private mvarlNumIntDocParc As Long
Private mvariGeraRelac As Integer
Private mvarsEmail As String
Private mvarsMensagem As String
Private mvariFilialCliRelac As Integer
Private mvariLinha As Integer
Private mvarcolDadosAdicionais As New AdmColCampoValor
Private mvariCobranca As Integer
Private mvariConfirmacaoLeitura As Integer
Private mvarsEmailResp As String
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocModelo As Long

Public Property Let lNumIntDocModelo(ByVal vData As Long)
    mvarlNumIntDocModelo = vData
End Property

Public Property Get lNumIntDocModelo() As Long
    lNumIntDocModelo = mvarlNumIntDocModelo
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sEmailResp(ByVal vData As String)
    mvarsEmailResp = vData
End Property

Public Property Get sEmailResp() As String
    sEmailResp = mvarsEmailResp
End Property

Public Property Let iConfirmacaoLeitura(ByVal vData As Integer)
    mvariConfirmacaoLeitura = vData
End Property

Public Property Get iConfirmacaoLeitura() As Integer
    iConfirmacaoLeitura = mvariConfirmacaoLeitura
End Property

Public Property Let iCobranca(ByVal vData As Integer)
    mvariCobranca = vData
End Property

Public Property Get iCobranca() As Integer
    iCobranca = mvariCobranca
End Property

Public Property Get colDadosAdicionais() As AdmColCampoValor
    Set colDadosAdicionais = mvarcolDadosAdicionais
End Property

Public Property Set colDadosAdicionais(ByVal vData As AdmColCampoValor)
    Set mvarcolDadosAdicionais = vData
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property

Public Property Let iFilialCliRelac(ByVal vData As Integer)
    mvariFilialCliRelac = vData
End Property

Public Property Get iFilialCliRelac() As Integer
    iFilialCliRelac = mvariFilialCliRelac
End Property

Public Property Let sMensagem(ByVal vData As String)
    mvarsMensagem = vData
End Property

Public Property Get sMensagem() As String
    sMensagem = mvarsMensagem
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let iGeraRelac(ByVal vData As Integer)
    mvariGeraRelac = vData
End Property

Public Property Get iGeraRelac() As Integer
    iGeraRelac = mvariGeraRelac
End Property

Public Property Let lNumIntDocParc(ByVal vData As Long)
    mvarlNumIntDocParc = vData
End Property

Public Property Get lNumIntDocParc() As Long
    lNumIntDocParc = mvarlNumIntDocParc
End Property

Public Property Let lClienteRelac(ByVal vData As Long)
    mvarlClienteRelac = vData
End Property

Public Property Get lClienteRelac() As Long
    lClienteRelac = mvarlClienteRelac
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let sNomeExibicao(ByVal vData As String)
    mvarsNomeExibicao = vData
End Property

Public Property Get sNomeExibicao() As String
    sNomeExibicao = mvarsNomeExibicao
End Property

Public Property Let sDe(ByVal vData As String)
    mvarsDe = vData
End Property

Public Property Get sDe() As String
    sDe = mvarsDe
End Property

Public Property Let sAnexo(ByVal vData As String)
    mvarsAnexo = vData
End Property

Public Property Get sAnexo() As String
    sAnexo = mvarsAnexo
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let sCC(ByVal vData As String)
    mvarsCC = vData
End Property

Public Property Get sCC() As String
    sCC = mvarsCC
End Property

Public Property Let sAssunto(ByVal vData As String)
    mvarsAssunto = vData
End Property

Public Property Get sAssunto() As String
    sAssunto = mvarsAssunto
End Property

Public Property Let sTextoRelac(ByVal vData As String)
    mvarsTextoRelac = vData
End Property

Public Property Get sTextoRelac() As String
    sTextoRelac = mvarsTextoRelac
End Property

