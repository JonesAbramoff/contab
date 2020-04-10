VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrancaEmailPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvariAtrasoDe As Integer
Private mvariAtrasoAte As Integer
Private mvarsCC As String
Private mvarsAssunto As String
Private mvarsAnexo As String
Private mvarsMensagem As String
Private mvarsCCValor As String
Private mvarsAssuntoValor As String
Private mvarsAnexoValor As String
Private mvarsMensagemValor As String
Private mvarsModelo As String
Private mvariTipo As Integer
Private mvarsDe As String
Private mvarsNomeExibicao As String
Private mvarsUsuarioExclusivo As String
Private mvariConfirmacaoLeitura As Integer
Private mvarsEmailResp As String

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

Public Property Let sUsuarioExclusivo(ByVal vData As String)
    mvarsUsuarioExclusivo = vData
End Property

Public Property Get sUsuarioExclusivo() As String
    sUsuarioExclusivo = mvarsUsuarioExclusivo
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

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sAnexo(ByVal vData As String)
    mvarsAnexo = vData
End Property

Public Property Get sAnexo() As String
    sAnexo = mvarsAnexo
End Property

Public Property Let sAnexoValor(ByVal vData As String)
    mvarsAnexoValor = vData
End Property

Public Property Get sAnexoValor() As String
    sAnexoValor = mvarsAnexoValor
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iAtrasoDe(ByVal vData As Integer)
    mvariAtrasoDe = vData
End Property

Public Property Get iAtrasoDe() As Integer
    iAtrasoDe = mvariAtrasoDe
End Property

Public Property Let iAtrasoAte(ByVal vData As Integer)
    mvariAtrasoAte = vData
End Property

Public Property Get iAtrasoAte() As Integer
    iAtrasoAte = mvariAtrasoAte
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

Public Property Let sMensagem(ByVal vData As String)
    mvarsMensagem = vData
End Property

Public Property Get sMensagem() As String
    sMensagem = mvarsMensagem
End Property

Public Property Let sCCValor(ByVal vData As String)
    mvarsCCValor = vData
End Property

Public Property Get sCCValor() As String
    sCCValor = mvarsCCValor
End Property

Public Property Let sAssuntoValor(ByVal vData As String)
    mvarsAssuntoValor = vData
End Property

Public Property Get sAssuntoValor() As String
    sAssuntoValor = mvarsAssuntoValor
End Property

Public Property Let sMensagemValor(ByVal vData As String)
    mvarsMensagemValor = vData
End Property

Public Property Get sMensagemValor() As String
    sMensagemValor = mvarsMensagemValor
End Property
