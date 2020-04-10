VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEmailConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsUsuario As String
Private mvarsSMTP As String
Private mvarsSMTPUsu As String
Private mvarsSMTPSenha As String
Private mvarlSMTPPorta As Long
Private mvariSSL As Integer
Private mvarsEmail As String
Private mvarsNome As String
Private mvariConfirmacaoLeitura As Integer
Private mvariPgmEmail As Integer

Public Property Let iPgmEmail(ByVal vData As Integer)
    mvariPgmEmail = vData
End Property

Public Property Get iPgmEmail() As Integer
    iPgmEmail = mvariPgmEmail
End Property

Public Property Let iConfirmacaoLeitura(ByVal vData As Integer)
    mvariConfirmacaoLeitura = vData
End Property

Public Property Get iConfirmacaoLeitura() As Integer
    iConfirmacaoLeitura = mvariConfirmacaoLeitura
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let iSSL(ByVal vData As Integer)
    mvariSSL = vData
End Property

Public Property Get iSSL() As Integer
    iSSL = mvariSSL
End Property

Public Property Let lSMTPPorta(ByVal vData As Long)
    mvarlSMTPPorta = vData
End Property

Public Property Get lSMTPPorta() As Long
    lSMTPPorta = mvarlSMTPPorta
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let sSMTP(ByVal vData As String)
    mvarsSMTP = vData
End Property

Public Property Get sSMTP() As String
    sSMTP = mvarsSMTP
End Property

Public Property Let sSMTPUsu(ByVal vData As String)
    mvarsSMTPUsu = vData
End Property

Public Property Get sSMTPUsu() As String
    sSMTPUsu = mvarsSMTPUsu
End Property

Public Property Let sSMTPSenha(ByVal vData As String)
    mvarsSMTPSenha = vData
End Property

Public Property Get sSMTPSenha() As String
    sSMTPSenha = mvarsSMTPSenha
End Property
