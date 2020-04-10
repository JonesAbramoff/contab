VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntBordero As Long
Private mvariExcluido As Integer
Private mvariCodConta As Integer
Private mvardtDataEmissao As Date
Private mvarlNumero As Long
Private mvarsNomeArq As String
Private mvardtDataEnvio As Date
Private mvardtDataVencimento As Date
Private mvariTipoDeCobranca As Integer
Private mvariTitOutroBanco As Integer
Private mvariNumArqRemessa As Integer
Private mvarobjTela As Object

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let iNumArqRemessa(ByVal vData As Integer)
    mvariNumArqRemessa = vData
End Property

Public Property Get iNumArqRemessa() As Integer
    iNumArqRemessa = mvariNumArqRemessa
End Property

Public Property Let iTipoDeCobranca(ByVal vData As Integer)
    mvariTipoDeCobranca = vData
End Property

Public Property Get iTipoDeCobranca() As Integer
    iTipoDeCobranca = mvariTipoDeCobranca
End Property

Public Property Let iTitOutroBanco(ByVal vData As Integer)
    mvariTitOutroBanco = vData
End Property

Public Property Get iTitOutroBanco() As Integer
    iTitOutroBanco = mvariTitOutroBanco
End Property

Public Property Let dtDataEnvio(ByVal vData As Date)
    mvardtDataEnvio = vData
End Property

Public Property Get dtDataEnvio() As Date
    dtDataEnvio = mvardtDataEnvio
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let sNomeArq(ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iExcluido(ByVal vData As Integer)
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property

Public Property Let lNumIntBordero(ByVal vData As Long)
    mvarlNumIntBordero = vData
End Property

Public Property Get lNumIntBordero() As Long
    lNumIntBordero = mvarlNumIntBordero
End Property



