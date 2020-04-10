VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPVouPassageiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsTipoVou As String
Private mvarsSerie As String
Private mvarlNumVou As Long
Private mvariSeq As Integer
Private mvarsNome As String
Private mvardtDataNascimento As Date
Private mvarsTipoDocumento As String
Private mvarsNumeroDocumento As String
Private mvarsSexo As String
Private mvardValorPago As Double
Private mvariStatus As Integer
Private mvariTitular As Integer
Private mvardValorPagoEmi As Double
Private mvarsPrimeiroNome As String
Private mvarsSobreNome As String

Public Property Let sSobreNome(ByVal vData As String)
    mvarsSobreNome = vData
End Property

Public Property Get sSobreNome() As String
    sSobreNome = mvarsSobreNome
End Property

Public Property Let sPrimeiroNome(ByVal vData As String)
    mvarsPrimeiroNome = vData
End Property

Public Property Get sPrimeiroNome() As String
    sPrimeiroNome = mvarsPrimeiroNome
End Property

Public Property Let dValorPagoEmi(ByVal vData As Double)
    mvardValorPagoEmi = vData
End Property

Public Property Get dValorPagoEmi() As Double
    dValorPagoEmi = mvardValorPagoEmi
End Property

Public Property Let sTipoVou(ByVal vData As String)
    mvarsTipoVou = vData
End Property

Public Property Get sTipoVou() As String
    sTipoVou = mvarsTipoVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let dtDataNascimento(ByVal vData As Date)
    mvardtDataNascimento = vData
End Property

Public Property Get dtDataNascimento() As Date
    dtDataNascimento = mvardtDataNascimento
End Property

Public Property Let sTipoDocumento(ByVal vData As String)
    mvarsTipoDocumento = vData
End Property

Public Property Get sTipoDocumento() As String
    sTipoDocumento = mvarsTipoDocumento
End Property

Public Property Let sNumeroDocumento(ByVal vData As String)
    mvarsNumeroDocumento = vData
End Property

Public Property Get sNumeroDocumento() As String
    sNumeroDocumento = mvarsNumeroDocumento
End Property

Public Property Let sSexo(ByVal vData As String)
    mvarsSexo = vData
End Property

Public Property Get sSexo() As String
    sSexo = mvarsSexo
End Property

Public Property Let dValorPago(ByVal vData As Double)
    mvardValorPago = vData
End Property

Public Property Get dValorPago() As Double
    dValorPago = mvardValorPago
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iTitular(ByVal vData As Integer)
    mvariTitular = vData
End Property

Public Property Get iTitular() As Integer
    iTitular = mvariTitular
End Property
