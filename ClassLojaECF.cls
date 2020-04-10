VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariImprimeItemAItem As Integer
Private mvarsNatOpPadrao As String
Private mvarsMensagemCupom As String
Private mvariOperadorIgualVendedor As Integer
Private mvariTabelaPreco As Integer
Private mvariTelaVendaMP As Integer
Private mvarobjECF As New ClassECF
Private mvariAtualizacaoECF As Integer
Private mvarsSimboloMoeda As String
Private mvariHorarioVerao As Integer
Private mvarcolAliquotaICMS As New Collection
Private mvariCupomDescreveFormaPagto As Integer
Private mvarlEspacoEntreLinhas As Long
Private mvariLinhasEntreCupons As Integer
Private mvarsNumLimRO As String
Private mvarsTruncamentoArredondamento As String
Private mvariGerenteAutoriza As Integer
Private mvariVendedorObrigatorio As Integer
Private mvariAbreAposFechamento As Integer
Private mvarsFTPURL As String
Private mvarsFTPUserName As String
Private mvarsFTPPassword As String
Private mvarsFTPDiretorio As String
Private mvarlIntervaloTrans As Long
Private mvardtTime As Date
Private mvardtData As Date
Private mvarsDirDadosCCC As String

Public Property Let sDirDadosCCC(ByVal vData As String)
    mvarsDirDadosCCC = vData
End Property

Public Property Get sDirDadosCCC() As String
    sDirDadosCCC = mvarsDirDadosCCC
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtTime(ByVal vData As Date)
    mvardtTime = vData
End Property

Public Property Get dtTime() As Date
    dtTime = mvardtTime
End Property

Public Property Let lIntervaloTrans(ByVal vData As Long)
    mvarlIntervaloTrans = vData
End Property

Public Property Get lIntervaloTrans() As Long
    lIntervaloTrans = mvarlIntervaloTrans
End Property

Public Property Let sFTPDiretorio(ByVal vData As String)
    mvarsFTPDiretorio = vData
End Property

Public Property Get sFTPDiretorio() As String
    sFTPDiretorio = mvarsFTPDiretorio
End Property

Public Property Let sFTPPassword(ByVal vData As String)
    mvarsFTPPassword = vData
End Property

Public Property Get sFTPPassword() As String
    sFTPPassword = mvarsFTPPassword
End Property

Public Property Let sFTPUserName(ByVal vData As String)
    mvarsFTPUserName = vData
End Property

Public Property Get sFTPUserName() As String
    sFTPUserName = mvarsFTPUserName
End Property

Public Property Let sFTPURL(ByVal vData As String)
    mvarsFTPURL = vData
End Property

Public Property Get sFTPURL() As String
    sFTPURL = mvarsFTPURL
End Property

Public Property Let iAbreAposFechamento(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariAbreAposFechamento = vData
End Property


Public Property Get iAbreAposFechamento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iAbreAposFechamento = mvariAbreAposFechamento
End Property


Public Property Let iVendedorObrigatorio(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariVendedorObrigatorio = vData
End Property


Public Property Get iVendedorObrigatorio() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iVendedorObrigatorio = mvariVendedorObrigatorio
End Property

Public Property Let iGerenteAutoriza(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariGerenteAutoriza = vData
End Property


Public Property Get iGerenteAutoriza() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iGerenteAutoriza = mvariGerenteAutoriza
End Property

Public Property Let sTruncamentoArredondamento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTruncamentoArredondamento = 5
    mvarsTruncamentoArredondamento = vData
End Property


Public Property Get sTruncamentoArredondamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTruncamentoArredondamento
    sTruncamentoArredondamento = mvarsTruncamentoArredondamento
End Property



Public Property Let sNumLimRO(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNumLimRO = 5
    mvarsNumLimRO = vData
End Property


Public Property Get sNumLimRO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNumLimRO
    sNumLimRO = mvarsNumLimRO
End Property



Public Property Let iLinhasEntreCupons(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLinhasEntreCupons = 5
    mvariLinhasEntreCupons = vData
End Property


Public Property Get iLinhasEntreCupons() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLinhasEntreCupons
    iLinhasEntreCupons = mvariLinhasEntreCupons
End Property



Public Property Let lEspacoEntreLinhas(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lEspacoEntreLinhas = 5
    mvarlEspacoEntreLinhas = vData
End Property


Public Property Get lEspacoEntreLinhas() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lEspacoEntreLinhas
    lEspacoEntreLinhas = mvarlEspacoEntreLinhas
End Property



Public Property Let iCupomDescreveFormaPagto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCupomDescreveFormaPagto = 5
    mvariCupomDescreveFormaPagto = vData
End Property


Public Property Get iCupomDescreveFormaPagto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCupomDescreveFormaPagto
    iCupomDescreveFormaPagto = mvariCupomDescreveFormaPagto
End Property



Public Property Set colAliquotaICMS(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colAliquotaICMS = Form1
    Set mvarcolAliquotaICMS = vData
End Property


Public Property Get colAliquotaICMS() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colAliquotaICMS
    Set colAliquotaICMS = mvarcolAliquotaICMS
End Property



Public Property Let iHorarioVerao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iHorarioVerao = 5
    mvariHorarioVerao = vData
End Property


Public Property Get iHorarioVerao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iHorarioVerao
    iHorarioVerao = mvariHorarioVerao
End Property



Public Property Let sSimboloMoeda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSimboloMoeda = 5
    mvarsSimboloMoeda = vData
End Property


Public Property Get sSimboloMoeda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSimboloMoeda
    sSimboloMoeda = mvarsSimboloMoeda
End Property



Public Property Let iAtualizacaoECF(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAtualizacaoECF = 5
    mvariAtualizacaoECF = vData
End Property


Public Property Get iAtualizacaoECF() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAtualizacaoECF
    iAtualizacaoECF = mvariAtualizacaoECF
End Property



Friend Property Set objECF(ByVal vData As ClassECF)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objECF = Form1
    Set mvarobjECF = vData
End Property


Friend Property Get objECF() As ClassECF
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objECF
    Set objECF = mvarobjECF
End Property




Public Property Let iTelaVendaMP(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTelaVendaMP = 5
    mvariTelaVendaMP = vData
End Property


Public Property Get iTelaVendaMP() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTelaVendaMP
    iTelaVendaMP = mvariTelaVendaMP
End Property



Public Property Let iTabelaPreco(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTabelaPreco = 5
    mvariTabelaPreco = vData
End Property


Public Property Get iTabelaPreco() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTabelaPreco
    iTabelaPreco = mvariTabelaPreco
End Property



Public Property Let iOperadorIgualVendedor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iOperadorIgualVendedor = 5
    mvariOperadorIgualVendedor = vData
End Property


Public Property Get iOperadorIgualVendedor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iOperadorIgualVendedor
    iOperadorIgualVendedor = mvariOperadorIgualVendedor
End Property



Public Property Let sMensagemCupom(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMensagemCupom = 5
    mvarsMensagemCupom = vData
End Property


Public Property Get sMensagemCupom() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMensagemCupom
    sMensagemCupom = mvarsMensagemCupom
End Property



Public Property Let sNatOpPadrao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNatOpPadrao = 5
    mvarsNatOpPadrao = vData
End Property


Public Property Get sNatOpPadrao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNatOpPadrao
    sNatOpPadrao = mvarsNatOpPadrao
End Property



Public Property Let iImprimeItemAItem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iImprimeItemAItem = 5
    mvariImprimeItemAItem = vData
End Property


Public Property Get iImprimeItemAItem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iImprimeItemAItem
    iImprimeItemAItem = mvariImprimeItemAItem
End Property

