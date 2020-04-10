VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMOUtiPerPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeq As Integer
Private mvariCodMO As Integer
Private mvarsDescricaoMO As String
Private mvardCustoPrevInf As Double
Private mvardQuantidadePrevInf As Double
Private mvardCustoPrevCalc As Double
Private mvardQuantidadePrevCalc As Double
Private mvardCustoRealInf As Double
Private mvardQuantidadeRealInf As Double
Private mvardCustoRealCalc As Double
Private mvardQuantidadeRealCalc As Double
Private mvarsUM As String
Private mvarlNumIntDocEtapaPRJMO As Long
Private mvarlNumIntDocPRJ As Long
Private mvarsProjeto As String
Private mvarsDescricaoPRJ As String
Private mvarlCliente As Long
Private mvarlNumIntDocEtapa As Long
Private mvarsEtapa As String
Private mvarsDescricaoEtapa As String
Private mvardtData As Date
Private mvardPercEtapa As Double
Private mvariTipo As Integer
Private mvarsObs As String

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let dPercEtapa(ByVal vData As Double)
    mvardPercEtapa = vData
End Property

Public Property Get dPercEtapa() As Double
    dPercEtapa = mvardPercEtapa
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iCodMO(ByVal vData As Integer)
    mvariCodMO = vData
End Property

Public Property Get iCodMO() As Integer
    iCodMO = mvariCodMO
End Property

Public Property Let sDescricaoMO(ByVal vData As String)
    mvarsDescricaoMO = vData
End Property

Public Property Get sDescricaoMO() As String
    sDescricaoMO = mvarsDescricaoMO
End Property

Public Property Let dCustoPrevInf(ByVal vData As Double)
    mvardCustoPrevInf = vData
End Property

Public Property Get dCustoPrevInf() As Double
    dCustoPrevInf = mvardCustoPrevInf
End Property

Public Property Let dQuantidadePrevInf(ByVal vData As Double)
    mvardQuantidadePrevInf = vData
End Property

Public Property Get dQuantidadePrevInf() As Double
    dQuantidadePrevInf = mvardQuantidadePrevInf
End Property

Public Property Let dCustoPrevCalc(ByVal vData As Double)
    mvardCustoPrevCalc = vData
End Property

Public Property Get dCustoPrevCalc() As Double
    dCustoPrevCalc = mvardCustoPrevCalc
End Property

Public Property Let dQuantidadePrevCalc(ByVal vData As Double)
    mvardQuantidadePrevCalc = vData
End Property

Public Property Get dQuantidadePrevCalc() As Double
    dQuantidadePrevCalc = mvardQuantidadePrevCalc
End Property

Public Property Let dCustoRealInf(ByVal vData As Double)
    mvardCustoRealInf = vData
End Property

Public Property Get dCustoRealInf() As Double
    dCustoRealInf = mvardCustoRealInf
End Property

Public Property Let dQuantidadeRealInf(ByVal vData As Double)
    mvardQuantidadeRealInf = vData
End Property

Public Property Get dQuantidadeRealInf() As Double
    dQuantidadeRealInf = mvardQuantidadeRealInf
End Property

Public Property Let dCustoRealCalc(ByVal vData As Double)
    mvardCustoRealCalc = vData
End Property

Public Property Get dCustoRealCalc() As Double
    dCustoRealCalc = mvardCustoRealCalc
End Property

Public Property Let dQuantidadeRealCalc(ByVal vData As Double)
    mvardQuantidadeRealCalc = vData
End Property

Public Property Get dQuantidadeRealCalc() As Double
    dQuantidadeRealCalc = mvardQuantidadeRealCalc
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let lNumIntDocEtapaPRJMO(ByVal vData As Long)
    mvarlNumIntDocEtapaPRJMO = vData
End Property

Public Property Get lNumIntDocEtapaPRJMO() As Long
    lNumIntDocEtapaPRJMO = mvarlNumIntDocEtapaPRJMO
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let sProjeto(ByVal vData As String)
    mvarsProjeto = vData
End Property

Public Property Get sProjeto() As String
    sProjeto = mvarsProjeto
End Property

Public Property Let sDescricaoPRJ(ByVal vData As String)
    mvarsDescricaoPRJ = vData
End Property

Public Property Get sDescricaoPRJ() As String
    sDescricaoPRJ = mvarsDescricaoPRJ
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Public Property Let sEtapa(ByVal vData As String)
    mvarsEtapa = vData
End Property

Public Property Get sEtapa() As String
    sEtapa = mvarsEtapa
End Property

Public Property Let sDescricaoEtapa(ByVal vData As String)
    mvarsDescricaoEtapa = vData
End Property

Public Property Get sDescricaoEtapa() As String
    sDescricaoEtapa = mvarsDescricaoEtapa
End Property
