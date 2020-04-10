VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassApontProdPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc AS Long
Private mvarlNumIntDocApont AS Long
Private mvariSeq AS Integer
Private mvarsProduto AS String
Private mvarsUM AS String
Private mvardQtd AS Double
Private mvardCusto AS Double
Private mvarsObs AS String

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let lNumIntDocApont (ByVal vData As Long)
    mvarlNumIntDocApont = vData
End Property

Public Property Get lNumIntDocApont () AS Long
    lNumIntDocApont= mvarlNumIntDocApont
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let sUM (ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM () AS String
    sUM= mvarsUM
End Property

Public Property Let dQtd (ByVal vData As Double)
    mvardQtd = vData
End Property

Public Property Get dQtd () AS Double
    dQtd= mvardQtd
End Property

Public Property Let dCusto (ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto () AS Double
    dCusto= mvardCusto
End Property

Public Property Let sObs (ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs () AS String
    sObs= mvarsObs
End Property
