VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTRVProdLimites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsProduto AS String
Private mvarlCodServ AS Long
Private mvarsDescProd AS String
Private mvardValorLimite AS Double
Private mvariMoeda AS Integer
Private mvarsDescricao1 AS String
Private mvarsDescricao2 AS String
Private mvariTipo AS Integer

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let lCodServ (ByVal vData As Long)
    mvarlCodServ = vData
End Property

Public Property Get lCodServ () AS Long
    lCodServ= mvarlCodServ
End Property

Public Property Let sDescProd (ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd () AS String
    sDescProd= mvarsDescProd
End Property

Public Property Let dValorLimite (ByVal vData As Double)
    mvardValorLimite = vData
End Property

Public Property Get dValorLimite () AS Double
    dValorLimite= mvardValorLimite
End Property

Public Property Let iMoeda (ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda () AS Integer
    iMoeda= mvariMoeda
End Property

Public Property Let sDescricao1 (ByVal vData As String)
    mvarsDescricao1 = vData
End Property

Public Property Get sDescricao1 () AS String
    sDescricao1= mvarsDescricao1
End Property

Public Property Let sDescricao2 (ByVal vData As String)
    mvarsDescricao2 = vData
End Property

Public Property Get sDescricao2 () AS String
    sDescricao2= mvarsDescricao2
End Property

Public Property Let iTipo (ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo () AS Integer
    iTipo= mvariTipo
End Property
