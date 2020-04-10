VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassCodServMun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodIBGE AS String
Private mvarsCodServ AS String
Private mvarsISSQNBase AS String
Private mvariPadraoISSQN AS Integer
Private mvarsDescricao1 AS String
Private mvarsDescricao2 AS String
Private mvardAliquota AS Double

Public Property Let sCodIBGE (ByVal vData As String)
    mvarsCodIBGE = vData
End Property

Public Property Get sCodIBGE () AS String
    sCodIBGE= mvarsCodIBGE
End Property

Public Property Let sCodServ (ByVal vData As String)
    mvarsCodServ = vData
End Property

Public Property Get sCodServ () AS String
    sCodServ= mvarsCodServ
End Property

Public Property Let sISSQNBase (ByVal vData As String)
    mvarsISSQNBase = vData
End Property

Public Property Get sISSQNBase () AS String
    sISSQNBase= mvarsISSQNBase
End Property

Public Property Let iPadraoISSQN (ByVal vData As Integer)
    mvariPadraoISSQN = vData
End Property

Public Property Get iPadraoISSQN () AS Integer
    iPadraoISSQN= mvariPadraoISSQN
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

Public Property Let dAliquota (ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota () AS Double
    dAliquota= mvardAliquota
End Property
