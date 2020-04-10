VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassFamiliasTipoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodInfo AS Integer
Private mvarsDescricao AS String
Private mvarsSigla AS String
Private mvariValidoPara AS Integer
Private mvariPosicao AS Integer

Public Property Let iCodInfo (ByVal vData As Integer)
    mvariCodInfo = vData
End Property

Public Property Get iCodInfo () AS Integer
    iCodInfo= mvariCodInfo
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

Public Property Let sSigla (ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla () AS String
    sSigla= mvarsSigla
End Property

Public Property Let iValidoPara (ByVal vData As Integer)
    mvariValidoPara = vData
End Property

Public Property Get iValidoPara () AS Integer
    iValidoPara= mvariValidoPara
End Property

Public Property Let iPosicao (ByVal vData As Integer)
    mvariPosicao = vData
End Property

Public Property Get iPosicao () AS Integer
    iPosicao= mvariPosicao
End Property
