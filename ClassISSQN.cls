VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassISSQN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodigo AS String
Private mvarsDescricao AS String
Private mvariTipo AS Integer
Private mvarlCodServNFe AS Long

Public Property Let sCodigo (ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo () AS String
    sCodigo= mvarsCodigo
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

Public Property Let iTipo (ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo () AS Integer
    iTipo= mvariTipo
End Property

Public Property Let lCodServNFe (ByVal vData As Long)
    mvarlCodServNFe = vData
End Property

Public Property Get lCodServNFe () AS Long
    lCodServNFe= mvarlCodServNFe
End Property
