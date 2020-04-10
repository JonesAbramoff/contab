VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTRPVouContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsTipoVou AS String
Private mvarsSerie AS String
Private mvarlNumVou AS Long
Private mvariSeq AS Integer
Private mvarsNome AS String
Private mvarsTelefone AS String

Public Property Let sTipoVou (ByVal vData As String)
    mvarsTipoVou = vData
End Property

Public Property Get sTipoVou () AS String
    sTipoVou= mvarsTipoVou
End Property

Public Property Let sSerie (ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie () AS String
    sSerie= mvarsSerie
End Property

Public Property Let lNumVou (ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou () AS Long
    lNumVou= mvarlNumVou
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let sNome (ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome () AS String
    sNome= mvarsNome
End Property

Public Property Let sTelefone (ByVal vData As String)
    mvarsTelefone = vData
End Property

Public Property Get sTelefone () AS String
    sTelefone= mvarsTelefone
End Property
