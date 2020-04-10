VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIntegracaoLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariTipoArq As Integer
Private mvariSeq As Integer
Private mvarsCritica As String
Private mvariAviso As Integer
Private mvariEtapa As Integer
Private mvarsChave As String

Public Property Let sChave(ByVal vData As String)
    mvarsChave = vData
End Property

Public Property Get sChave() As String
    sChave = mvarsChave
End Property

Public Property Let iEtapa(ByVal vData As Integer)
    mvariEtapa = vData
End Property

Public Property Get iEtapa() As Integer
    iEtapa = mvariEtapa
End Property

Public Property Let iAviso(ByVal vData As Integer)
    mvariAviso = vData
End Property

Public Property Get iAviso() As Integer
    iAviso = mvariAviso
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipoArq(ByVal vData As Integer)
    mvariTipoArq = vData
End Property

Public Property Get iTipoArq() As Integer
    iTipoArq = mvariTipoArq
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sCritica(ByVal vData As String)
    mvarsCritica = vData
End Property

Public Property Get sCritica() As String
    sCritica = mvarsCritica
End Property
