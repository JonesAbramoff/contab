VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoLogAtu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntArq As Long
Private mvariTabela As Integer
Private mvarlNumRegArqTexto As Long
Private mvariSeq As Integer
Private mvarsCritica As String
Private mvariAviso As Integer

Public Property Let iAviso(ByVal vData As Integer)
    mvariAviso = vData
End Property

Public Property Get iAviso() As Integer
    iAviso = mvariAviso
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let iTabela(ByVal vData As Integer)
    mvariTabela = vData
End Property

Public Property Get iTabela() As Integer
    iTabela = mvariTabela
End Property

Public Property Let lNumRegArqTexto(ByVal vData As Long)
    mvarlNumRegArqTexto = vData
End Property

Public Property Get lNumRegArqTexto() As Long
    lNumRegArqTexto = mvarlNumRegArqTexto
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
