VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasosTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOcrCaso As Long
Private mvariTipoPai As Integer
Private mvariSeqPai As Integer
Private mvariSeq As Integer
Private mvarsTexto As String

Public Property Let lNumIntDocOcrCaso(ByVal vData As Long)
    mvarlNumIntDocOcrCaso = vData
End Property

Public Property Get lNumIntDocOcrCaso() As Long
    lNumIntDocOcrCaso = mvarlNumIntDocOcrCaso
End Property

Public Property Let iTipoPai(ByVal vData As Integer)
    mvariTipoPai = vData
End Property

Public Property Get iTipoPai() As Integer
    iTipoPai = mvariTipoPai
End Property

Public Property Let iSeqPai(ByVal vData As Integer)
    mvariSeqPai = vData
End Property

Public Property Get iSeqPai() As Integer
    iSeqPai = mvariSeqPai
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sTexto(ByVal vData As String)
    mvarsTexto = vData
End Property

Public Property Get sTexto() As String
    sTexto = mvarsTexto
End Property
