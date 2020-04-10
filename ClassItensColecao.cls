VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensColecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCor As Integer
Private mvariVariacao As Integer
Private mvariPintura As Integer
Private mvariSeq As Integer
Private mvarlColecao As Long

Public Property Let lColecao(ByVal vData As Long)
    mvarlColecao = vData
End Property

Public Property Get lColecao() As Long
    lColecao = mvarlColecao
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iPintura(ByVal vData As Integer)
    mvariPintura = vData
End Property

Public Property Get iPintura() As Integer
    iPintura = mvariPintura
End Property

Public Property Let iCor(ByVal vData As Integer)
    mvariCor = vData
End Property

Public Property Get iCor() As Integer
    iCor = mvariCor
End Property

Public Property Let iVariacao(ByVal vData As Integer)
    mvariVariacao = vData
End Property

Public Property Get iVariacao() As Integer
    iVariacao = mvariVariacao
End Property

