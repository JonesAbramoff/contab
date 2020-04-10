VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmValorCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "armazena o valor do campo"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipo As Integer
Private mvariPosicaoTela As Integer
Private mvariSubTipo As Integer
Private mvariAlinhamento As Integer
Private mvariTamanho As Integer

Public Property Let iTamanho(ByVal vData As Integer)
    mvariTamanho = vData
End Property

Public Property Get iTamanho() As Integer
    iTamanho = mvariTamanho
End Property

Public Property Let iAlinhamento(ByVal vData As Integer)
    mvariAlinhamento = vData
End Property

Public Property Get iAlinhamento() As Integer
    iAlinhamento = mvariAlinhamento
End Property

Public Property Let iSubTipo(ByVal vData As Integer)
    mvariSubTipo = vData
End Property

Public Property Get iSubTipo() As Integer
    iSubTipo = mvariSubTipo
End Property

Public Property Let iPosicaoTela(ByVal vData As Integer)
    mvariPosicaoTela = vData
End Property

Public Property Get iPosicaoTela() As Integer
    iPosicaoTela = mvariPosicaoTela
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property



