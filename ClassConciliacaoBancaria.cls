VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassConciliacaoBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodConta As Variant
Private mvarlSequencialMovto As Long
Private mvariNumExtrato As Integer
Private mvarlSeqExtrBco As Long
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lSeqExtrBco(ByVal vData As Long)
    mvarlSeqExtrBco = vData
End Property

Public Property Get lSeqExtrBco() As Long
    lSeqExtrBco = mvarlSeqExtrBco
End Property

Public Property Let iNumExtrato(ByVal vData As Integer)
    mvariNumExtrato = vData
End Property

Public Property Get iNumExtrato() As Integer
    iNumExtrato = mvariNumExtrato
End Property

Public Property Let lSequencialMovto(ByVal vData As Long)
    mvarlSequencialMovto = vData
End Property

Public Property Get lSequencialMovto() As Long
    lSequencialMovto = mvarlSequencialMovto
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
        iCodConta = mvariCodConta
End Property



