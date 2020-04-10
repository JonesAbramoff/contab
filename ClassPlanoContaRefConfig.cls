VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanoContaRefConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigoModelo As Long
Private mvarsConta As String
Private mvariSeq As Integer
Private mvarsContaInicial As String
Private mvarsContaFinal As String
Private mvarsCclInicial As String
Private mvarsCclFinal As String
Private mvariSubtrai As Integer

Public Property Let lCodigoModelo(ByVal vData As Long)
    mvarlCodigoModelo = vData
End Property

Public Property Get lCodigoModelo() As Long
    lCodigoModelo = mvarlCodigoModelo
End Property

Public Property Let sConta(ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sContaInicial(ByVal vData As String)
    mvarsContaInicial = vData
End Property

Public Property Get sContaInicial() As String
    sContaInicial = mvarsContaInicial
End Property

Public Property Let sContaFinal(ByVal vData As String)
    mvarsContaFinal = vData
End Property

Public Property Get sContaFinal() As String
    sContaFinal = mvarsContaFinal
End Property

Public Property Let sCclInicial(ByVal vData As String)
    mvarsCclInicial = vData
End Property

Public Property Get sCclInicial() As String
    sCclInicial = mvarsCclInicial
End Property

Public Property Let sCclFinal(ByVal vData As String)
    mvarsCclFinal = vData
End Property

Public Property Get sCclFinal() As String
    sCclFinal = mvarsCclFinal
End Property

Public Property Let iSubtrai(ByVal vData As Integer)
    mvariSubtrai = vData
End Property

Public Property Get iSubtrai() As Integer
    iSubtrai = mvariSubtrai
End Property
