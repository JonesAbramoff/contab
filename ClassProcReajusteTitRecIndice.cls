VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProcReajTitRecIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocProc As Long
Private mvariMoeda As Integer
Private mvariPeriodicidade As Integer
Private mvardIndice As Double

Public Property Let lNumIntDocProc(ByVal vData As Long)
    mvarlNumIntDocProc = vData
End Property

Public Property Get lNumIntDocProc() As Long
    lNumIntDocProc = mvarlNumIntDocProc
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iPeriodicidade(ByVal vData As Integer)
    mvariPeriodicidade = vData
End Property

Public Property Get iPeriodicidade() As Integer
    iPeriodicidade = mvariPeriodicidade
End Property

Public Property Let dIndice(ByVal vData As Double)
    mvardIndice = vData
End Property

Public Property Get dIndice() As Double
    dIndice = mvardIndice
End Property
