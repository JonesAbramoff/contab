VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProcReajTitRecCcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocProc As Long
Private mvarsCcl As String
Private mvardValorReajuste As Double

Public Property Let lNumIntDocProc(ByVal vData As Long)
    mvarlNumIntDocProc = vData
End Property

Public Property Get lNumIntDocProc() As Long
    lNumIntDocProc = mvarlNumIntDocProc
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let dValorReajuste(ByVal vData As Double)
    mvardValorReajuste = vData
End Property

Public Property Get dValorReajuste() As Double
    dValorReajuste = mvardValorReajuste
End Property
