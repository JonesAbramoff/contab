VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCTMaqProgTurnoItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocCTMaqPTurno As Long
Private mvariTurno As Integer
Private mvardHoras As Double

Public Property Let lNumIntDocCTMaqPTurno(ByVal vData As Long)
    mvarlNumIntDocCTMaqPTurno = vData
End Property

Public Property Get lNumIntDocCTMaqPTurno() As Long
    lNumIntDocCTMaqPTurno = mvarlNumIntDocCTMaqPTurno
End Property

Public Property Let dHoras(ByVal vData As Double)
    mvardHoras = vData
End Property

Public Property Get dHoras() As Double
    dHoras = mvardHoras
End Property

Public Property Let iTurno(ByVal vData As Integer)
    mvariTurno = vData
End Property

Public Property Get iTurno() As Integer
    iTurno = mvariTurno
End Property
