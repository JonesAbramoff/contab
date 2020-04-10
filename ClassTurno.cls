VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarlNumIntDocCT As Long
Private mvarlNumIntDocMaq As Long
Private mvariTurno As Integer
Private mvardQtdHoras(1 To 7) As Double

Public Property Let dQtdHoras(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQtdHoras(iIndice) = vData
End Property

Public Property Get dQtdHoras(ByVal iIndice As Integer) As Double
    dQtdHoras = mvardQtdHoras(iIndice)
End Property

Public Property Let iTurno(ByVal vData As Integer)
    mvariTurno = vData
End Property

Public Property Get iTurno() As Integer
    iTurno = mvariTurno
End Property

Public Property Let lNumIntDocMaq(ByVal vData As Long)
    mvarlNumIntDocMaq = vData
End Property

Public Property Get lNumIntDocMaq() As Long
    lNumIntDocMaq = mvarlNumIntDocMaq
End Property

Public Property Let lNumIntDocCT(ByVal vData As Long)
    mvarlNumIntDocCT = vData
End Property

Public Property Get lNumIntDocCT() As Long
    lNumIntDocCT = mvarlNumIntDocCT
End Property



