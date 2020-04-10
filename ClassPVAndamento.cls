VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVAndamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariAuto As Integer
Private mvariFatorAuto As Integer
Private mvariOrd As Integer

Public Property Let iOrd(ByVal vData As Integer)
    mvariOrd = vData
End Property

Public Property Get iOrd() As Integer
    iOrd = mvariOrd
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iAuto(ByVal vData As Integer)
    mvariAuto = vData
End Property

Public Property Get iAuto() As Integer
    iAuto = mvariAuto
End Property

Public Property Let iFatorAuto(ByVal vData As Integer)
    mvariFatorAuto = vData
End Property

Public Property Get iFatorAuto() As Integer
    iFatorAuto = mvariFatorAuto
End Property
