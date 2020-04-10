VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNatMovCtaAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNatureza As String
Private mvarlCliForn As Long
Private mvarsNomeCliForn As String
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvardValor As Double
Private mvariTipoCliForn As Integer

Public Property Let iTipoCliForn(ByVal vData As Integer)
    mvariTipoCliForn = vData
End Property
Public Property Get iTipoCliForn() As Integer
    iTipoCliForn = mvariTipoCliForn
End Property

Public Property Let lCliForn(ByVal vData As Long)
    mvarlCliForn = vData
End Property
Public Property Get lCliForn() As Long
    lCliForn = mvarlCliForn
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property
Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property
Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property
Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let sNomeCliForn(ByVal vData As String)
    mvarsNomeCliForn = vData
End Property
Public Property Get sNomeCliForn() As String
    sNomeCliForn = mvarsNomeCliForn
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property
Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

