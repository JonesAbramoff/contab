VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoteiroSRVOperMaq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOper As Long
Private mvariSeq As Integer
Private mvariCodMaq As Integer
Private mvardHoras As Double
Private mvariQtd As Integer
Private mvarsObs As String
Private mvariFilialEmpMaq As Integer

Public Property Let iFilialEmpMaq(ByVal vData As Integer)
    mvariFilialEmpMaq = vData
End Property

Public Property Get iFilialEmpMaq() As Integer
    iFilialEmpMaq = mvariFilialEmpMaq
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocOper(ByVal vData As Long)
    mvarlNumIntDocOper = vData
End Property

Public Property Get lNumIntDocOper() As Long
    lNumIntDocOper = mvarlNumIntDocOper
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iCodMaq(ByVal vData As Integer)
    mvariCodMaq = vData
End Property

Public Property Get iCodMaq() As Integer
    iCodMaq = mvariCodMaq
End Property

Public Property Let dHoras(ByVal vData As Double)
    mvardHoras = vData
End Property

Public Property Get dHoras() As Double
    dHoras = mvardHoras
End Property

Public Property Let iQtd(ByVal vData As Integer)
    mvariQtd = vData
End Property

Public Property Get iQtd() As Integer
    iQtd = mvariQtd
End Property

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property
