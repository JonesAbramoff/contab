VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJDadosCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataIniPrev As Date
Private mvardtDataFimPrev As Date
Private mvardtDataIniReal As Date
Private mvardtDataFimReal As Date
Private mvardPercentualComplet As Double
Private mvariStatus As Integer

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dPercentualComplet(ByVal vData As Double)
    mvardPercentualComplet = vData
End Property

Public Property Get dPercentualComplet() As Double
    dPercentualComplet = mvardPercentualComplet
End Property

Public Property Let dtDataIniReal(ByVal vData As Date)
    mvardtDataIniReal = vData
End Property

Public Property Get dtDataIniReal() As Date
    dtDataIniReal = mvardtDataIniReal
End Property

Public Property Let dtDataFimReal(ByVal vData As Date)
    mvardtDataFimReal = vData
End Property

Public Property Get dtDataFimReal() As Date
    dtDataFimReal = mvardtDataFimReal
End Property

Public Property Let dtDataIniPrev(ByVal vData As Date)
    mvardtDataIniPrev = vData
End Property

Public Property Get dtDataIniPrev() As Date
    dtDataIniPrev = mvardtDataIniPrev
End Property

Public Property Let dtDataFimPrev(ByVal vData As Date)
    mvardtDataFimPrev = vData
End Property

Public Property Get dtDataFimPrev() As Date
    dtDataFimPrev = mvardtDataFimPrev
End Property

Private Sub Class_Initialize()

    mvardtDataIniPrev = DATA_NULA
    mvardtDataFimPrev = DATA_NULA
    mvardtDataIniReal = DATA_NULA
    mvardtDataFimReal = DATA_NULA
    
End Sub
