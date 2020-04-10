VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedContabil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjSpedContabil As ClassSpedContabil
Private mvarsConta As String
Private mvarsCcl As String
Private mvardSldIni As Double
Private mvardCred As Double
Private mvardDeb As Double
Private mvardtDataIni As Date
Private mvardtDataFim As Date
Private mvarsContaBD As String
Private mvariNaturezaSped As Integer

Public Property Let iNaturezaSped(ByVal vData As Integer)
    mvariNaturezaSped = vData
End Property

Public Property Get iNaturezaSped() As Integer
    iNaturezaSped = mvariNaturezaSped
End Property

Public Property Let sContaBD(ByVal vData As String)
    mvarsContaBD = vData
End Property

Public Property Get sContaBD() As String
    sContaBD = mvarsContaBD
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataIni(ByVal vData As Date)
    mvardtDataIni = vData
End Property

Public Property Get dtDataIni() As Date
    dtDataIni = mvardtDataIni
End Property

Public Property Set objSpedContabil(ByVal vData As ClassSpedContabil)
    Set mvarobjSpedContabil = vData
End Property

Public Property Get objSpedContabil() As ClassSpedContabil
    Set objSpedContabil = mvarobjSpedContabil
End Property

Public Property Let dDeb(ByVal vData As Double)
    mvardDeb = vData
End Property

Public Property Get dDeb() As Double
    dDeb = mvardDeb
End Property

Public Property Let dCred(ByVal vData As Double)
    mvardCred = vData
End Property

Public Property Get dCred() As Double
    dCred = mvardCred
End Property

Public Property Let dSldIni(ByVal vData As Double)
    mvardSldIni = vData
End Property

Public Property Get dSldIni() As Double
    dSldIni = mvardSldIni
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sConta(ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

