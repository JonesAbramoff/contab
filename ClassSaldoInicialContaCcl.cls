VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSaldoInicialContaCcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsConta As String
Private mvarsCcl As String
Private mvardSldIni As Double
Private mvarsDescCcl As String

Public Property Let sDescCcl(ByVal vData As String)
    mvarsDescCcl = vData
End Property

Public Property Get sDescCcl() As String
    sDescCcl = mvarsDescCcl
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

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



