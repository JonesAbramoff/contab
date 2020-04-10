VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDASAliquotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariAno As Integer
Private mvariMes As Integer
Private mvardAliquotaICMS As Double
Private mvardAliquotaTotal As Double
Private mvardAliquotaICMSServ As Double

Public iFilialEmpresa As Integer

Public Property Let dAliquotaICMSServ(ByVal vData As Double)
    mvardAliquotaICMSServ = vData
End Property

Public Property Get dAliquotaICMSServ() As Double
    dAliquotaICMSServ = mvardAliquotaICMSServ
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dAliquotaTotal(ByVal vData As Double)
    mvardAliquotaTotal = vData
End Property

Public Property Get dAliquotaTotal() As Double
    dAliquotaTotal = mvardAliquotaTotal
End Property
