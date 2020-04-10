VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaArqFisAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodECF As Integer
Private mvariFilialEmpresa As Integer
Private mvardtData As Date
Private mvarsSituacaoTrib As String
Private mvardTotalizador As Double

Public Property Let dTotalizador(ByVal vData As Double)
    mvardTotalizador = vData
End Property

Public Property Get dTotalizador() As Double
    dTotalizador = mvardTotalizador
End Property

Public Property Let sSituacaoTrib(ByVal vData As String)
    mvarsSituacaoTrib = vData
End Property

Public Property Get sSituacaoTrib() As String
    sSituacaoTrib = mvarsSituacaoTrib
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodECF(ByVal vData As Integer)
    mvariCodECF = vData
End Property

Public Property Get iCodECF() As Integer
    iCodECF = mvariCodECF
End Property





