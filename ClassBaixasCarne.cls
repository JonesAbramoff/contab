VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixasCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvardtDatabaixa As Date
Private mvarlCodigo As Long
Private mvarcolParcelas As New Collection

Public Property Set colParcelas(ByVal vdata As Collection)
    Set mvarcolParcelas = vdata
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let lCodigo(ByVal vdata As Long)
    mvarlCodigo = vdata
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let lNumIntDoc(ByVal vdata As Long)
    mvarlNumIntDoc = vdata
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dtDataBaixa(ByVal vdata As Date)
    mvardtDatabaixa = vdata
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDatabaixa
End Property

Public Property Let iFilialEmpresa(ByVal vdata As Integer)
    mvariFilialEmpresa = vdata
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

