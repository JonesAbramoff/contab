VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequeBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntCheque As Long
Private mvariFilialEmpresa As Integer
Private mvarlNumBordero As Long
Private mvariTipoBordero As Integer

Public Property Let iTipoBordero(ByVal vData As Integer)
    mvariTipoBordero = vData
End Property

Public Property Get iTipoBordero() As Integer
    iTipoBordero = mvariTipoBordero
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
        iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntCheque(ByVal vData As Long)
    mvarlNumIntCheque = vData
End Property

Public Property Get lNumIntCheque() As Long
    lNumIntCheque = mvarlNumIntCheque
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



