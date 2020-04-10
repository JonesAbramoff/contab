VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntSped As Long
Private mvarlSeq As Long
Private mvariFilialEmpresa As Integer
Private mvarsTipo As String
Private mvarsRegistro As String
Private mvarlSeqPai As Long

Public Property Let lSeqPai(ByVal vData As Long)
    mvarlSeqPai = vData
End Property

Public Property Get lSeqPai() As Long
    lSeqPai = mvarlSeqPai
End Property

Public Property Let sRegistro(ByVal vData As String)
    mvarsRegistro = vData
End Property

Public Property Get sRegistro() As String
    sRegistro = mvarsRegistro
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lSeq(ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Let lNumIntSped(ByVal vData As Long)
    mvarlNumIntSped = vData
End Property

Public Property Get lNumIntSped() As Long
    lNumIntSped = mvarlNumIntSped
End Property






