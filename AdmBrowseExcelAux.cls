VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseExcelAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCampo As String
Private mvariFormula As Integer
Private mvariPosicao As Integer
Private mvarobjBrowseUsuarioCampo As New AdmBrowseUsuarioCampo
Private mvariSeq As Integer

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Set objBrowseUsuarioCampo(ByVal vData As Object)
    Set mvarobjBrowseUsuarioCampo = vData
End Property

Public Property Get objBrowseUsuarioCampo() As AdmBrowseUsuarioCampo
    Set objBrowseUsuarioCampo = mvarobjBrowseUsuarioCampo
End Property

Public Property Let iPosicao(ByVal vData As Integer)
    mvariPosicao = vData
End Property

Public Property Get iPosicao() As Integer
    iPosicao = mvariPosicao
End Property

Public Property Let iFormula(ByVal vData As Integer)
    mvariFormula = vData
End Property

Public Property Get iFormula() As Integer
    iFormula = mvariFormula
End Property

Public Property Let sCampo(ByVal vData As String)
    mvarsCampo = vData
End Property

Public Property Get sCampo() As String
    sCampo = mvarsCampo
End Property
