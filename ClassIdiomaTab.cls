VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIdiomaTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTabela As String
Private mvariChaveDocI As Integer
Private mvarlChaveDocL As Long
Private mvarsChaveDocS As String
Private mvarcolCampos As New Collection
Private mvariLinhaGrid As Integer

Public Property Let iLinhaGrid(ByVal vData As Integer)
    mvariLinhaGrid = vData
End Property

Public Property Get iLinhaGrid() As Integer
    iLinhaGrid = mvariLinhaGrid
End Property

Public Property Set colCampos(ByVal vData As Collection)
    Set mvarcolCampos = vData
End Property

Public Property Get colCampos() As Collection
    Set colCampos = mvarcolCampos
End Property

Public Property Let sChaveDocS(ByVal vData As String)
    mvarsChaveDocS = vData
End Property

Public Property Get sChaveDocS() As String
    sChaveDocS = mvarsChaveDocS
End Property

Public Property Let lChaveDocL(ByVal vData As Long)
    mvarlChaveDocL = vData
End Property

Public Property Get lChaveDocL() As Long
    lChaveDocL = mvarlChaveDocL
End Property

Public Property Let iChaveDocI(ByVal vData As Integer)
    mvariChaveDocI = vData
End Property

Public Property Get iChaveDocI() As Integer
    iChaveDocI = mvariChaveDocI
End Property

Public Property Let sNomeTabela(ByVal vData As String)
    mvarsNomeTabela = vData
End Property

Public Property Get sNomeTabela() As String
    sNomeTabela = mvarsNomeTabela
End Property

Public Function Reiniciar() As Long
Dim objCampo As ClassIdiomaTabCampo
    mvariChaveDocI = 0
    mvarlChaveDocL = 0
    mvarsChaveDocS = ""
    For Each objCampo In mvarcolCampos
        Call objCampo.Reiniciar
    Next
End Function

Public Function Associar(ByVal sChaveDocSAux As String, ByVal lChaveDocLAux As Long, ByVal iChaveDocIAux As Integer, Optional bTrazerDados As Boolean = False) As Long
    mvariChaveDocI = iChaveDocIAux
    mvarlChaveDocL = lChaveDocLAux
    mvarsChaveDocS = sChaveDocSAux
    If bTrazerDados Then
        'LÊ AS INFORMAÇÕES SALVAS
    End If
End Function
