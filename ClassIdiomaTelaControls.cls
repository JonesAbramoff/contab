VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIdiomaTelaControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsNomeControle As String
Private mvarsNomeTelaExibicao As String
Private mvarsNomeControleExibicao As String
Private mvarsNomeTabela As String
Private mvarsNomeCampo As String
Private mvariComMaxLen As Integer
Private mvariComMultiLine As Integer
Private mvariEmGrid As Integer

Public Property Let iEmGrid(ByVal vData As Integer)
    mvariEmGrid = vData
End Property

Public Property Get iEmGrid() As Integer
    iEmGrid = mvariEmGrid
End Property

Public Property Let iComMultiLine(ByVal vData As Integer)
    mvariComMultiLine = vData
End Property

Public Property Get iComMultiLine() As Integer
    iComMultiLine = mvariComMultiLine
End Property

Public Property Let iComMaxLen(ByVal vData As Integer)
    mvariComMaxLen = vData
End Property

Public Property Get iComMaxLen() As Integer
    iComMaxLen = mvariComMaxLen
End Property

Public Property Let sNomeTabela(ByVal vData As String)
    mvarsNomeTabela = vData
End Property

Public Property Get sNomeTabela() As String
    sNomeTabela = mvarsNomeTabela
End Property

Public Property Let sNomeCampo(ByVal vData As String)
    mvarsNomeCampo = vData
End Property

Public Property Get sNomeCampo() As String
    sNomeCampo = mvarsNomeCampo
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sNomeControle(ByVal vData As String)
    mvarsNomeControle = vData
End Property

Public Property Get sNomeControle() As String
    sNomeControle = mvarsNomeControle
End Property

Public Property Let sNomeTelaExibicao(ByVal vData As String)
    mvarsNomeTelaExibicao = vData
End Property

Public Property Get sNomeTelaExibicao() As String
    sNomeTelaExibicao = mvarsNomeTelaExibicao
End Property

Public Property Let sNomeControleExibicao(ByVal vData As String)
    mvarsNomeControleExibicao = vData
End Property

Public Property Get sNomeControleExibicao() As String
    sNomeControleExibicao = mvarsNomeControleExibicao
End Property

