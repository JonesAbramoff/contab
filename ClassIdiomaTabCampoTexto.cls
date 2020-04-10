VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIdiomaTabCampoTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariIdioma As Integer
Private mvarsTexto As String

'==========AUX=============
Private mvarsNomeTabela As String
Private mvariChaveDocI As Integer
Private mvarlChaveDocL As Long
Private mvarsChaveDocS As String
Private mvarsNomeCampo As String
Private mvariSeq As Integer
'==========AUX=============

Public Property Let iIdioma(ByVal vData As Integer)
    mvariIdioma = vData
End Property

Public Property Get iIdioma() As Integer
    iIdioma = mvariIdioma
End Property

Public Property Let sTexto(ByVal vData As String)
    mvarsTexto = vData
End Property

Public Property Get sTexto() As String
    sTexto = mvarsTexto
End Property

'==========AUX=============
Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sNomeCampo(ByVal vData As String)
    mvarsNomeCampo = vData
End Property

Public Property Get sNomeCampo() As String
    sNomeCampo = mvarsNomeCampo
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
'==========AUX=============

Public Function Reiniciar() As Long
    mvariChaveDocI = 0
    mvarlChaveDocL = 0
    mvarsChaveDocS = ""
    mvarsTexto = ""
End Function
