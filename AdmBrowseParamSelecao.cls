VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseParamSelecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvariOrdem As Integer
Private mvarsProjeto As String
Private mvarsClasse As String
Private mvarsProperty As String

Public Property Let sProperty(ByVal vData As String)
    mvarsProperty = vData
End Property

Public Property Get sProperty() As String
    sProperty = mvarsProperty
End Property

Public Property Let sClasse(ByVal vData As String)
    mvarsClasse = vData
End Property

Public Property Get sClasse() As String
    sClasse = mvarsClasse
End Property

Public Property Let sProjeto(ByVal vData As String)
    mvarsProjeto = vData
End Property

Public Property Get sProjeto() As String
    sProjeto = mvarsProjeto
End Property

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property
