VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseOpcaoOrdenacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsOpcao As String
Private mvariIndice As Integer
Private mvarsSelecaoSQL1 As String
Private mvarsSelecaoSQL1Usuario As String
Private mvarsNomeIndice As String

Public Property Let sNomeIndice(ByVal vData As String)
    mvarsNomeIndice = vData
End Property

Public Property Get sNomeIndice() As String
    sNomeIndice = mvarsNomeIndice
End Property

Public Property Let sSelecaoSQL1Usuario(ByVal vData As String)
    mvarsSelecaoSQL1Usuario = vData
End Property

Public Property Get sSelecaoSQL1Usuario() As String
    sSelecaoSQL1Usuario = mvarsSelecaoSQL1Usuario
End Property

Public Property Let sSelecaoSQL1(ByVal vData As String)
    mvarsSelecaoSQL1 = vData
End Property

Public Property Get sSelecaoSQL1() As String
    sSelecaoSQL1 = mvarsSelecaoSQL1
End Property

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let sOpcao(ByVal vData As String)
    mvarsOpcao = vData
End Property

Public Property Get sOpcao() As String
    sOpcao = mvarsOpcao
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Private Sub Class_Initialize()

    mvarsNomeTela = String(STRING_NOME_TELA, 0)

End Sub

