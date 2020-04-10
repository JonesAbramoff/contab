VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmBrowsePesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsSelecao As String
Private mvariTelaOk As Integer

Public Property Let iTelaOk(ByVal vData As Integer)
    mvariTelaOk = vData
End Property

Public Property Get iTelaOk() As Integer
    iTelaOk = mvariTelaOk
End Property

Public Property Let sSelecao(ByVal vData As String)
    mvarsSelecao = vData
End Property

Public Property Get sSelecao() As String
    sSelecao = mvarsSelecao
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

