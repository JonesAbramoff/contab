VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoClienteCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariTipoDeCliente As Integer
Private mvarsCategoria As String
Private mvarsItem As String

Public Property Let sItem(ByVal vData As String)
    mvarsItem = vData
End Property

Public Property Get sItem() As String
    sItem = mvarsItem
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Let iTipoDeCliente(ByVal vData As Integer)
    mvariTipoDeCliente = vData
End Property

Public Property Get iTipoDeCliente() As Integer
    iTipoDeCliente = mvariTipoDeCliente
End Property




