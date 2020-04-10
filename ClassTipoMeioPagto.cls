VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoMeioPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvariExigeNumero As Integer
Private mvariInativo As Integer

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let iExigeNumero(ByVal vData As Integer)
    mvariExigeNumero = vData
End Property

Public Property Get iExigeNumero() As Integer
    iExigeNumero = mvariExigeNumero
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property



