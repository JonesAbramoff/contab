VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsSigla As String
Private mvarsDescricao As String
Private mvariVinculadoNF As Integer

Public Property Let iVinculadoNF(ByVal vData As Integer)
    mvariVinculadoNF = vData
End Property

Public Property Get iVinculadoNF() As Integer
    iVinculadoNF = mvariVinculadoNF
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property



