VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMnemonicoCTBValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsMnemonico As String
Private mvarsDescricao As String
Private mvarsValor As String

Public Property Let sValor(ByVal vData As String)
Attribute sValor.VB_Description = "Valor do Mnemonico"
    mvarsValor = vData
End Property

Public Property Get sValor() As String
    sValor = mvarsValor
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descricao do Mnemonico"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sMnemonico(ByVal vData As String)
Attribute sMnemonico.VB_Description = "Mnemonico"
    mvarsMnemonico = vData
End Property

Public Property Get sMnemonico() As String
    sMnemonico = mvarsMnemonico
End Property



