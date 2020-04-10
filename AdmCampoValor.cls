VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmCampoValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Nome do campo, valor do campo."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCampo As String
Private mvarvValor As Variant
Private mvariTamanhoString As Integer

Public Property Let iTamanhoString(ByVal vData As Integer)
Attribute iTamanhoString.VB_Description = "Tamanho de string no BD."
    mvariTamanhoString = vData
End Property

Public Property Get iTamanhoString() As Integer
    iTamanhoString = mvariTamanhoString
End Property

Public Property Let vValor(ByVal vData As Variant)
Attribute vValor.VB_Description = "Valor do Campo."
    mvarvValor = vData
End Property

Public Property Set vValor(ByVal vData As Object)
    Set mvarvValor = vData
End Property

Public Property Get vValor() As Variant
    If IsObject(mvarvValor) Then
        Set vValor = mvarvValor
    Else
        vValor = mvarvValor
    End If
End Property

Public Property Let sCampo(ByVal vData As String)
Attribute sCampo.VB_Description = "Nome do campo."
    mvarsCampo = vData
End Property

Public Property Get sCampo() As String
    sCampo = mvarsCampo
End Property



