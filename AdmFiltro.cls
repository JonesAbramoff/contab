VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum enumOperador
    OP_IGUAL
    OP_MAIOR
    OP_MENOR
    OP_MAIOR_OU_IGUAL
    OP_MENOR_OU_IGUAL
    OP_DIFERENTE
End Enum

Private mvarsCampo As String
Private mvarvValor As Variant
Private mvareOperador As enumOperador

Public Property Let eOperador(ByVal vData As enumOperador)
    mvareOperador = vData
End Property

Public Property Get eOperador() As enumOperador
    eOperador = mvareOperador
End Property

Public Property Let vValor(ByVal vData As Variant)
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
    mvarsCampo = vData
End Property

Public Property Get sCampo() As String
    sCampo = mvarsCampo
End Property



