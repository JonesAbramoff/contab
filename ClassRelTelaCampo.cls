VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRelTelaCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColRelTelaCampos"
Option Explicit

Private mvarsCodRel As String
Private mvariSequencial As Integer
Private mvarsNome As String
Private mvariTipo As Integer
Private mvariTamanho As Integer

Public Property Let iTamanho(ByVal vData As Integer)
    mvariTamanho = vData
End Property

Public Property Get iTamanho() As Integer
    iTamanho = mvariTamanho
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let sCodRel(ByVal vData As String)
    mvarsCodRel = vData
End Property

Public Property Get sCodRel() As String
    sCodRel = mvarsCodRel
End Property



