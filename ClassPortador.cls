VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPortador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariBanco As Integer
Private mvariInativo As Integer
Private mvariCodigo As Integer
Private mvarsNome As String
Private mvarsNomeReduzido As String

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property
