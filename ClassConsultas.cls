VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvarsSigla As String
Private mvarsConsulta As String
Private mvariPosicao As Integer
Private mvariNivel As Integer
Private mvarsTelaRelacionada As String
Private mvariIconeModulo As Integer
Private mvariIconeConsulta As Integer
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iIconeConsulta(ByVal vData As Integer)
    mvariIconeConsulta = vData
End Property

Public Property Get iIconeConsulta() As Integer
    iIconeConsulta = mvariIconeConsulta
End Property

Public Property Let iIconeModulo(ByVal vData As Integer)
    mvariIconeModulo = vData
End Property

Public Property Get iIconeModulo() As Integer
    iIconeModulo = mvariIconeModulo
End Property

Public Property Let sTelaRelacionada(ByVal vData As String)
    mvarsTelaRelacionada = vData
End Property

Public Property Get sTelaRelacionada() As String
    sTelaRelacionada = mvarsTelaRelacionada
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let iPosicao(ByVal vData As Integer)
    mvariPosicao = vData
End Property

Public Property Get iPosicao() As Integer
    iPosicao = mvariPosicao
End Property

Public Property Let sConsulta(ByVal vData As String)
    mvarsConsulta = vData
End Property

Public Property Get sConsulta() As String
    sConsulta = mvarsConsulta
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property
