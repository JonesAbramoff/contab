VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvarsCodUsuario As String
Private mvarlTopo As Long
Private mvarlEsquerda As Long
Private mvarlLargura As Long
Private mvarlAltura As Long

Public Property Let lAltura(ByVal vData As Long)
    mvarlAltura = vData
End Property

Public Property Get lAltura() As Long
    lAltura = mvarlAltura
End Property

Public Property Let lLargura(ByVal vData As Long)
    mvarlLargura = vData
End Property

Public Property Get lLargura() As Long
    lLargura = mvarlLargura
End Property

Public Property Let lEsquerda(ByVal vData As Long)
    mvarlEsquerda = vData
End Property

Public Property Get lEsquerda() As Long
    lEsquerda = mvarlEsquerda
End Property

Public Property Let lTopo(ByVal vData As Long)
    mvarlTopo = vData
End Property

Public Property Get lTopo() As Long
    lTopo = mvarlTopo
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property



