VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBloqueioPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlPedidoDeVendas As Long
Private mvariSequencial As Integer
Private mvariTipoDeBloqueio As Integer
Private mvarsCodUsuario As String
Private mvarsResponsavel As String
Private mvardtData As Date
Private mvarsCodUsuarioLib As String
Private mvarsResponsavelLib As String
Private mvardtDataLib As Date
Private mvarsObservacao As String

Public Property Let dtDataLib(ByVal vData As Date)
    mvardtDataLib = vData
End Property

Public Property Get dtDataLib() As Date
    dtDataLib = mvardtDataLib
End Property

Public Property Let sResponsavelLib(ByVal vData As String)
    mvarsResponsavelLib = vData
End Property

Public Property Get sResponsavelLib() As String
    sResponsavelLib = mvarsResponsavelLib
End Property

Public Property Let sCodUsuarioLib(ByVal vData As String)
    mvarsCodUsuarioLib = vData
End Property

Public Property Get sCodUsuarioLib() As String
    sCodUsuarioLib = mvarsCodUsuarioLib
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let iTipoDeBloqueio(ByVal vData As Integer)
    mvariTipoDeBloqueio = vData
End Property

Public Property Get iTipoDeBloqueio() As Integer
    iTipoDeBloqueio = mvariTipoDeBloqueio
End Property

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let lPedidoDeVendas(ByVal vData As Long)
    mvarlPedidoDeVendas = vData
End Property

Public Property Get lPedidoDeVendas() As Long
    lPedidoDeVendas = mvarlPedidoDeVendas
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property


