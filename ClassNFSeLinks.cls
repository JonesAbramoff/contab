VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFSeLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodIBGE As String
Private mvarsLinkConsulta As String
Private mvarsParamConsulta As String
Private mvarsLinkVerificacao As String
Private mvarsLinkSite As String
Private mvarsEmail As String

'AUXILIAR
Private mvariFilialEmpresa As Integer
Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodIBGE(ByVal vData As String)
    mvarsCodIBGE = vData
End Property

Public Property Get sCodIBGE() As String
    sCodIBGE = mvarsCodIBGE
End Property

Public Property Let sLinkConsulta(ByVal vData As String)
    mvarsLinkConsulta = vData
End Property

Public Property Get sLinkConsulta() As String
    sLinkConsulta = mvarsLinkConsulta
End Property

Public Property Let sParamConsulta(ByVal vData As String)
    mvarsParamConsulta = vData
End Property

Public Property Get sParamConsulta() As String
    sParamConsulta = mvarsParamConsulta
End Property

Public Property Let sLinkVerificacao(ByVal vData As String)
    mvarsLinkVerificacao = vData
End Property

Public Property Get sLinkVerificacao() As String
    sLinkVerificacao = mvarsLinkVerificacao
End Property

Public Property Let sLinkSite(ByVal vData As String)
    mvarsLinkSite = vData
End Property

Public Property Get sLinkSite() As String
    sLinkSite = mvarsLinkSite
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property
