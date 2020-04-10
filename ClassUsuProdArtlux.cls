VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUsuProdArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariFilialEmpresa As Integer
Private mvarsCodUsuario As String
Private mvariAcessoCorte As Integer
Private mvariAcessoForro As Integer
Private mvariAcessoMontagem As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let iAcessoCorte(ByVal vData As Integer)
    mvariAcessoCorte = vData
End Property

Public Property Get iAcessoCorte() As Integer
    iAcessoCorte = mvariAcessoCorte
End Property

Public Property Let iAcessoForro(ByVal vData As Integer)
    mvariAcessoForro = vData
End Property

Public Property Get iAcessoForro() As Integer
    iAcessoForro = mvariAcessoForro
End Property

Public Property Let iAcessoMontagem(ByVal vData As Integer)
    mvariAcessoMontagem = vData
End Property

Public Property Get iAcessoMontagem() As Integer
    iAcessoMontagem = mvariAcessoMontagem
End Property
