VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDicConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariID As Integer
Private mvariLimiteFiliais As Integer
Private mvariLimiteLogs As Integer
Private mvarsSerie As String
Private mvariTipoVersao As Integer
Private mvariLimiteEmpresas As Integer
Private mvardtValidade As Date
Private mvarsSenha As String
Private mvardtDataSenha As Date
Private mvardtValidadeAte As Date
Private mvarcolModulosLib As New Collection
Public Property Set colModulosLib(ByVal vData As Collection)
    Set mvarcolModulosLib = vData
End Property

Public Property Get colModulosLib() As Collection
    Set colModulosLib = mvarcolModulosLib
End Property

Public Property Let dtValidadeAte(ByVal vData As Date)
    mvardtValidadeAte = vData
End Property

Public Property Get dtValidadeAte() As Date
    dtValidadeAte = mvardtValidadeAte
End Property

Public Property Let dtDataSenha(ByVal vData As Date)
    mvardtDataSenha = vData
End Property

Public Property Get dtDataSenha() As Date
    dtDataSenha = mvardtDataSenha
End Property

Public Property Let sSenha(ByVal vData As String)
    mvarsSenha = vData
End Property

Public Property Get sSenha() As String
    sSenha = mvarsSenha
End Property

Public Property Let dtValidadeDe(ByVal vData As Date)
    mvardtValidade = vData
End Property
Public Property Get dtValidadeDe() As Date
    dtValidadeDe = mvardtValidade
End Property

Public Property Let iLimiteEmpresas(ByVal vData As Integer)
    mvariLimiteEmpresas = vData
End Property
Public Property Get iLimiteEmpresas() As Integer
    iLimiteEmpresas = mvariLimiteEmpresas
End Property

Public Property Let iTipoVersao(ByVal vData As Integer)
    mvariTipoVersao = vData
End Property

Public Property Get iTipoVersao() As Integer
    iTipoVersao = mvariTipoVersao
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property
Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iLimiteLogs(ByVal vData As Integer)
    mvariLimiteLogs = vData
End Property
Public Property Get iLimiteLogs() As Integer
    iLimiteLogs = mvariLimiteLogs
End Property

Public Property Let iLimiteFiliais(ByVal vData As Integer)
    mvariLimiteFiliais = vData
End Property
Public Property Get iLimiteFiliais() As Integer
    iLimiteFiliais = mvariLimiteFiliais
End Property

Public Property Let iID(ByVal vData As Integer)
    mvariID = vData
End Property
Public Property Get iID() As Integer
    iID = mvariID
End Property
