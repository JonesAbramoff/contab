VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClientesLimites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodCliente As Long
Private mvariCodFilial As Integer
Private mvarsSerie As String
Private mvariTipoVersao As Integer
Private mvarsSenha As String
Private mvardtDataSenha As Date
Private mvardtValidadeAte As Date
Private mvarsVersao As String
Private mvariLimiteLogs As Integer
Private mvariLimiteFiliais As Integer
Private mvariLimiteEmpresas As Integer
Private mvarcolSiglasModulosLib As New Collection
Private mvarsCGC As String
Public Property Let sCGC(ByVal vData As String)
    mvarsCGC = vData
End Property
Public Property Get sCGC() As String
    sCGC = mvarsCGC
End Property
Public Property Set colSiglasModulosLib(ByVal vData As Collection)
    Set mvarcolSiglasModulosLib = vData
End Property
Public Property Get colSiglasModulosLib() As Collection
    Set colSiglasModulosLib = mvarcolSiglasModulosLib
End Property
Public Property Let iLimiteEmpresas(ByVal vData As Integer)
    mvariLimiteEmpresas = vData
End Property
Public Property Get iLimiteEmpresas() As Integer
    iLimiteEmpresas = mvariLimiteEmpresas
End Property
Public Property Let iLimiteFiliais(ByVal vData As Integer)
    mvariLimiteFiliais = vData
End Property
Public Property Get iLimiteFiliais() As Integer
    iLimiteFiliais = mvariLimiteFiliais
End Property
Public Property Let iLimiteLogs(ByVal vData As Integer)
    mvariLimiteLogs = vData
End Property
Public Property Get iLimiteLogs() As Integer
    iLimiteLogs = mvariLimiteLogs
End Property
Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property
Public Property Get sVersao() As String
    sVersao = mvarsVersao
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
Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property
Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property
Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property
Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property
