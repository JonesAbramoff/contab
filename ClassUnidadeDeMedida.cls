VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUnidadeDeMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsSigla As String
Private mvarsNome As String
Private mvariClasse As Integer
Private mvarsSiglaUMBase As String
Private mvardQuantidade As Double
Private mvariEditavel As Integer
Private mvardtDataLog As Date
Private mvariAtivo As Integer

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iEditavel(ByVal vData As Integer)
    mvariEditavel = vData
End Property

Public Property Get iEditavel() As Integer
    iEditavel = mvariEditavel
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sSiglaUMBase(ByVal vData As String)
    mvarsSiglaUMBase = vData
End Property

Public Property Get sSiglaUMBase() As String
    sSiglaUMBase = mvarsSiglaUMBase
End Property

Public Property Let iClasse(ByVal vData As Integer)
    mvariClasse = vData
End Property

Public Property Get iClasse() As Integer
    iClasse = mvariClasse
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property



