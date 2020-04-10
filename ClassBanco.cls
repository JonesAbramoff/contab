VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodBanco As Integer
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvarsLayoutCheque As String
Private mvarsLayoutBoleto As String
Private mvariAtivo As Integer
Private mvardtDataLog As Date

Public iLayoutCnabConciliacao As Integer

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let sLayoutBoleto(ByVal vData As String)
    mvarsLayoutBoleto = vData
End Property

Public Property Get sLayoutBoleto() As String
    sLayoutBoleto = mvarsLayoutBoleto
End Property

Public Property Let sLayoutCheque(ByVal vData As String)
    mvarsLayoutCheque = vData
End Property

Public Property Get sLayoutCheque() As String
    sLayoutCheque = mvarsLayoutCheque
End Property

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

Public Property Let iCodBanco(ByVal vData As Integer)
    mvariCodBanco = vData
End Property

Public Property Get iCodBanco() As Integer
    iCodBanco = mvariCodBanco
End Property



