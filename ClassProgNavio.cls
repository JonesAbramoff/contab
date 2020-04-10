VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProgNavio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsNavio As String
Private mvarsTerminal As String
Private mvarsArmador As String
Private mvarsAgMaritima As String
Private mvarsViagem As String
Private mvardtDataChegada As Date
Private mvardtHoraChegada As Date
Private mvardtDataDeadLine As Date
Private mvardtHoraDeadLine As Date
Private mvarsObservacao As String

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dtHoraDeadLine(ByVal vData As Date)
    mvardtHoraDeadLine = vData
End Property

Public Property Get dtHoraDeadLine() As Date
    dtHoraDeadLine = mvardtHoraDeadLine
End Property

Public Property Let dtDataDeadLine(ByVal vData As Date)
    mvardtDataDeadLine = vData
End Property

Public Property Get dtDataDeadLine() As Date
    dtDataDeadLine = mvardtDataDeadLine
End Property

Public Property Let dtHoraChegada(ByVal vData As Date)
    mvardtHoraChegada = vData
End Property

Public Property Get dtHoraChegada() As Date
    dtHoraChegada = mvardtHoraChegada
End Property

Public Property Let dtDataChegada(ByVal vData As Date)
    mvardtDataChegada = vData
End Property

Public Property Get dtDataChegada() As Date
    dtDataChegada = mvardtDataChegada
End Property

Public Property Let sViagem(ByVal vData As String)
    mvarsViagem = vData
End Property

Public Property Get sViagem() As String
    sViagem = mvarsViagem
End Property

Public Property Let sAgMaritima(ByVal vData As String)
    mvarsAgMaritima = vData
End Property

Public Property Get sAgMaritima() As String
    sAgMaritima = mvarsAgMaritima
End Property

Public Property Let sArmador(ByVal vData As String)
    mvarsArmador = vData
End Property

Public Property Get sArmador() As String
    sArmador = mvarsArmador
End Property

Public Property Let sTerminal(ByVal vData As String)
    mvarsTerminal = vData
End Property

Public Property Get sTerminal() As String
    sTerminal = mvarsTerminal
End Property

Public Property Let sNavio(ByVal vData As String)
    mvarsNavio = vData
End Property

Public Property Get sNavio() As String
    sNavio = mvarsNavio
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property



