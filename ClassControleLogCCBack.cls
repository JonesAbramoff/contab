VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassControleLogCCBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlSequencial As Long
Private mvarlNumIntDocInicial As Long
Private mvarlNumIntDocFinal As Long
Private mvardtData As Date

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumIntDocFinal(ByVal vData As Long)
    mvarlNumIntDocFinal = vData
End Property

Public Property Get lNumIntDocFinal() As Long
    lNumIntDocFinal = mvarlNumIntDocFinal
End Property

Public Property Let lNumIntDocInicial(ByVal vData As Long)
    mvarlNumIntDocInicial = vData
End Property

Public Property Get lNumIntDocInicial() As Long
    lNumIntDocInicial = mvarlNumIntDocInicial
End Property

Public Property Let lSequencial(ByVal vData As Long)
    mvarlSequencial = vData
End Property

Public Property Get lSequencial() As Long
    lSequencial = mvarlSequencial
End Property

