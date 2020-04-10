VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNatOpPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvariTipoOperacao As Integer
Private mvariTipoAtividadeEmp As Integer
Private mvariTipoTribEmp As Integer
Private mvarsCFOPEmp As String
Private mvariTipoAtividadeExt As Integer
Private mvariTipoTribExt As Integer
Private mvarsCFOPExt As String
Private mvariPadrao As Integer

Public Property Let iPadrao(ByVal vData As Integer)
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

Public Property Let sCFOPExt(ByVal vData As String)
    mvarsCFOPExt = vData
End Property

Public Property Get sCFOPExt() As String
    sCFOPExt = mvarsCFOPExt
End Property

Public Property Let iTipoTribExt(ByVal vData As Integer)
    mvariTipoTribExt = vData
End Property

Public Property Get iTipoTribExt() As Integer
    iTipoTribExt = mvariTipoTribExt
End Property

Public Property Let iTipoAtividadeExt(ByVal vData As Integer)
    mvariTipoAtividadeExt = vData
End Property

Public Property Get iTipoAtividadeExt() As Integer
    iTipoAtividadeExt = mvariTipoAtividadeExt
End Property

Public Property Let sCFOPEmp(ByVal vData As String)
    mvarsCFOPEmp = vData
End Property

Public Property Get sCFOPEmp() As String
    sCFOPEmp = mvarsCFOPEmp
End Property

Public Property Let iTipoTribEmp(ByVal vData As Integer)
    mvariTipoTribEmp = vData
End Property

Public Property Get iTipoTribEmp() As Integer
    iTipoTribEmp = mvariTipoTribEmp
End Property

Public Property Let iTipoAtividadeEmp(ByVal vData As Integer)
    mvariTipoAtividadeEmp = vData
End Property

Public Property Get iTipoAtividadeEmp() As Integer
    iTipoAtividadeEmp = mvariTipoAtividadeEmp
End Property

Public Property Let iTipoOperacao(ByVal vData As Integer)
    mvariTipoOperacao = vData
End Property

Public Property Get iTipoOperacao() As Integer
    iTipoOperacao = mvariTipoOperacao
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property
