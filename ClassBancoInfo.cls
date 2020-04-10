VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBancoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodBanco As Integer
Private mvariInfoCodigo As Integer
Private mvarsInfoTexto As String
Private mvarsInfoDescricao As String
Private mvariInfoNivel As Integer

Public Property Let iInfoNivel(ByVal vData As Integer)
    mvariInfoNivel = vData
End Property

Public Property Get iInfoNivel() As Integer
    iInfoNivel = mvariInfoNivel
End Property

Public Property Let sInfoDescricao(ByVal vData As String)
    mvarsInfoDescricao = vData
End Property

Public Property Get sInfoDescricao() As String
    sInfoDescricao = mvarsInfoDescricao
End Property

Public Property Let sInfoTexto(ByVal vData As String)
    mvarsInfoTexto = vData
End Property

Public Property Get sInfoTexto() As String
    sInfoTexto = mvarsInfoTexto
End Property

Public Property Let iInfoCodigo(ByVal vData As Integer)
    mvariInfoCodigo = vData
End Property

Public Property Get iInfoCodigo() As Integer
    iInfoCodigo = mvariInfoCodigo
End Property

Public Property Let iCodBanco(ByVal vData As Integer)
    mvariCodBanco = vData
End Property

Public Property Get iCodBanco() As Integer
    iCodBanco = mvariCodBanco
End Property


