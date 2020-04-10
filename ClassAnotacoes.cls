VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAnotacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsID As String
Private mvardtDataAlteracao As Date
Private mvariTipoDocOrigem As Integer
Private mvarsTitulo As String
Private mvarlNumIntDoc As Long
Private mvarsTextoCompleto As String
Private mvarsIdExterno As String
Private mvarcolTexto As Collection '??? lixo, apagar

Public Property Set colTexto(ByVal vData As Collection)
    Set mvarcolTexto = vData
End Property

Public Property Get colTexto() As Collection
    Set colTexto = mvarcolTexto
End Property

Public Property Let sIdExterno(ByVal vData As String)
    mvarsIdExterno = vData
End Property

Public Property Get sIdExterno() As String
    sIdExterno = mvarsIdExterno
End Property


Public Property Let sTextoCompleto(ByVal vData As String)
    mvarsTextoCompleto = vData
End Property

Public Property Get sTextoCompleto() As String
    sTextoCompleto = mvarsTextoCompleto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sTitulo(ByVal vData As String)
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property

Public Property Let iTipoDocOrigem(ByVal vData As Integer)
    mvariTipoDocOrigem = vData
End Property

Public Property Get iTipoDocOrigem() As Integer
    iTipoDocOrigem = mvariTipoDocOrigem
End Property

Public Property Let dtDataAlteracao(ByVal vData As Date)
    mvardtDataAlteracao = vData
End Property

Public Property Get dtDataAlteracao() As Date
    dtDataAlteracao = mvardtDataAlteracao
End Property

Public Property Let sID(ByVal vData As String)
    mvarsID = vData
End Property

Public Property Get sID() As String
    sID = mvarsID
End Property
