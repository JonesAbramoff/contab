VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJEtapaVistorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntPRJEtapa As Long
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvardtDataValidade As Date
Private mvarsResponsavel As String
Private mvarsLaudo As String
Private mvarobjAnexos As Object

Public Property Set objAnexos(ByVal vData As Object)
    Set mvarobjAnexos = vData
End Property

Public Property Get objAnexos() As Object
    Set objAnexos = mvarobjAnexos
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntPRJEtapa(ByVal vData As Long)
    mvarlNumIntPRJEtapa = vData
End Property

Public Property Get lNumIntPRJEtapa() As Long
    lNumIntPRJEtapa = mvarlNumIntPRJEtapa
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let sLaudo(ByVal vData As String)
    mvarsLaudo = vData
End Property

Public Property Get sLaudo() As String
    sLaudo = mvarsLaudo
End Property
