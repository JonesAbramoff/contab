VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqImportacaoArq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Private mvarobjFile As Object
Private mvarobjTipoArq As New ClassTipoArqIntegracao
Private mvarlTotalRegistros As Long
Private mvardtData As Date
Private mvarsNomeArquivo As String

Public Property Let sNomeArquivo(ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lTotalRegistros(ByVal vData As Long)
    mvarlTotalRegistros = vData
End Property

Public Property Get lTotalRegistros() As Long
    lTotalRegistros = mvarlTotalRegistros
End Property

Public Property Set objTipoArq(ByVal vData As ClassTipoArqIntegracao)
    Set mvarobjTipoArq = vData
End Property

Public Property Get objTipoArq() As ClassTipoArqIntegracao
    Set objTipoArq = mvarobjTipoArq
End Property

Public Property Set objFile(ByVal vData As Object)
    Set mvarobjFile = vData
End Property

Public Property Get objFile() As Object
    Set objFile = mvarobjFile
End Property
