VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqImportacaoAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarcolArquivos As New Collection
Private mvarsDiretorio As String
Private mvariValidacaoManual As Integer

Public Property Let iValidacaoManual(ByVal vData As Integer)
    mvariValidacaoManual = vData
End Property

Public Property Get iValidacaoManual() As Integer
    iValidacaoManual = mvariValidacaoManual
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Set colArquivos(ByVal vData As Object)
    Set mvarcolArquivos = vData
End Property

Public Property Get colArquivos() As Collection
    Set colArquivos = mvarcolArquivos
End Property
