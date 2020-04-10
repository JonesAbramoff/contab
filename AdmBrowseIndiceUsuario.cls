VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseIndiceUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsCodUsuario As String
Private mvariIndice As Integer
Private mvarsNomeIndice As String
Private mvarsOrdenacaoSQL As String
Private mvarsSelecaoSQL As String

Public Property Let sSelecaoSQL(ByVal vData As String)
    mvarsSelecaoSQL = vData
End Property

Public Property Get sSelecaoSQL() As String
    sSelecaoSQL = mvarsSelecaoSQL
End Property

Public Property Let sOrdenacaoSQL(ByVal vData As String)
    mvarsOrdenacaoSQL = vData
End Property

Public Property Get sOrdenacaoSQL() As String
    sOrdenacaoSQL = mvarsOrdenacaoSQL
End Property

Public Property Let sNomeIndice(ByVal vData As String)
    mvarsNomeIndice = vData
End Property

Public Property Get sNomeIndice() As String
    sNomeIndice = mvarsNomeIndice
End Property

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

