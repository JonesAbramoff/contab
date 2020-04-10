VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsSetor As String
Private mvarsTelefone As String
Private mvarsEmail As String
Private mvarsFax As String
Private mvarsContato As String
Private mvarlNumIntDocOrigem As Long
Private mvariTipoNumIntDocOrigem As Integer
Private mvarlNumIntDoc As Long

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
    mvariTipoNumIntDocOrigem = vData
End Property

Public Property Get iTipoNumIntDocOrigem() As Integer
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sFax(ByVal vData As String)
    mvarsFax = vData
End Property

Public Property Get sFax() As String
    sFax = mvarsFax
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let sTelefone(ByVal vData As String)
    mvarsTelefone = vData
End Property

Public Property Get sTelefone() As String
    sTelefone = mvarsTelefone
End Property

Public Property Let sSetor(ByVal vData As String)
    mvarsSetor = vData
End Property

Public Property Get sSetor() As String
    sSetor = mvarsSetor
End Property

