VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvariCodigo As Integer
Private mvarsContato As String
Private mvarsSetor As String
Private mvarsCargo As String
Private mvarsTelefone As String
Private mvarsFax As String
Private mvarsEmail As String
Private mvariPadrao As Integer
Private mvardtDataNasc As Date
Private mvarsOutrosMeioComunic As String

Public Property Let sOutrosMeioComunic(ByVal vData As String)
    mvarsOutrosMeioComunic = vData
End Property

Public Property Get sOutrosMeioComunic() As String
    sOutrosMeioComunic = mvarsOutrosMeioComunic
End Property

Public Property Let dtDataNasc(ByVal vData As Date)
    mvardtDataNasc = vData
End Property

Public Property Get dtDataNasc() As Date
    dtDataNasc = mvardtDataNasc
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sSetor(ByVal vData As String)
    mvarsSetor = vData
End Property

Public Property Get sSetor() As String
    sSetor = mvarsSetor
End Property

Public Property Let sCargo(ByVal vData As String)
    mvarsCargo = vData
End Property

Public Property Get sCargo() As String
    sCargo = mvarsCargo
End Property

Public Property Let sTelefone(ByVal vData As String)
    mvarsTelefone = vData
End Property

Public Property Get sTelefone() As String
    sTelefone = mvarsTelefone
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

Public Property Let iPadrao(ByVal vData As Integer)
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

