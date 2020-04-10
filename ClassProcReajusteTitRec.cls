VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProcReajTitRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvardtDataProc As Date
Private mvardHoraProc As Double
Private mvarsUsuario As String
Private mvardtAtualizadoAte As Date
Private mvariFilialEmpresa As Integer
Private mvarcolCcl As New Collection

Public Property Set colCcl(ByVal vData As Object)
    Set mvarcolCcl = vData
End Property

Public Property Get colCcl() As Collection
    Set colCcl = mvarcolCcl
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dtDataProc(ByVal vData As Date)
    mvardtDataProc = vData
End Property

Public Property Get dtDataProc() As Date
    dtDataProc = mvardtDataProc
End Property

Public Property Let dHoraProc(ByVal vData As Double)
    mvardHoraProc = vData
End Property

Public Property Get dHoraProc() As Double
    dHoraProc = mvardHoraProc
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dtAtualizadoAte(ByVal vData As Date)
    mvardtAtualizadoAte = vData
End Property

Public Property Get dtAtualizadoAte() As Date
    dtAtualizadoAte = mvardtAtualizadoAte
End Property
