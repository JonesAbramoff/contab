VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOSApMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigoAp As Long
Private mvariFilialEmpresa As Integer
Private mvariSeq As Integer
Private mvarlCodigoMO As Long
Private mvardHorasGastas As Double
Private mvarsOS As String
Private mvarsProdutoOS As String
Private mvariTipo As Integer
Private mvarlNumIntDoc As Long

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lCodigoAp(ByVal vData As Long)
    mvarlCodigoAp = vData
End Property

Public Property Get lCodigoAp() As Long
    lCodigoAp = mvarlCodigoAp
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lCodigoMO(ByVal vData As Long)
    mvarlCodigoMO = vData
End Property

Public Property Get lCodigoMO() As Long
    lCodigoMO = mvarlCodigoMO
End Property

Public Property Let dHorasGastas(ByVal vData As Double)
    mvardHorasGastas = vData
End Property

Public Property Get dHorasGastas() As Double
    dHorasGastas = mvardHorasGastas
End Property

Public Property Let sOS(ByVal vData As String)
    mvarsOS = vData
End Property

Public Property Get sOS() As String
    sOS = mvarsOS
End Property

Public Property Let sProdutoOS(ByVal vData As String)
    mvarsProdutoOS = vData
End Property

Public Property Get sProdutoOS() As String
    sProdutoOS = mvarsProdutoOS
End Property
