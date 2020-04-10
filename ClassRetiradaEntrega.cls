VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRetiradaEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarsCNPJCPFRet As String
Private mvarsCNPJCPFEnt As String
Private mvarlClienteRet As Long
Private mvariFilialCliRet As Integer
Private mvarlFornecedorRet As Long
Private mvariFilialFornRet As Integer
Private mvarlClienteEnt As Long
Private mvariFilialCliEnt As Integer
Private mvarlFornecedorEnt As Long
Private mvariFilialFornEnt As Integer
Private mvarobjEnderecoRet As ClassEndereco
Private mvarobjEnderecoEnt As ClassEndereco
Private mvarlEnderecoRet As Long
Private mvarlEnderecoEnt As Long
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lEnderecoEnt(ByVal vData As Long)
    mvarlEnderecoEnt = vData
End Property

Public Property Get lEnderecoEnt() As Long
    lEnderecoEnt = mvarlEnderecoEnt
End Property

Public Property Let lEnderecoRet(ByVal vData As Long)
    mvarlEnderecoRet = vData
End Property

Public Property Get lEnderecoRet() As Long
    lEnderecoRet = mvarlEnderecoRet
End Property

Public Property Set objEnderecoEnt(ByVal vData As Object)
    Set mvarobjEnderecoEnt = vData
End Property

Public Property Get objEnderecoEnt() As Object
    Set objEnderecoEnt = mvarobjEnderecoEnt
End Property

Public Property Set objEnderecoRet(ByVal vData As Object)
    Set mvarobjEnderecoRet = vData
End Property

Public Property Get objEnderecoRet() As Object
    Set objEnderecoRet = mvarobjEnderecoRet
End Property


Public Property Let iFilialFornEnt(ByVal vData As Integer)
    mvariFilialFornEnt = vData
End Property

Public Property Get iFilialFornEnt() As Integer
    iFilialFornEnt = mvariFilialFornEnt
End Property

Public Property Let lFornecedorEnt(ByVal vData As Long)
    mvarlFornecedorEnt = vData
End Property

Public Property Get lFornecedorEnt() As Long
    lFornecedorEnt = mvarlFornecedorEnt
End Property

Public Property Let iFilialCliEnt(ByVal vData As Integer)
    mvariFilialCliEnt = vData
End Property

Public Property Get iFilialCliEnt() As Integer
    iFilialCliEnt = mvariFilialCliEnt
End Property

Public Property Let lClienteEnt(ByVal vData As Long)
    mvarlClienteEnt = vData
End Property

Public Property Get lClienteEnt() As Long
    lClienteEnt = mvarlClienteEnt
End Property

Public Property Let iFilialFornRet(ByVal vData As Integer)
    mvariFilialFornRet = vData
End Property

Public Property Get iFilialFornRet() As Integer
    iFilialFornRet = mvariFilialFornRet
End Property

Public Property Let lFornecedorRet(ByVal vData As Long)
    mvarlFornecedorRet = vData
End Property

Public Property Get lFornecedorRet() As Long
    lFornecedorRet = mvarlFornecedorRet
End Property

Public Property Let iFilialCliRet(ByVal vData As Integer)
    mvariFilialCliRet = vData
End Property

Public Property Get iFilialCliRet() As Integer
    iFilialCliRet = mvariFilialCliRet
End Property

Public Property Let lClienteRet(ByVal vData As Long)
    mvarlClienteRet = vData
End Property

Public Property Get lClienteRet() As Long
    lClienteRet = mvarlClienteRet
End Property

Public Property Let sCNPJCPFEnt(ByVal vData As String)
    mvarsCNPJCPFEnt = vData
End Property

Public Property Get sCNPJCPFEnt() As String
    sCNPJCPFEnt = mvarsCNPJCPFEnt
End Property

Public Property Let sCNPJCPFRet(ByVal vData As String)
    mvarsCNPJCPFRet = vData
End Property

Public Property Get sCNPJCPFRet() As String
    sCNPJCPFRet = mvarsCNPJCPFRet
End Property
