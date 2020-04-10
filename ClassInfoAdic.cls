VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjRetEnt As ClassRetiradaEntrega
Private mvarobjExportacao As ClassInfoAdicExportacao
Private mvarobjCompra As ClassInfoAdicCompra
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvarsNatureza As String
Private mvarsCcl As String

Public objAnexos As New ClassAnexos

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

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

Public Property Set objCompra(ByVal vData As ClassInfoAdicCompra)
    Set mvarobjCompra = vData
End Property

Public Property Get objCompra() As ClassInfoAdicCompra
    Set objCompra = mvarobjCompra
End Property

Public Property Set objExportacao(ByVal vData As ClassInfoAdicExportacao)
    Set mvarobjExportacao = vData
End Property

Public Property Get objExportacao() As ClassInfoAdicExportacao
    Set objExportacao = mvarobjExportacao
End Property

Public Property Set objRetEnt(ByVal vData As ClassRetiradaEntrega)
    Set mvarobjRetEnt = vData
End Property

Public Property Get objRetEnt() As ClassRetiradaEntrega
    Set objRetEnt = mvarobjRetEnt
End Property
