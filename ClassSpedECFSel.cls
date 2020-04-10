VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedECFSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public NUM_MAX_COMANDOS As Integer

Public iUltimoCmdUsado As Integer
Public objFilialEmpresa As New AdmFiliais
Public colSpedRegistro As New Collection
Public objTela As Object
Public objSpedFiscal As New ClassSpedFiscal
Public lFreeFile As Long
Public lModelo As Long
Public iTipoModelo As Integer

Private mvarsDiretorio As String
Private mvardtDataIni As Date
Private mvardtDataFim As Date
Private mvariFilialEmpresa As Integer
Private mvarlNumIntSped As Long
Private mvariCodVersao As Integer
Private mvarlComandoInsert As Long
Private mvarlComandoInsertReg As Long
Private mvaralComando(0 To 100) As Long
Private mvarlNumTotalPrevisto As Long
Private mvarlNumTotalReal As Long
Private mvarlNumTotalProcessados As Long
Private colRegistros As New Collection

Public Property Let lNumTotalProcessados(ByVal vData As Long)
    mvarlNumTotalProcessados = vData
End Property

Public Property Get lNumTotalProcessados() As Long
    lNumTotalProcessados = mvarlNumTotalProcessados
End Property

Public Property Let lNumTotalReal(ByVal vData As Long)
    mvarlNumTotalReal = vData
End Property

Public Property Get lNumTotalReal() As Long
    lNumTotalReal = mvarlNumTotalReal
End Property

Public Property Let lNumTotalPrevisto(ByVal vData As Long)
    mvarlNumTotalPrevisto = vData
End Property

Public Property Get lNumTotalPrevisto() As Long
    lNumTotalPrevisto = mvarlNumTotalPrevisto
End Property

Public Property Get alComando(ByVal iIndice As Integer) As Long
    alComando = mvaralComando(iIndice)
End Property

Public Property Let alComando(ByVal iIndice As Integer, ByVal vData As Long)
    mvaralComando(iIndice) = vData
End Property

Public Property Get lComandoInsert() As Long
    lComandoInsert = alComando(0)
End Property

Public Property Get lComandoInsertReg() As Long
    lComandoInsertReg = alComando(1)
End Property

Public Property Let iCodVersao(ByVal vData As Integer)
    mvariCodVersao = vData
End Property

Public Property Get iCodVersao() As Integer
    iCodVersao = mvariCodVersao
End Property

Public Property Let lNumIntSped(ByVal vData As Long)
    mvarlNumIntSped = vData
End Property

Public Property Get lNumIntSped() As Long
    lNumIntSped = mvarlNumIntSped
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataIni(ByVal vData As Date)
    mvardtDataIni = vData
End Property

Public Property Get dtDataIni() As Date
    dtDataIni = mvardtDataIni
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Private Sub Class_Initialize()
    iUltimoCmdUsado = 10
    NUM_MAX_COMANDOS = 100
End Sub

Public Function ProxComando() As Long
    iUltimoCmdUsado = iUltimoCmdUsado + 1
    ProxComando = alComando(iUltimoCmdUsado)
End Function

Public Function UltimoCmdUsado() As Long
    UltimoCmdUsado = alComando(iUltimoCmdUsado)
End Function

Public Sub Inicia_Reg()
    Set colRegistros = New Collection
End Sub

Public Sub Reg_Campo_add(ByVal iIndice As Integer, ByVal sCampo As String, ByVal sValor As String)
    Dim objCampo As New ClassSpedFiscalReg
    objCampo.iIndice = iIndice
    objCampo.sCampo = sCampo
    objCampo.sValor = Replace(sValor, "|", " ")
    colRegistros.Add objCampo
End Sub

Public Sub Finaliza_Reg()
    Call Monta_Reg
    Set colRegistros = Nothing
End Sub

Private Sub Monta_Reg()
Dim sReg As String
Dim objCampo As New ClassSpedFiscalReg
    sReg = "|"
    For Each objCampo In colRegistros
        sReg = sReg & objCampo.sValor & "|"
    Next
    objSpedFiscal.lSeq = objSpedFiscal.lSeq + 1
    objSpedFiscal.sRegistro = sReg
End Sub
