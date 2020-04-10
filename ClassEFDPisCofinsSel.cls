VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEFDPisCofinsSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public NUM_MAX_COMANDOS As Integer

Public colUM As New Collection
Public objFilialEmpresa As New AdmFiliais
Public iUltimoCmdUsado As Integer
Public colSpedRegistro As New Collection
Public objTela As Object
Public objSpedFiscal As New ClassSpedFiscal
Public colModulo As New AdmColModulo

Private mvarsDiretorio As String
Private mvarsReciboAnterior As String
Private mvardtDataIni As Date
Private mvardtDataFim As Date
Private mvariFilialEmpresa As Integer
Private mvariSituacaoEspecial As Integer
Private mvariTipo As Integer
Private mvarlNumIntSped As Long
Private mvariCodVersao As Integer
Private mvarlComandoInsert As Long
Private mvarlComandoInsertReg As Long
Private mvaralComando(0 To 200) As Long
Private colRegistros As New Collection
Private mvarlNumTotalItensNF As Long
Private mvarlNumTotalPrevisto As Long
Private mvarlNumTotalReal As Long
Private mvarlNumTotalProcessados As Long
Private mvariRegimeApuracao As Integer

Private mvaralComando0200() As Long
Private mvaralComando0150() As Long
Private mvaralComando0145() As Long
Private mvaralComando0450() As Long

Private mvarlComando0190 As Long

Public Property Let iRegimeApuracao(ByVal vData As Integer)
    mvariRegimeApuracao = vData
End Property

Public Property Get iRegimeApuracao() As Integer
    iRegimeApuracao = mvariRegimeApuracao
End Property

Public Property Get TRATA_REG_CAIXA() As Long
    TRATA_REG_CAIXA = 1
End Property

Public Property Get TRATA_REG_COMPETENCIA() As Long
    TRATA_REG_COMPETENCIA = 0
End Property

Public Sub alComando0145_AbreComandos(ByVal iQtde As Integer)
Dim iIndice As Integer

On Error GoTo Erro_alComando0145_AbreComandos

    If mvaralComando0145(1) = 0 Then
    
        'Abertura de Comando
        For iIndice = 1 To iQtde
            mvaralComando0145(iIndice) = Me.ProxComando
        Next

    End If
    
    Exit Sub
    
Erro_alComando0145_AbreComandos:

    ReDim mvaralComando0145(1 To iQtde) As Long
    Resume Next
    
End Sub

Public Property Get alComando0145(ByVal iIndice As Integer) As Long
    alComando0145 = mvaralComando0145(iIndice)
End Property

Public Sub alComando0450_AbreComandos(ByVal iQtde As Integer)
Dim iIndice As Integer

On Error GoTo Erro_alComando0450_AbreComandos

    If mvaralComando0450(1) = 0 Then
    
        'Abertura de Comando
        For iIndice = 1 To iQtde
            mvaralComando0450(iIndice) = Me.ProxComando
        Next

    End If
    
    Exit Sub
    
Erro_alComando0450_AbreComandos:

    ReDim mvaralComando0450(1 To iQtde) As Long
    Resume Next
    
End Sub

Public Property Get alComando0450(ByVal iIndice As Integer) As Long
    alComando0450 = mvaralComando0450(iIndice)
End Property

Public Property Let alComando0450(ByVal iIndice As Integer, ByVal vData As Long)
    mvaralComando0450(iIndice) = vData
End Property

Public Property Get lComando0190() As Long
    
    If mvarlComando0190 = 0 Then
        mvarlComando0190 = Me.ProxComando
    End If
    
    lComando0190 = mvarlComando0190
    
End Property

Public Sub alComando0150_AbreComandos(ByVal iQtde As Integer)
Dim iIndice As Integer

On Error GoTo Erro_alComando0150_AbreComandos

    If mvaralComando0150(1) = 0 Then
    
        'Abertura de Comando
        For iIndice = 1 To iQtde
            mvaralComando0150(iIndice) = Me.ProxComando
        Next

    End If
    
    Exit Sub
    
Erro_alComando0150_AbreComandos:

    ReDim mvaralComando0150(1 To iQtde) As Long
    Resume Next
    
End Sub

Public Property Get alComando0150(ByVal iIndice As Integer) As Long
    alComando0150 = mvaralComando0150(iIndice)
End Property

Public Property Let alComando0150(ByVal iIndice As Integer, ByVal vData As Long)
    mvaralComando0150(iIndice) = vData
End Property

Public Sub alComando0200_AbreComandos(ByVal iQtde As Integer)
Dim iIndice As Integer

On Error GoTo Erro_alComando0200_AbreComandos

    If mvaralComando0200(1) = 0 Then
    
        'Abertura de Comando
        For iIndice = 1 To iQtde
            mvaralComando0200(iIndice) = Me.ProxComando
        Next

    End If
    
    Exit Sub
    
Erro_alComando0200_AbreComandos:

    ReDim mvaralComando0200(1 To iQtde) As Long
    Resume Next
    
End Sub

Public Property Get alComando0200(ByVal iIndice As Integer) As Long
    alComando0200 = mvaralComando0200(iIndice)
End Property

Public Property Let alComando0200(ByVal iIndice As Integer, ByVal vData As Long)
    mvaralComando0200(iIndice) = vData
End Property

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

Public Property Let lNumTotalItensNF(ByVal vData As Long)
    mvarlNumTotalItensNF = vData
End Property

Public Property Get lNumTotalItensNF() As Long
    lNumTotalItensNF = mvarlNumTotalItensNF
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

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iSituacaoEspecial(ByVal vData As Integer)
    mvariSituacaoEspecial = vData
End Property

Public Property Get iSituacaoEspecial() As Integer
    iSituacaoEspecial = mvariSituacaoEspecial
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

Public Property Let sReciboAnterior(ByVal vData As String)
    mvarsReciboAnterior = vData
End Property

Public Property Get sReciboAnterior() As String
    sReciboAnterior = mvarsReciboAnterior
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Private Sub Class_Initialize()
    iUltimoCmdUsado = 10
    NUM_MAX_COMANDOS = 200
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
