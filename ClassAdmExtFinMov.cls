VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdmExtFinMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntDoc As Long
Public lNumIntArq As Long
Public iTipo As Integer
Public iCodConta As Integer
Public dtData As Date
Public dValorBruto As Double
Public dValorComissao As Double
Public dValorRejeitado As Double
Public dValorLiq As Double
Public lNumMovto As Long
Public sEstabelecimento As String

Public colRO As New Collection
Public colMovDet As New Collection
Public colMovDetCred As New Collection
Public iFilialEmpresa As Integer

'########################################
'AUXILIARES
Public sNomeCtaCorrente As String

Private mvariQtdDep As Integer
Private mvariQtdParc As Integer
Private mvariQtdParcEnc As Integer
Private mvardVlrDep As Double
Private mvardVlrParc As Double
Private mvardVlrParcEnc As Double
Private mvariTotalReg As Integer
Private mvariBandeira As Integer

Public Property Let iBandeira(ByVal vData As Integer)
    mvariBandeira = vData
End Property

Public Property Get iBandeira() As Integer
    iBandeira = mvariBandeira
End Property

Public Property Let iTotalReg(ByVal vData As Integer)
    mvariTotalReg = vData
End Property

Public Property Get iTotalReg() As Integer
    iTotalReg = mvariTotalReg
End Property

Public Property Let dVlrDep(ByVal vData As Double)
    mvardVlrDep = vData
End Property

Public Property Get dVlrDep() As Double
    dVlrDep = mvardVlrDep
End Property

Public Property Let dVlrParc(ByVal vData As Double)
    mvardVlrParc = vData
End Property

Public Property Get dVlrParc() As Double
    dVlrParc = mvardVlrParc
End Property

Public Property Let dVlrParcEnc(ByVal vData As Double)
    mvardVlrParcEnc = vData
End Property

Public Property Get dVlrParcEnc() As Double
    dVlrParcEnc = mvardVlrParcEnc
End Property

Public Property Let iQtdDep(ByVal vData As Integer)
    mvariQtdDep = vData
End Property

Public Property Get iQtdDep() As Integer
    iQtdDep = mvariQtdDep
End Property

Public Property Let iQtdParc(ByVal vData As Integer)
    mvariQtdParc = vData
End Property

Public Property Get iQtdParc() As Integer
    iQtdParc = mvariQtdParc
End Property

Public Property Let iQtdParcEnc(ByVal vData As Integer)
    mvariQtdParcEnc = vData
End Property

Public Property Get iQtdParcEnc() As Integer
    iQtdParcEnc = mvariQtdParcEnc
End Property
'AUXILIARES
'########################################
