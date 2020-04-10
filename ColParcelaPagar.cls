VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColParcelaPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassParcelaPagar"
Attribute VB_Ext_KEY = "Member0" ,"ClassParcelaPagar"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntDoc As Long, lNumIntTitulo As Long, iNumParcela As Integer, iStatus As Integer, dtDataVencimento As Date, dtDataVencimentoReal As Date, dSaldo As Double, dValor As Double, iPortador As Integer, iProxSeqBaixa As Integer, iTipoCobranca As Integer, iBancoCobrador As Integer, sNossoNumero As String, sCodigoDeBarras As String, Optional sKey As String, Optional ByVal dValorOriginal As Double, Optional ByVal iMotivoDiferenca As Integer, Optional ByVal sCodUsuarioLib As String = "", Optional ByVal dtDataLib As Date = DATA_NULA) As ClassParcelaPagar

Dim objNewMember As ClassParcelaPagar
Set objNewMember = New ClassParcelaPagar

    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.lNumIntTitulo = lNumIntTitulo
    objNewMember.iNumParcela = iNumParcela
    objNewMember.iStatus = iStatus
    objNewMember.dtDataVencimento = dtDataVencimento
    objNewMember.dtDataVencimentoReal = dtDataVencimentoReal
    objNewMember.dSaldo = dSaldo
    objNewMember.dValor = dValor
    objNewMember.iPortador = iPortador
    objNewMember.iProxSeqBaixa = iProxSeqBaixa
    objNewMember.iTipoCobranca = iTipoCobranca
    objNewMember.iBancoCobrador = iBancoCobrador
    objNewMember.sNossoNumero = sNossoNumero
    objNewMember.sCodigoDeBarras = sCodigoDeBarras
    objNewMember.sCodUsuarioLib = sCodUsuarioLib
    objNewMember.dtDataLib = dtDataLib

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    If Not IsMissing(dValorOriginal) Then
        objNewMember.dValorOriginal = dValorOriginal
    End If
    
    If Not IsMissing(iMotivoDiferenca) Then
        objNewMember.iMotivoDiferenca = iMotivoDiferenca
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassParcelaPagar
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

