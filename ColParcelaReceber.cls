VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColParcelaReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassParcelaReceber"
Attribute VB_Ext_KEY = "Member0" ,"ClassParcelaReceber"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ClassParcelaReceber"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntDoc As Long, lNumIntTitulo As Long, iNumParcela As Integer, iStatus As Integer, dtDataVencimento As Date, dtDataVencimentoReal As Date, dSaldo As Double, dValor As Double, iProxSeqBaixa As Integer, iCarteiraCobranca As Integer, iCobrador As Integer, sNumTitCobrador As String, lNumBordero As Long, iCodOcorrencia As Integer, iInstrucaoPrimaria As Integer, iInstrucaoSecundaria As Integer, dJuros As Double, iDiasDeProtesto As Integer, iDesconto1Codigo As Integer, dtDesconto1Ate As Date, dDesconto1Valor As Double, iDesconto2Codigo As Integer, dtDesconto2Ate As Date, dDesconto2Valor As Double, iDesconto3Codigo As Integer, dtDesconto3Ate As Date, dDesconto3Valor As Double, lNumIntCheque As Long, iAceite As Integer, iDescontada As Integer, iProxSeqOcorr As Integer, iPrevisao As Integer, sObservacao As String, dValorOriginal As Double, Optional sKey As String) As ClassParcelaReceber

Dim objNewMember As ClassParcelaReceber
Set objNewMember = New ClassParcelaReceber

    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.lNumIntTitulo = lNumIntTitulo
    objNewMember.iNumParcela = iNumParcela
    objNewMember.iStatus = iStatus
    objNewMember.dtDataVencimento = dtDataVencimento
    objNewMember.dtDataVencimentoReal = dtDataVencimentoReal
    objNewMember.dSaldo = dSaldo
    objNewMember.dValor = dValor
    objNewMember.iCobrador = iCobrador
    objNewMember.iCarteiraCobranca = iCarteiraCobranca
    objNewMember.sNumTitCobrador = sNumTitCobrador
    objNewMember.lNumIntCheque = lNumIntCheque
    objNewMember.iProxSeqBaixa = iProxSeqBaixa
    objNewMember.iProxSeqOcorr = iProxSeqOcorr
    objNewMember.iDesconto1Codigo = iDesconto1Codigo
    objNewMember.dtDesconto1Ate = dtDesconto1Ate
    objNewMember.dDesconto1Valor = dDesconto1Valor
    objNewMember.iDesconto2Codigo = iDesconto2Codigo
    objNewMember.dtDesconto2Ate = dtDesconto2Ate
    objNewMember.dDesconto2Valor = dDesconto2Valor
    objNewMember.iDesconto3Codigo = iDesconto3Codigo
    objNewMember.dtDesconto3Ate = dtDesconto3Ate
    objNewMember.dDesconto3Valor = dDesconto3Valor
    objNewMember.iAceite = iAceite
    objNewMember.iDescontada = iDescontada
    objNewMember.iPrevisao = iPrevisao
    objNewMember.sObservacao = sObservacao
    objNewMember.dValorOriginal = dValorOriginal
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassParcelaReceber
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

Public Function AddObj(ByVal objParcRec As ClassParcelaReceber, Optional sKey As String) As ClassParcelaReceber

    If Len(sKey) = 0 Then
        mCol.Add objParcRec
    Else
        mCol.Add objParcRec, sKey
    End If

    Set AddObj = objParcRec

End Function


