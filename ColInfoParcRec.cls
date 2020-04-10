VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColInfoParcRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassInfoParcRec"
Attribute VB_Ext_KEY = "Member0" ,"ClassInfoParcRec"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(sSiglaDocumento As String, iMarcada As Integer, sNomeRedCliente As String, sRazaoSocialCli As String, iFilialCliente As Integer, lCliente As Long, lNumTitulo As Long, iNumParcela As Integer, lNumIntParc As Long, dValorDesconto As Double, dValorMulta As Double, dValorJuros As Double, dValorOriginal As Double, dValor As Double, dtVencimento As Date, iPadraoCobranca As Integer, Optional sKey As String) As ClassInfoParcRec
    
Dim objNewMember As ClassInfoParcRec
Set objNewMember = New ClassInfoParcRec

    objNewMember.sSiglaDocumento = sSiglaDocumento
    objNewMember.iMarcada = iMarcada
    objNewMember.sNomeRedCliente = sNomeRedCliente
    objNewMember.sRazaoSocialCli = sRazaoSocialCli
    objNewMember.iFilialCliente = iFilialCliente
    objNewMember.lCliente = lCliente
    objNewMember.lNumTitulo = lNumTitulo
    objNewMember.iNumParcela = iNumParcela
    objNewMember.lNumIntParc = lNumIntParc
    objNewMember.dValorDesconto = dValorDesconto
    objNewMember.dValorMulta = dValorMulta
    objNewMember.dValorJuros = dValorJuros
    objNewMember.dValorOriginal = dValorOriginal
    objNewMember.dValor = dValor
    objNewMember.dtVencimento = dtVencimento
    objNewMember.iPadraoCobranca = iPadraoCobranca

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassInfoParcRec
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

