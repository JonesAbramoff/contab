VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColPagtoAntecipado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassAntecipPag"
Attribute VB_Ext_KEY = "Member0" ,"ClassAntecipPag"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ClassAntecipPag"
Option Explicit

Private mCol As Collection

Public Function Add(lFornecedor As Long, iFilial As Integer, iCodConta As Integer, lSequencial As Long, dtData As Date, dValor As Double, iTipoMeioPagto As Integer, lNumero As Long, sHistorico As String, lNumMovto As Long, dSaldoNaoApropriado As Double, lNumIntPag As Long, sDescricaoMeioPagto As String, dValorBaixado As Double, iExcluido As Integer, Optional sKey As String) As ClassAntecipPag
    
Dim objNewMember As ClassAntecipPag
Set objNewMember = New ClassAntecipPag

    objNewMember.lFornecedor = lFornecedor
    objNewMember.iFilial = iFilial
    objNewMember.iCodConta = iCodConta
    objNewMember.lSequencial = lSequencial
    objNewMember.dtData = dtData
    objNewMember.dValor = dValor
    objNewMember.iTipoMeioPagto = iTipoMeioPagto
    objNewMember.lNumero = lNumero
    objNewMember.sHistorico = sHistorico
    objNewMember.lNumMovto = lNumMovto
    objNewMember.dSaldoNaoApropriado = dSaldoNaoApropriado
    objNewMember.lNumIntPag = lNumIntPag
    objNewMember.sDescricaoMeioPagto = sDescricaoMeioPagto
    objNewMember.dValorBaixado = dValorBaixado
    objNewMember.iExcluido = iExcluido

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassAntecipPag
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

