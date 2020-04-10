VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColComissoesNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassComissaoNF"
Attribute VB_Ext_KEY = "Member0" ,"ClassComissaoNF"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntDoc As Long, iCodVendedor As Integer, dValorBase As Double, dPercentual As Double, dValor As Double, dPercentualEmissao As Double, dValorEmissao As Double, Optional sKey As String) As ClassComissaoNF

Dim objNewMember As ClassComissaoNF
Set objNewMember = New ClassComissaoNF

    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.iCodVendedor = iCodVendedor
    objNewMember.dValorBase = dValorBase
    objNewMember.dPercentual = dPercentual
    objNewMember.dValor = dValor
    objNewMember.dPercentualEmissao = dPercentualEmissao
    objNewMember.dValorEmissao = dValorEmissao

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassComissaoNF
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

