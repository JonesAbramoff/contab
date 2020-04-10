VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColInfoComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassInfoComissao"
Attribute VB_Ext_KEY = "Member0" ,"ClassInfoComissao"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Public Function Add(lNumIntCom As Long, iTipoTitulo As Integer, lNumIntDoc As Long, iCodVendedor As Integer, dtDataBaixa As Date, dPercentual As Double, dValorBase As Double, dValor As Double, iStatus As Integer, iFilialEmpresa As Integer, sVendedorNomeRed As String, dtDataGeracao As Date, Optional sKey As String) As ClassInfoComissao

Dim objNewMember As ClassInfoComissao
Set objNewMember = New ClassInfoComissao
    
    objNewMember.lNumIntCom = lNumIntCom
    objNewMember.iTipoTitulo = iTipoTitulo
    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.iCodVendedor = iCodVendedor
    objNewMember.dtDataBaixa = dtDataBaixa
    objNewMember.dPercentual = dPercentual
    objNewMember.dValorBase = dValorBase
    objNewMember.dValor = dValor
    objNewMember.iStatus = iStatus
    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.sVendedorNomeRed = sVendedorNomeRed
    objNewMember.dtDataGeracao = dtDataGeracao

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassInfoComissao
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

