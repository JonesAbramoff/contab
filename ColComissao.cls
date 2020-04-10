VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassComissao"
Attribute VB_Ext_KEY = "Member0" ,"ClassComissao"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ClassComissao"
Option Explicit

Private mCol As Collection

Public Function AddObj(objNewMember As ClassComissao, Optional sKey As String) As ClassComissao

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set AddObj = objNewMember

End Function

Public Function Add(lNumIntCom As Long, iFilialEmpresa As Integer, iStatus As Integer, iTipoTitulo As Integer, lNumIntDoc As Long, iCodVendedor As Integer, dtDataBaixa As Date, dPercentual As Double, dValorBase As Double, dValor As Double, dtDataGeracao As Date, Optional sKey As String) As ClassComissao

Dim objNewMember As ClassComissao
Set objNewMember = New ClassComissao

    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.iStatus = iStatus
    objNewMember.lNumIntCom = lNumIntCom
    objNewMember.iTipoTitulo = iTipoTitulo
    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.iCodVendedor = iCodVendedor
    objNewMember.dtDataBaixa = dtDataBaixa
    objNewMember.dPercentual = dPercentual
    objNewMember.dValorBase = dValorBase
    objNewMember.dValor = dValor
    objNewMember.dtDataGeracao = dtDataGeracao

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassComissao
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

