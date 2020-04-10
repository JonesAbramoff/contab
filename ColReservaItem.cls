VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColReservaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassReservaItem"
Attribute VB_Ext_KEY = "Member0" ,"ClassReservaItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(iAlmoxarifado As Integer, dQuantidade As Double, sResponsavel As String, sAlmoxarifado As String, dtDataValidade As Date, lNumIntDoc As Long, Optional sKey As String) As ClassReservaItem

Dim objNewMember As ClassReservaItem
Set objNewMember = New ClassReservaItem

    objNewMember.iAlmoxarifado = iAlmoxarifado
    objNewMember.dQuantidade = dQuantidade
    objNewMember.sResponsavel = sResponsavel
    objNewMember.sAlmoxarifado = sAlmoxarifado
    objNewMember.dtDataValidade = dtDataValidade
    objNewMember.lNumIntDoc = lNumIntDoc

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassReservaItem
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

