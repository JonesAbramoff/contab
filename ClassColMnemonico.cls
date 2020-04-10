VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassColMnemonico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassMnemonico"
Attribute VB_Ext_KEY = "Member0" ,"ClassMnemonico"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(sModulo As String, iTransacao As Integer, sMnemonico As String, iTipo As Integer, iNumParam As Integer, iParam1 As Integer, iParam2 As Integer, iParam3 As Integer, sGrid As String, sMnemonicoCombo As String, sMnemonicoDesc As String, Optional sKey As String) As ClassMnemonico
    
    Dim objNewMember As ClassMnemonico
    Set objNewMember = New ClassMnemonico


    objNewMember.sModulo = sModulo
    objNewMember.iTransacao = iTransacao
    objNewMember.sMnemonico = sMnemonico
    objNewMember.iTipo = iTipo
    objNewMember.iNumParam = iNumParam
    objNewMember.iParam1 = iParam1
    objNewMember.iParam2 = iParam2
    objNewMember.iParam3 = iParam3
    objNewMember.sGrid = sGrid
    objNewMember.sMnemonicoCombo = sMnemonicoCombo
    objNewMember.sMnemonicoDesc = sMnemonicoDesc

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As ClassMnemonico
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

