VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColInfoChequePag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassInfoChequePag"
Attribute VB_Ext_KEY = "Member0" ,"ClassInfoChequePag"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mCol As Collection

Public Function Add(iSeqCheque As Integer, sFavorecido As String, dValor As Double, lNumRealCheque As Long, iImprimir As Integer, iPortador As Integer, Optional sKey As String) As ClassInfoChequePag
    
Dim objNewMember As ClassInfoChequePag
Set objNewMember = New ClassInfoChequePag

    objNewMember.iSeqCheque = iSeqCheque
    objNewMember.sFavorecido = sFavorecido
    objNewMember.dValor = dValor
    objNewMember.lNumRealCheque = lNumRealCheque
    objNewMember.iImprimir = iImprimir
    objNewMember.iPortador = iPortador
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassInfoChequePag
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


