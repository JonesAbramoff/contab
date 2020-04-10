VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColCRFATConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassCRFATConfig"
Attribute VB_Ext_KEY = "Member0" ,"ClassCRFATConfig"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Public Function Add(sCodigo As String, iFilialEmpresa As Integer, sDescricao As String, iTipo As Integer, sConteudo As String, Optional sKey As String) As ClassCRFATConfig
    
    'create a new object
    Dim objNewMember As ClassCRFATConfig
    Set objNewMember = New ClassCRFATConfig

    'set the properties passed into the method
    objNewMember.sCodigo = sCodigo
    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.sDescricao = sDescricao
    objNewMember.iTipo = iTipo
    objNewMember.sConteudo = sConteudo

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassCRFATConfig
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

