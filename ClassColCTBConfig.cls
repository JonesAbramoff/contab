VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassColCTBConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Function Add(sCodigo As String, iFilialEmpresa As Integer, sDescricao As String, iTipo As Integer, sConteudo As String, Optional sKey As String) As ClassCTBConfig
    
    Dim objNewMember As ClassCTBConfig
    Set objNewMember = New ClassCTBConfig

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

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassCTBConfig
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

