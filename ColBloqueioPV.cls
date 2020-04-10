VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColBloqueioPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassBloqueioPV"
Attribute VB_Ext_KEY = "Member0" ,"ClassBloqueioPV"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(iFilialEmpresa As Integer, lPedidoDeVendas As Long, iSequencial As Integer, iTipoDeBloqueio As Integer, sCodUsuario As String, sResponsavel As String, dtData As Date, sCodUsuarioLib As String, sResponsavelLib As String, dtDataLib As Date, sObservacao As String, Optional sKey As String) As ClassBloqueioPV
    
Dim objNewMember As ClassBloqueioPV
Set objNewMember = New ClassBloqueioPV

    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.lPedidoDeVendas = lPedidoDeVendas
    objNewMember.iSequencial = iSequencial
    objNewMember.iTipoDeBloqueio = iTipoDeBloqueio
    objNewMember.sCodUsuario = sCodUsuario
    objNewMember.sResponsavel = sResponsavel
    objNewMember.dtData = dtData
    objNewMember.sCodUsuarioLib = sCodUsuarioLib
    objNewMember.sResponsavelLib = sResponsavelLib
    objNewMember.dtDataLib = dtDataLib
    objNewMember.sObservacao = sObservacao

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassBloqueioPV
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

Public Function AddObj(objNewMember As Object, Optional sKey As String) As Object

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set AddObj = objNewMember

End Function

