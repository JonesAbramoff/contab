VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member0" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member2" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member3" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member4" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member5" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member6" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member7" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member8" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member9" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member10" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member11" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member12" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member13" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member14" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member15" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member16" ,"ClassEndereco"
Attribute VB_Ext_KEY = "Member17" ,"ClassEndereco"
Option Explicit

Private mCol As Collection

Public Function Add(sEndereco As String, sBairro As String, sCidade As String, sSiglaEstado As String, iCodigoPais As Integer, sCEP As String, sTelefone1 As String, sTelefone2 As String, sEmail As String, sFax As String, sContato As String, lCodigo As Long, Optional sKey As String) As ClassEndereco

Dim objNewMember As ClassEndereco
Set objNewMember = New ClassEndereco

    objNewMember.sEndereco = sEndereco
    objNewMember.sBairro = sBairro
    objNewMember.sCidade = sCidade
    objNewMember.sSiglaEstado = sSiglaEstado
    objNewMember.iCodigoPais = iCodigoPais
    objNewMember.sCEP = sCEP
    objNewMember.sTelefone1 = sTelefone1
    objNewMember.sTelefone2 = sTelefone2
    objNewMember.sEmail = sEmail
    objNewMember.sFax = sFax
    objNewMember.sContato = sContato
    objNewMember.lCodigo = lCodigo

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassEndereco
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

Public Function AddObj(ByVal objEndereco As ClassEndereco, Optional sKey As String) As ClassEndereco

    If Len(sKey) = 0 Then
        mCol.Add objEndereco
    Else
        mCol.Add objEndereco, sKey
    End If

    Set AddObj = objEndereco

End Function
